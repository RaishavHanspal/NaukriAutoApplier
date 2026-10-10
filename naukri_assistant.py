#!/usr/bin/env python3
"""
Naukri assistant - one-shot daily pipeline.

  python naukri_assistant.py login     # local only: log in by hand, saves session.json
  python naukri_assistant.py run       # scrape -> score -> tailor -> apply (to daily target) -> reports
  python naukri_assistant.py report    # rebuild applied_jobs.md / .csv from applied_log.json
  python naukri_assistant.py scrape | score | tailor   # individual steps (offline-friendly)

Environment:
  NAUKRI_SESSION_B64   base64 of session.json (GitHub secret). Used when session.json is absent.
  ANTHROPIC_API_KEY    optional; lets Claude answer free-text screening questions from your master resume.
  DRY_RUN=1            do everything except click Apply.
  CI=true              headless browser.

Selectors on Naukri change often. All of them live in SEL below.
"""
import base64
import csv
import json
import os
import random
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

from docx import Document
from docx.shared import Pt

BASE = Path(__file__).parent
SESSION = BASE / "session.json"
JOBS_FILE = BASE / "jobs.json"
LOG_FILE = BASE / "applied_log.json"
UNANSWERED_FILE = BASE / "unanswered_questions.json"
REPORT_MD = BASE / "applied_jobs.md"
REPORT_CSV = BASE / "applied_jobs.csv"
OUT = BASE / "resumes"
DEBUG = BASE / "debug"

IST = timezone(timedelta(hours=5, minutes=30))
DRY = os.environ.get("DRY_RUN", "").lower() in ("1", "true", "yes")
HEADLESS = os.environ.get("CI", "").lower() in ("1", "true", "yes")


def load_json(path, default):
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else default


def save_json(path, data):
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


CFG = load_json(BASE / "config.json", {})
MASTER = load_json(BASE / "master_resume.json", {})
ANSWERS = load_json(BASE / "answers.json", {"rules": []})


def now_ist():
    return datetime.now(IST)


def pause(a=3.0, b=7.0):
    time.sleep(random.uniform(a, b))


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


# ------------------------------------------------------------------ selectors (verify in DevTools)
SEL = {
    "card": ".srp-jobtuple-wrapper",
    "jd": '[class*="dang-inner-html"]',
    "apply_btn": '#apply-button, button:has-text("Apply"):not(:has-text("company site")):not(:has-text("Applied"))',
    "company_btn": '#company-site-button, button:has-text("company site"), a:has-text("company site")',
    "already": '#already-applied, button:text-is("Applied"), :text-matches("^\\\\s*Already Applied\\\\s*$", "i")',
    "success": "text=/successfully applied|application (has been )?(sent|submitted)|you have applied/i",
    "chat_root": '[class*="chatbot_DrawerContentWrapper"], [class*="chatbot_Drawer"]',
    "bot_msg": '[class*="botMsg"]',
    "chip": '[class*="chatbot_Chip"], [class*="ssrc__label"], [class*="chatbot_ChipContainer"] label',
    "text_input": '[class*="chatbot_InputContainer"] [contenteditable="true"], '
                  '[class*="chatbot_InputContainer"] input, [class*="chatbot"] textarea',
    "send": '[class*="sendMsg"], button:has-text("Save"), button:has-text("Submit")',
    "resume_input": 'input[type="file"]',
}

CARD_JS = """
els => els.map(e => {
  const q = s => e.querySelector(s);
  const t = q('a.title');
  return {
    title: t ? t.innerText.trim() : '',
    url: t ? t.href.split('?')[0] : '',
    company: q('.comp-name') ? q('.comp-name').innerText.trim() : '',
    experience: q('.exp-wrap') ? q('.exp-wrap').innerText.trim() : '',
    location: q('.loc-wrap') ? q('.loc-wrap').innerText.trim() : '',
    snippet: q('.job-desc') ? q('.job-desc').innerText.trim() : '',
  };
})
"""


class Blocked(Exception):
    pass


# ------------------------------------------------------------------ browser / session
def ensure_session():
    b64 = os.environ.get("NAUKRI_SESSION_B64", "").strip()
    if b64 and not SESSION.exists():
        SESSION.write_bytes(base64.b64decode(b64))
    if not SESSION.exists():
        sys.exit("No session. Run `python naukri_assistant.py login` locally, "
                 "or set the NAUKRI_SESSION_B64 secret.")


def new_context(p):
    ensure_session()
    browser = p.chromium.launch(headless=HEADLESS, args=["--disable-blink-features=AutomationControlled"])
    ctx = browser.new_context(
        storage_state=str(SESSION),
        viewport={"width": 1366, "height": 850},
        locale="en-IN",
        timezone_id="Asia/Kolkata",
        user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    )
    return browser, ctx


def goto_checked(page, url):
    page.goto(url, wait_until="domcontentloaded")
    title = (page.title() or "").lower()
    if "access denied" in title or "captcha" in page.url.lower():
        screenshot(page, "blocked")
        raise Blocked(f"Naukri blocked this browser (bot detection) at {page.url} | page title: {page.title()!r}")
    if "nlogin" in page.url:
        raise Blocked("Session expired - run `login` again and refresh the NAUKRI_SESSION_B64 secret")


def login():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context()
        page = ctx.new_page()
        page.goto("https://www.naukri.com/nlogin/login")
        input("Log in in the browser (OTP/captcha included), then press Enter here... ")
        ctx.storage_state(path=str(SESSION))
        browser.close()
    print(f"Saved {SESSION}. For GitHub Actions run:  base64 -w0 session.json  and store it as NAUKRI_SESSION_B64")


# ------------------------------------------------------------------ applied log (dedupe + report data)
DONE = {"applied", "already_applied", "company_site", "unconfirmed"}


def load_log():
    return load_json(LOG_FILE, [])


def migrate_log():
    """First run: seed the log from the old `applied` flags in jobs.json."""
    if LOG_FILE.exists():
        return
    jobs = load_json(JOBS_FILE, {})
    log = []
    for url, j in jobs.items():
        if j.get("applied"):
            rp = j.get("resume_path", "")
            rp = ("resumes/" + re.split(r"[\\/]", rp)[-1]) if rp else ""
            log.append({"date": "before-automation", "url": url, "title": j["title"], "company": j["company"],
                        "score": j.get("match", {}).get("score"), "resume": rp, "status": "applied",
                        "detail": "migrated from jobs.json", "attempts": 1})
    save_json(LOG_FILE, log)
    print(f"Migrated {len(log)} previously applied jobs into {LOG_FILE.name}")


def log_result(j, status, detail=""):
    log = load_log()
    prev = [e for e in log if e["url"] == j["url"]]
    attempts = (prev[-1].get("attempts", 1) + 1) if prev else 1
    log = [e for e in log if e["url"] != j["url"]]
    log.append({"date": now_ist().strftime("%Y-%m-%d %H:%M"), "url": j["url"], "title": j["title"],
                "company": j["company"], "location": j.get("location", ""),
                "score": j.get("match", {}).get("score"),
                "resume": j.get("resume_path", ""), "status": status, "detail": detail, "attempts": attempts})
    save_json(LOG_FILE, log)


def blocked_urls():
    """Jobs we must not try again."""
    out = set()
    for e in load_log():
        if e["status"] in DONE or e.get("attempts", 1) >= CFG.get("max_attempts_per_job", 2):
            out.add(e["url"])
    return out


def applied_today():
    d = now_ist().strftime("%Y-%m-%d")
    return sum(1 for e in load_log() if e["status"] == "applied" and e["date"].startswith(d))


# ------------------------------------------------------------------ scrape
def scrape(page, pages):
    jobs = load_json(JOBS_FILE, {})
    loc = CFG["location"]
    for kw in CFG["keywords"]:
        for pg in range(1, pages + 1):
            suffix = "" if pg == 1 else f"-{pg}"
            url = (f"https://www.naukri.com/{slug(kw)}-jobs-in-{slug(loc)}{suffix}"
                   f"?k={quote(kw)}&l={quote(loc)}")
            print("Search:", url)
            goto_checked(page, url)
            try:
                page.wait_for_selector(SEL["card"], timeout=15000)
            except Exception:
                print("  no cards (end of results or layout change)")
                break
            for c in page.eval_on_selector_all(SEL["card"], CARD_JS):
                if c["url"] and c["url"] not in jobs:
                    jobs[c["url"]] = c
            pause()
    skip = blocked_urls()
    pending = [j for j in jobs.values() if not j.get("description") and j["url"] not in skip]
    for j in pending[: CFG.get("max_new_descriptions", 40)]:
        goto_checked(page, j["url"])
        try:
            page.wait_for_selector(SEL["jd"], timeout=10000)
            j["description"] = page.inner_text(SEL["jd"])
        except Exception:
            j["description"] = j.get("snippet", "")
        print("  JD:", j["title"], "-", j["company"])
        save_json(JOBS_FILE, jobs)
        pause()
    save_json(JOBS_FILE, jobs)
    print(f"{len(jobs)} jobs known")


# ------------------------------------------------------------------ score
def master_text():
    parts = [MASTER.get("summary", ""), " ".join(MASTER.get("skills", []))]
    for e in MASTER["experience"]:
        parts += [e["title"]] + e["bullets"]
    return " ".join(parts).lower()


def has_term(term, text):
    return re.search(r"(?<![a-z0-9+#])" + re.escape(term.lower()) + r"(?![a-z0-9+#])", text) is not None


def score_job(j):
    jd = (j.get("title", "") + " " + j.get("description", "")).lower()
    jd_skills = [s for s in CFG["skills_vocab"] if has_term(s, jd)]
    mtxt = master_text()
    matched = [s for s in jd_skills if has_term(s, mtxt)]
    missing = [s for s in jd_skills if s not in matched]
    pct = round(100 * len(matched) / len(jd_skills)) if len(jd_skills) >= CFG["min_jd_skills"] else 0
    area = any(a in (j.get("location", "") + jd).lower() for a in CFG["area_terms"])
    return {"score": pct, "matched": matched, "missing": missing, "area_match": area}


def min_exp(j):
    m = re.match(r"\s*(\d+)\s*-\s*(\d+)", j.get("experience", ""))
    return int(m.group(1)) if m else 0


def score_all():
    jobs = load_json(JOBS_FILE, {})
    for j in jobs.values():
        j["match"] = score_job(j)
    save_json(JOBS_FILE, jobs)
    ranked = sorted(jobs.values(), key=lambda r: r["match"]["score"], reverse=True)
    for r in ranked[:25]:
        print(f"{r['match']['score']:>3}%  {r['title'][:40]:40} {r['company'][:28]:28} {r.get('experience','')}")


def candidates():
    jobs = load_json(JOBS_FILE, {})
    skip = blocked_urls()
    out = []
    for j in jobs.values():
        m = j.get("match") or {}
        if j["url"] in skip or m.get("score", 0) < CFG["threshold"]:
            continue
        if min_exp(j) > CFG.get("max_min_experience", 3):
            continue
        if CFG.get("require_area_match") and not m.get("area_match"):
            continue
        out.append(j)
    out.sort(key=lambda j: (not j["match"]["area_match"], -j["match"]["score"]))
    return out


# ------------------------------------------------------------------ tailor
def build_resume(j, path):
    jd = (j or {}).get("description", "").lower()
    jd_skills = set((j or {}).get("match", {}).get("matched", []))
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Pt(36)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)
    doc.add_heading(MASTER["name"], 0)
    doc.add_paragraph(" | ".join(MASTER.get("contact", [])))
    doc.add_heading("Summary", 1)
    doc.add_paragraph(MASTER["summary"])
    skills = sorted(MASTER["skills"], key=lambda s: (s.lower() not in jd_skills, not has_term(s, jd)))
    doc.add_heading("Skills", 1)
    doc.add_paragraph(", ".join(skills))
    doc.add_heading("Experience", 1)
    for e in MASTER["experience"]:
        doc.add_paragraph().add_run(f"{e['title']} - {e['company']} ({e['dates']})").bold = True
        bullets = sorted(e["bullets"], key=lambda b: -sum(has_term(s, b.lower()) for s in jd_skills))
        for b in bullets:
            doc.add_paragraph(b, style="List Bullet")
    doc.add_heading("Education", 1)
    for ed in MASTER.get("education", []):
        doc.add_paragraph(ed)
    doc.save(path)


def tailor_for(j):
    OUT.mkdir(exist_ok=True)
    fn = OUT / f"{slug(MASTER['name'])}_{slug(j['company'])}_{slug(j['title'])[:40]}.docx"
    build_resume(j, fn)
    j["resume_path"] = f"resumes/{fn.name}"
    return fn


def tailor_all():
    jobs = load_json(JOBS_FILE, {})
    n = 0
    for j in jobs.values():
        if j.get("match", {}).get("score", 0) >= CFG["threshold"]:
            tailor_for(j)
            n += 1
    save_json(JOBS_FILE, jobs)
    print(f"{n} tailored resumes in {OUT}")


def upload_profile_resume(page, path):
    """Naukri applies with the resume on your profile, so swap it before applying."""
    goto_checked(page, "https://www.naukri.com/mnjuser/profile")
    pause(2, 4)
    page.locator(SEL["resume_input"]).first.set_input_files(str(path))
    pause(4, 6)


# ------------------------------------------------------------------ questionnaire answering
def usable(answer):
    if answer is None:
        return False
    items = answer if isinstance(answer, list) else [answer]
    return all(a and "REPLACE_ME" not in str(a) for a in items)


def match_option(candidates_, options):
    cands = candidates_ if isinstance(candidates_, list) else [candidates_]
    low = {o.lower().strip(): o for o in options}
    for c in cands:
        c = str(c).lower().strip()
        if c in low:
            return low[c]
    for c in cands:
        c = str(c).lower().strip()
        for k, o in low.items():
            if c and (c in k or k in c):
                return o
    return None


def llm_answer(question, options):
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key or not CFG.get("use_llm_for_questions", True):
        return None
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        facts = [r for r in ANSWERS.get("rules", []) if usable(r.get("answer"))]
        system = (
            "You answer job-application screening questions on behalf of the candidate, using ONLY the "
            "facts in the master resume and profile facts provided. Never invent facts, numbers, salary, "
            "dates, employers or skills. You may infer years of use of a tool only from the dated roles in "
            "the resume. If the answer is not clearly supported, reply with exactly UNKNOWN. "
            "If options are given, reply with exactly one option copied verbatim. For free text reply in "
            "1-3 plain first-person sentences, no markdown.")
        user = (f"MASTER RESUME:\n{json.dumps(MASTER)}\n\nPROFILE FACTS:\n{json.dumps(facts)}\n\n"
                f"QUESTION: {question}\nOPTIONS: {options if options else 'none (free text)'}")
        msg = client.messages.create(model=CFG.get("llm_model", "claude-haiku-5-5"), max_tokens=300,
                                     system=system, messages=[{"role": "user", "content": user}])
        text = msg.content[0].text.strip()
        if text.upper().startswith("UNKNOWN"):
            return None
        return match_option(text, options) if options else text
    except Exception as ex:
        print("  LLM answer failed:", ex)
        return None


def answer_question(question, options):
    """Returns an answer string or None (None = we don't know, so a human must answer)."""
    ql = question.lower()
    for rule in ANSWERS.get("rules", []):
        if re.search(rule["pattern"], ql) and usable(rule.get("answer")):
            a = rule["answer"]
            if options:
                m = match_option(a, options)
                if m:
                    return m
            else:
                return a[0] if isinstance(a, list) else a
    return llm_answer(question, options)


def record_unanswered(j, question, options):
    data = load_json(UNANSWERED_FILE, [])
    if not any(d["question"] == question for d in data):
        data.append({"question": question, "options": options, "example_job": j["url"],
                     "how_to_fix": "add a rule for this to answers.json"})
        save_json(UNANSWERED_FILE, data)


# ------------------------------------------------------------------ apply
def confirm_applied(page, url):
    if page.locator(SEL["success"]).count():
        return True
    goto_checked(page, url)
    pause(1.5, 3)
    return page.locator(SEL["already"]).count() > 0


def run_questionnaire(page, j):
    last_q, same = None, 0
    for _ in range(CFG.get("max_questions", 15)):
        if page.locator(SEL["success"]).count():
            return "applied", ""
        root = page.locator(SEL["chat_root"]).first
        if not page.locator(SEL["chat_root"]).count():
            break
        msgs = root.locator(SEL["bot_msg"])
        n = msgs.count()
        question = msgs.nth(n - 1).inner_text().strip() if n else ""
        options = [c.inner_text().strip() for c in root.locator(SEL["chip"]).all()]
        options = [o for o in options if o]
        if not question:
            pause(1, 2)
            continue
        same = same + 1 if question == last_q else 0
        last_q = question
        if same >= 2:
            record_unanswered(j, question, options)
            return "needs_manual", f"stuck on: {question}"
        ans = answer_question(question, options)
        if ans is None:
            record_unanswered(j, question, options)
            return "needs_manual", f"no answer for: {question}"
        print(f"  Q: {question[:80]} -> {ans[:60]}")
        try:
            if options:
                for part in [a.strip() for a in ans.split("||")]:
                    root.locator(SEL["chip"]).filter(
                        has_text=re.compile(rf"^\s*{re.escape(part)}\s*$", re.I)).first.click()
                if root.locator(SEL["send"]).count():
                    root.locator(SEL["send"]).first.click()
            else:
                box = root.locator(SEL["text_input"]).first
                box.click()
                box.fill(ans)
                if root.locator(SEL["send"]).count():
                    root.locator(SEL["send"]).first.click()
                else:
                    page.keyboard.press("Enter")
        except Exception as ex:
            return "needs_manual", f"could not answer '{question}': {ex}"
        pause(1.5, 3)
    return ("applied", "") if page.locator(SEL["success"]).count() else ("unconfirmed", "")


def apply_to(page, j):
    goto_checked(page, j["url"])
    pause(2, 4)
    if page.locator(SEL["already"]).count():
        return "already_applied", ""
    has_apply = page.locator(SEL["apply_btn"]).count() > 0
    if not has_apply:
        if page.locator(SEL["company_btn"]).count():
            return "company_site", "Apply on company site - open the link and apply manually"
        return "failed", "no Apply button found (layout change or expired listing)"
    if DRY:
        return "dry_run", "would apply"
    page.locator(SEL["apply_btn"]).first.click()
    pause(2.5, 4.5)
    status, detail = run_questionnaire(page, j)
    if status in ("applied", "unconfirmed"):
        status = "applied" if confirm_applied(page, j["url"]) else "unconfirmed"
    return status, detail


def screenshot(page, name):
    try:
        DEBUG.mkdir(exist_ok=True)
        page.screenshot(path=str(DEBUG / f"{slug(name)[:60]}.png"))
    except Exception:
        pass


def apply_loop(page, target):
    done_today = applied_today()
    new_applied, clicks = 0, 0
    master_docx = OUT / "_master_resume.docx"
    swapped = False
    for j in candidates():
        if done_today + new_applied >= target:
            break
        if clicks >= CFG.get("max_clicks_per_run", 15):
            print("Safety cap reached for this run (max_clicks_per_run)")
            break
        print(f"\n[{j['match']['score']}%] {j['title']} @ {j['company']} ({j.get('experience','')})")
        resume = tailor_for(j)
        try:
            if CFG.get("upload_tailored_resume", True) and not DRY:
                upload_profile_resume(page, resume)
                swapped = True
            status, detail = apply_to(page, j)
        except Blocked:
            raise
        except Exception as ex:
            status, detail = "failed", str(ex)[:200]
        print("  ->", status, detail)
        if status == "dry_run":
            continue
        clicks += 1 if status in ("applied", "unconfirmed", "needs_manual") else 0
        if status in ("failed", "needs_manual", "unconfirmed"):
            screenshot(page, f"{j['company']}-{j['title']}")
        log_result(j, status, detail)
        if status == "applied":
            new_applied += 1
        pause(*CFG.get("apply_delay", [10, 25]))
    if swapped and CFG.get("restore_master_resume", True):
        OUT.mkdir(exist_ok=True)
        build_resume(None, master_docx)
        try:
            upload_profile_resume(page, master_docx)
            print("Profile resume restored to master")
        except Exception as ex:
            print("WARNING: could not restore master resume on profile:", ex)
    return new_applied


# ------------------------------------------------------------------ reports
def write_reports():
    log = load_log()
    applied = sorted([e for e in log if e["status"] == "applied"], key=lambda e: e["date"], reverse=True)

    def link(e):
        return f"[{Path(e['resume']).name}]({quote(e['resume'])})" if e.get("resume") else "-"

    lines = ["# Applied jobs", "", f"Total applied: **{len(applied)}**  |  last updated "
             f"{now_ist().strftime('%Y-%m-%d %H:%M IST')}", "",
             "| Date | Role | Company | Match | Resume used | Job |", "|---|---|---|---|---|---|"]
    for e in applied:
        lines.append(f"| {e['date']} | {e['title']} | {e['company']} | {e.get('score','')}% | "
                     f"{link(e)} | [open]({e['url']}) |")
    comp = [e for e in log if e["status"] == "company_site"]
    if comp:
        lines += ["", "## Apply manually: company-site jobs", "", "| Role | Company | Match | Resume ready | Job |",
                  "|---|---|---|---|---|"]
        lines += [f"| {e['title']} | {e['company']} | {e.get('score','')}% | {link(e)} | [open]({e['url']}) |"
                  for e in comp]
    man = [e for e in log if e["status"] in ("needs_manual", "unconfirmed", "failed")]
    if man:
        lines += ["", "## Needs your attention", "", "| Role | Company | Status | Detail | Job |", "|---|---|---|---|---|"]
        lines += [f"| {e['title']} | {e['company']} | {e['status']} | {e.get('detail','')[:100]} | "
                  f"[open]({e['url']}) |" for e in man]
    REPORT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    with REPORT_CSV.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["date", "title", "company", "location", "match_pct", "status", "resume", "url", "detail"])
        for e in sorted(log, key=lambda e: e["date"], reverse=True):
            w.writerow([e["date"], e["title"], e["company"], e.get("location", ""), e.get("score", ""),
                        e["status"], e.get("resume", ""), e["url"], e.get("detail", "")])
    print(f"Reports written: {REPORT_MD.name}, {REPORT_CSV.name}")


# ------------------------------------------------------------------ one-go pipeline
def run():
    from playwright.sync_api import sync_playwright
    migrate_log()
    target = CFG.get("daily_target", 10)
    if applied_today() >= target:
        print(f"Already applied to {applied_today()} jobs today (target {target}).")
        write_reports()
        return
    with sync_playwright() as p:
        browser, ctx = new_context(p)
        page = ctx.new_page()
        try:
            for attempt, pages in enumerate([CFG.get("pages", 2), CFG.get("pages_if_short", 5)]):
                scrape(page, pages)
                score_all()
                have = len(candidates())
                need = target - applied_today()
                print(f"Eligible (>= {CFG['threshold']}% match, not yet applied): {have}; still need {need}")
                if have >= need or attempt == 1:
                    break
                print("Not enough fresh matches - scraping deeper pages...")
            n = apply_loop(page, target)
            print(f"\nApplied this run: {n} | today total: {applied_today()} / {target}")
            if applied_today() < target:
                print("Below target: not enough eligible matches. Add keywords in config.json or lower "
                      "`threshold` yourself - the script will not do that silently.")
        except Blocked as ex:
            print("STOPPED:", ex)
            sys.exit(2)
        finally:
            write_reports()
            browser.close()


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "login":
        login()
    elif cmd == "run":
        run()
    elif cmd == "report":
        migrate_log()
        write_reports()
    elif cmd == "score":
        score_all()
    elif cmd == "tailor":
        tailor_all()
    elif cmd == "scrape":
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b, c = new_context(p)
            scrape(c.new_page(), CFG.get("pages", 2))
            b.close()
    else:
        print(__doc__)
