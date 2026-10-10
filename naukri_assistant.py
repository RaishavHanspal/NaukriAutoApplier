#!/usr/bin/env python3
"""
Naukri assistant: login (manual, session saved) -> scrape -> score vs master resume
-> tailored resume (.docx) -> semi-automatic apply (you confirm every job).

Usage:
  python naukri_assistant.py login
  python naukri_assistant.py scrape
  python naukri_assistant.py score
  python naukri_assistant.py tailor
  python naukri_assistant.py apply
  python naukri_assistant.py all        # scrape + score + tailor (no apply)
"""
import json
import random
import re
import sys
import time
from pathlib import Path

from docx import Document
from docx.shared import Pt
from playwright.sync_api import sync_playwright

BASE = Path(__file__).parent
SESSION = BASE / "session.json"
JOBS_FILE = BASE / "jobs.json"
OUT = BASE / "resumes"
CFG = json.loads((BASE / "config.json").read_text())
MASTER = json.loads((BASE / "master_resume.json").read_text())


def pause(a=3, b=7):
    time.sleep(random.uniform(a, b))


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


# ---------------------------------------------------------------- login
def login():
    """Opens a real browser. YOU log in (OTP/captcha included). Cookies saved to session.json.
    Your password never touches this script."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context()
        page = ctx.new_page()
        page.goto("https://www.naukri.com/nlogin/login")
        input("Log in in the browser window, then press Enter here... ")
        ctx.storage_state(path=str(SESSION))
        browser.close()
    print(f"Session saved to {SESSION}")


def new_context(p, headless=False):
    if not SESSION.exists():
        sys.exit("No session found. Run: python naukri_assistant.py login")
    browser = p.chromium.launch(headless=headless)
    ctx = browser.new_context(storage_state=str(SESSION))
    return browser, ctx


# ---------------------------------------------------------------- scrape
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


def scrape():
    jobs = json.loads(JOBS_FILE.read_text()) if JOBS_FILE.exists() else {}
    loc = CFG["location"]
    with sync_playwright() as p:
        browser, ctx = new_context(p)
        page = ctx.new_page()
        for kw in CFG["keywords"]:
            for pg in range(1, CFG["pages"] + 1):
                suffix = "" if pg == 1 else f"-{pg}"
                url = (f"https://www.naukri.com/{slug(kw)}-jobs-in-{slug(loc)}{suffix}"
                       f"?k={kw.replace(' ', '%20')}&l={loc}")
                print("Search:", url)
                page.goto(url, wait_until="domcontentloaded")
                try:
                    page.wait_for_selector(".srp-jobtuple-wrapper", timeout=15000)
                except Exception:
                    print("  no cards (captcha / layout change / end of results)")
                    break
                cards = page.eval_on_selector_all(".srp-jobtuple-wrapper", CARD_JS)
                for c in cards:
                    if c["url"] and c["url"] not in jobs:
                        jobs[c["url"]] = c
                pause()

        # Full descriptions
        for url, j in list(jobs.items())[: CFG["max_jobs"]]:
            if j.get("description"):
                continue
            page.goto(url, wait_until="domcontentloaded")
            try:
                page.wait_for_selector('[class*="dang-inner-html"]', timeout=10000)
                j["description"] = page.inner_text('[class*="dang-inner-html"]')
            except Exception:
                j["description"] = j.get("snippet", "")
            print("  JD:", j["title"], "-", j["company"])
            JOBS_FILE.write_text(json.dumps(jobs, indent=2))
            pause()
        browser.close()
    JOBS_FILE.write_text(json.dumps(jobs, indent=2))
    print(f"{len(jobs)} jobs saved to {JOBS_FILE}")


# ---------------------------------------------------------------- score
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
    area = any(a in (j.get("location", "") + jd) .lower() for a in CFG["area_terms"])
    return {"score": pct, "matched": matched, "missing": missing, "area_match": area}


def score():
    jobs = json.loads(JOBS_FILE.read_text())
    rows = []
    for url, j in jobs.items():
        j["match"] = score_job(j)
        rows.append(j)
    JOBS_FILE.write_text(json.dumps(jobs, indent=2))
    rows.sort(key=lambda r: r["match"]["score"], reverse=True)
    for r in rows:
        m = r["match"]
        flag = "OK " if m["score"] >= CFG["threshold"] else "-- "
        print(f"{flag}{m['score']:>3}%  {r['title'][:45]:45} {r['company'][:25]:25} "
              f"{'[area]' if m['area_match'] else ''}  missing: {', '.join(m['missing'][:5])}")


# ---------------------------------------------------------------- tailor
def build_resume(j, path):
    jd = j["description"].lower()
    jd_skills = set(j["match"]["matched"])
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Pt(36)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)

    doc.add_heading(MASTER["name"], 0)
    doc.add_paragraph(" | ".join(MASTER.get("contact", [])))
    doc.add_heading("Summary", 1)
    doc.add_paragraph(MASTER["summary"])

    # Skills: JD-relevant ones first. Only skills that are already in the master resume.
    skills = sorted(MASTER["skills"], key=lambda s: (s.lower() not in jd_skills, not has_term(s, jd)))
    doc.add_heading("Skills", 1)
    doc.add_paragraph(", ".join(skills))

    doc.add_heading("Experience", 1)
    for e in MASTER["experience"]:
        doc.add_paragraph().add_run(f"{e['title']} - {e['company']} ({e['dates']})").bold = True
        # Reorder bullets by relevance to the JD; wording is never changed.
        bullets = sorted(e["bullets"], key=lambda b: -sum(has_term(s, b.lower()) for s in jd_skills))
        for b in bullets:
            doc.add_paragraph(b, style="List Bullet")

    doc.add_heading("Education", 1)
    for ed in MASTER.get("education", []):
        doc.add_paragraph(ed)
    doc.save(path)


def tailor():
    OUT.mkdir(exist_ok=True)
    jobs = json.loads(JOBS_FILE.read_text())
    n = 0
    for url, j in jobs.items():
        m = j.get("match") or {}
        if m.get("score", 0) < CFG["threshold"]:
            continue
        fn = OUT / f"{slug(MASTER['name'])}_{slug(j['company'])}_{slug(j['title'])[:40]}.docx"
        build_resume(j, fn)
        j["resume_path"] = str(fn)
        n += 1
    JOBS_FILE.write_text(json.dumps(jobs, indent=2))
    print(f"{n} tailored resumes in {OUT}")


# ---------------------------------------------------------------- apply (human in the loop)
def apply():
    jobs = json.loads(JOBS_FILE.read_text())
    todo = [j for j in jobs.values()
            if j.get("resume_path") and not j.get("applied")
            and (j["match"]["area_match"] or not CFG["require_area_match"])]
    if not todo:
        return print("Nothing to apply to. Run score + tailor first.")
    with sync_playwright() as p:
        browser, ctx = new_context(p)
        page = ctx.new_page()
        for j in todo:
            print(f"\n{j['title']} @ {j['company']} | {j['location']} | match {j['match']['score']}%")
            print("Resume:", j["resume_path"])
            if input("Open this job? [y/N] ").lower() != "y":
                continue
            page.goto(j["url"], wait_until="domcontentloaded")
            print("Naukri applies with the resume on your PROFILE. To use the tailored file, upload it under "
                  "Profile > Resume first, then click Apply.")
            if input("Click Apply now? [y/N] ").lower() == "y":
                try:
                    page.click('#apply-button, button:has-text("Apply"):not(:has-text("company site"))',
                               timeout=8000)
                    input("Complete any questionnaire in the browser, then press Enter... ")
                    j["applied"] = True
                except Exception as ex:
                    print("Could not click Apply (company-site job or layout change):", ex)
            JOBS_FILE.write_text(json.dumps(jobs, indent=2))
        browser.close()


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    {"login": login, "scrape": scrape, "score": score, "tailor": tailor, "apply": apply,
     "all": lambda: (scrape(), score(), tailor())}.get(cmd, lambda: print(__doc__))()
