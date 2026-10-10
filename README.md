# Naukri daily assistant (GitHub Actions)

Daily pipeline: scrape -> score vs master resume -> skip already-applied -> tailor resume -> apply
(answering screening questions) until the daily target is reached -> write `applied_jobs.md` / `.csv`.

## One-time setup
1. Create a **PRIVATE** GitHub repo and push this folder (it holds your resumes, phone and email).
   `session.json` is git-ignored: never commit it.
2. Locally: `pip install -r requirements.txt && playwright install chromium`
   then `python naukri_assistant.py login` (log in by hand), then `base64 -w0 session.json`
   (macOS: `base64 -i session.json`; Windows PowerShell: `[Convert]::ToBase64String([IO.File]::ReadAllBytes("session.json"))`).
3. Repo > Settings > Secrets > Actions: add `NAUKRI_SESSION_B64` (that output) and, optionally,
   `ANTHROPIC_API_KEY` (lets Claude answer free-text questions from your master resume).
4. Fill in `answers.json` (anything marked REPLACE_ME: current/expected CTC, relocation, ...).
   Jobs needing an unanswered question are skipped, not guessed.
5. Actions > naukri-daily > Run workflow with **dry_run ON** first. Check the log, then run it with dry_run off.

It then runs Mon-Sat at 10:00 IST (edit the cron in `.github/workflows/naukri-daily.yml`).

## Outputs (committed back to the repo each run)
- `applied_jobs.md` / `applied_jobs.csv` - applied jobs with the resume used, job link, match %.
  The .md also lists company-site jobs to apply to manually and jobs needing attention.
- `applied_log.json` - dedupe memory (a job is never retried once applied / company-site).
- `unanswered_questions.json` - questions the bot could not answer; add a rule to `answers.json`.
- `resumes/` - one tailored .docx per job.

## Things to know
- Session cookies expire. When the run fails with "Session expired", repeat step 2-3.
- Naukri may block datacenter IPs (GitHub runners) with a captcha. The run then stops and reports it.
- All CSS selectors are in `SEL` at the top of the script and are unverified against the live site;
  run a dry run, and fix selectors there if no jobs/buttons are found. Failure screenshots are uploaded as an artifact.
- "Apply on company site" jobs are not automated (every company site is different); they are listed for manual apply.
- Naukri applies with your *profile* resume, so the script uploads the tailored file to your profile
  before each apply and restores a master-resume .docx at the end.
- `daily_target` (10) is a target, not a guarantee: if fewer than 10 fresh jobs score >= `threshold`,
  it applies to fewer. Add keywords in `config.json`; the script never lowers the threshold silently.
- `max_clicks_per_run` is a safety stop in case confirmation detection breaks.
- Automated applying is against Naukri's terms and can get the account restricted. Use at your own risk.

## If the run stops with "Naukri blocked this browser"
GitHub-hosted runners use datacenter IPs, which Naukri's bot protection blocks. Run the same workflow on your own PC instead:
1. Repo > Settings > Actions > Runners > New self-hosted runner, and follow the steps on your PC (needs Git, Python 3.11+).
2. Repo > Settings > Secrets and variables > Actions > Variables: add `RUNNER_LABEL` = `self-hosted`.
3. The PC must be on at the scheduled time. A visible browser window opens during the run.
Alternative without GitHub: schedule `python naukri_assistant.py run` with Windows Task Scheduler or cron.
