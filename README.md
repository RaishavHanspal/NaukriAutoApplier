# Naukri assistant

Setup:
    pip install -r requirements.txt
    playwright install chromium

Edit master_resume.json (your real resume) and config.json (keywords, threshold, skills_vocab).

Run:
    python naukri_assistant.py login    # log in manually once; session.json is saved
    python naukri_assistant.py scrape   # collect listings + full descriptions
    python naukri_assistant.py score    # match % vs master resume
    python naukri_assistant.py tailor   # .docx per job scoring >= threshold
    python naukri_assistant.py apply    # opens each job; you confirm every Apply
