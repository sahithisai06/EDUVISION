# EduVision — OCR-Based Exam Sheet Reader

A Flask web app (mobile-first, installable as a home-screen app) that scans exam
answer sheets with the phone camera, pulls out **roll number, name, department
and marks** using OCR, and organizes everything into an Excel workbook —
sorted by roll number — so faculty stop transcribing marks by hand.

Built entirely with the stack already on your resume: **Python (Flask)** for
the backend, **HTML/CSS/JavaScript** for the frontend, plus `pytesseract`
(Tesseract OCR) and `openpyxl` for the Excel output — this is the same
tech description as your "EduVision – OCR Based Learning App" resume bullet,
just fully implemented.

## Features

1. **Sign-in page** — Gmail-style email/password login (session-based), with
   a "Continue with Google" button stubbed in for real OAuth later.
2. **Dashboard** — four tappable tiles: Scan Page, My Sheets, Export Excel,
   Account — plus a running count of sheets scanned.
3. **Scan Page** — opens the device camera directly (`getUserMedia`),
   captures a frame, runs OCR, and shows an editable review card so the
   faculty member can fix anything OCR got wrong before saving. Every saved
   record is flagged **High / Medium / Low confidence** so unclear scans get
   a second look — never silently trusted.
4. **My Sheets** — searchable history of every scan, plus a one-tap Excel
   download of the full master workbook.
5. **Account** — profile info and sign-out.

## Project structure

```
eduvision/
├── app.py                  # Flask app: routes, auth, OCR parsing, Excel writer
├── requirements.txt
├── templates/               # login, dashboard, scan, my_sheets, account
├── static/css/style.css     # visual identity (indigo + amber, exam-ledger theme)
├── static/js/                # (camera logic lives inline in scan.html)
└── data/                    # created at runtime: eduvision.db + marks_master.xlsx
```

## Running it locally

```bash
cd eduvision
pip install -r requirements.txt

# Tesseract OCR engine must be installed on the machine (not just the pip wrapper):
#   macOS:   brew install tesseract
#   Ubuntu:  sudo apt-get install tesseract-ocr
#   Windows: https://github.com/UB-Mannich/tesseract wiki (installer + add to PATH)

python app.py
```

Then open `http://localhost:5000` — on your phone, use your laptop's local
IP (`http://192.168.x.x:5000`) so the camera on the **phone itself** opens,
or deploy it (Render/Railway/PythonAnywhere all support Flask + Tesseract)
and open the deployed URL on your phone. Browsers only allow camera access
over `https://` or `localhost`, so for a real phone demo you'll want either
a deployed HTTPS URL or a tool like `ngrok` during development.

**Demo login:** `faculty@cmrcet.ac.in` / `eduvision123` (seeded automatically
on first run) — or tap "Create one" to sign up your own account.

## How the OCR parsing works

`extract_fields()` in `app.py` runs Tesseract on the captured frame, then
uses labeled-field regex (`Roll No:`, `Name:`, `Department:`, `Marks:`) to
pull structured values out of the raw text, falling back to a bare
roll-number pattern (`23H51A66D9`-style) if no label is found. Fields it
can't find confidently are left blank, and the overall record is marked
Low/Medium confidence — the review screen surfaces this instead of hiding
it, so faculty know exactly which entries to double check. This directly
addresses the "flag unclear marks for review" requirement from your project
brief.

## Turning this into a real installable phone app

Right now it's a mobile-responsive web app. To make it feel fully native:
- Add a `manifest.json` + service worker (a few lines) to make it a PWA —
  installable to the home screen with an offline shell, no app-store needed.
- Or wrap it with **Capacitor** or **React Native WebView** if you want an
  actual `.apk`/`.ipa` for the Play Store / App Store.

## Deploying so you can demo it live (good for interviews)

Render.com's free tier or PythonAnywhere both run Flask + Tesseract with
minimal config — either works well for a resume-linked live demo.

## Language recommendation for your resume

You're already well covered for this exact build — Python, HTML/CSS/JS,
MySQL — no changes needed for EduVision itself. Two additions worth
considering to round out your resume generally, based on projects you already
have:

- **SQL** as its own bullet under Languages (separate from "MySQL" under
  Tools) — you're already writing it in this project and the cancer-detection
  one; recruiters scan for it as its own keyword.
- **TypeScript** — since you already list React, adding TypeScript signals
  you can build production-grade frontends, which is one of the most
  requested pairings in current job postings (React + TS).

Not mandatory, but worth it only if you actually use them on a project —
adding a keyword with nothing behind it tends to fall apart in interviews.
