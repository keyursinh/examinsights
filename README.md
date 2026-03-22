# ExamInsights 🎓

**Question Bank Intelligence Platform** for B.E. Sem 6 exam papers.

## Features
- 📊 Dashboard with analytics — questions per subject, marks distribution
- 🔍 Manual Search with filters (subject, marks, exam type, year, section, similarity %)
- 📄 Check Paper — upload PDF/DOCX and check which questions exist in the bank
- ⊕ Insert Questions — extract from PDF/DOCX and auto-insert (duplicates skipped)
- ↓ Export Excel — download updated question bank

## Tech Stack
- **Backend:** Python, Flask
- **Data:** Pandas, OpenPyXL
- **File Parsing:** PyPDF2, python-docx
- **Server:** Gunicorn (Render) / Flask dev server (local)

## Run Locally

```bash
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000

## Deploy on Render

1. Push this repo to GitHub
2. Go to https://render.com → New → Web Service
3. Connect this repo
4. Build Command: `pip install -r requirements.txt`
5. Start Command: `gunicorn app:app`
6. Click Deploy

## Question Bank Format (Excel)

Each sheet = one subject. Required columns:
| Serial No | Subject | Exam | Section | Question No | Question | Marks |
