from flask import Flask, request, jsonify, render_template_string, send_file
import pandas as pd
import re
import os
from PyPDF2 import PdfReader
from docx import Document

app = Flask(__name__)
EXCEL_FILE = "question_bank.xlsx"

# ─────────────────────────────────────────────
# CORE UTILITIES
# ─────────────────────────────────────────────

def similarity(q1, q2):
    q1 = re.sub(r'[^a-zA-Z0-9 ]', '', q1.lower())
    q2 = re.sub(r'[^a-zA-Z0-9 ]', '', q2.lower())
    w1 = set(q1.split())
    w2 = set(q2.split())
    if not w1:
        return 0
    return len(w1 & w2) / len(w1)

def extract_exam_type(text):
    t = text.lower()
    if "summer" in t: return "Summer"
    if "winter" in t: return "Winter"
    return "Unknown"

def extract_semester(text):
    m = re.search(r'Sem[-\s]?(\d+)', text, re.IGNORECASE)
    return f"Sem {m.group(1)}" if m else ""

def extract_year(text):
    m = re.search(r'20\d{2}', text)
    return m.group(0) if m else ""

def load_db():
    sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    data = []
    for subject, df in sheets.items():
        df = df.fillna("")
        for _, row in df.iterrows():
            exam_text = str(row.get("Exam", ""))
            data.append({
                "serial":    str(row.get("Serial No", "")),
                "question":  str(row.get("Question", "")),
                "subject":   subject,
                "exam":      exam_text,
                "exam_type": extract_exam_type(exam_text),
                "semester":  extract_semester(exam_text),
                "year":      extract_year(exam_text),
                "section":   str(row.get("Section", "")),
                "q_no":      str(row.get("Question No", "")),
                "marks":     str(row.get("Marks", "")),
            })
    return data

def extract_pdf(file):
    reader = PdfReader(file)
    text = "".join(p.extract_text() or "" for p in reader.pages)
    return re.findall(r'(?:Q\.?\s*\d+|\d+\.)\s*(.*)', text)

def extract_docx(file):
    doc = Document(file)
    qs = []
    for p in doc.paragraphs:
        m = re.match(r'(?:Q\.?\s*\d+|\d+\.)\s*(.*)', p.text)
        if m:
            qs.append(m.group(1))
    return qs

# ─────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ExamInsight</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg:       #f9fafb;
  --white:    #ffffff;
  --ink:      #111827;
  --ink2:     #374151;
  --muted:    #9ca3af;
  --border:   #e5e7eb;
  --sidebar:  #111827;
  --accent:   #2563eb;
  --accent-h: #1d4ed8;
  --green:    #059669;
  --red:      #dc2626;
  --amber:    #d97706;
  --radius:   10px;
  --sidebar-w: 240px;
}

body {
  font-family: 'Plus Jakarta Sans', sans-serif;
  background: var(--bg);
  color: var(--ink);
  font-size: 14px;
  line-height: 1.5;
}

/* ── SIDEBAR ── */
.sidebar {
  position: fixed;
  top: 0; left: 0; bottom: 0;
  width: var(--sidebar-w);
  background: var(--sidebar);
  display: flex;
  flex-direction: column;
  z-index: 200;
  overflow: hidden;
}

.logo {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 20px 16px;
  border-bottom: 1px solid rgba(255,255,255,0.06);
  min-width: 0;
  flex-shrink: 0;
}

.logo-icon {
  width: 34px;
  height: 34px;
  flex-shrink: 0;
  background: var(--accent);
  border-radius: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.logo-icon svg {
  width: 18px;
  height: 18px;
  fill: none;
  stroke: #fff;
  stroke-width: 2;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.logo-text {
  min-width: 0;
  overflow: hidden;
}

.logo-name {
  font-size: 15px;
  font-weight: 700;
  color: #fff;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  line-height: 1.2;
}

.logo-tagline {
  font-size: 10px;
  color: rgba(255,255,255,0.35);
  text-transform: uppercase;
  letter-spacing: 0.8px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.nav {
  flex: 1;
  padding: 12px 10px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  overflow-y: auto;
}

.nav-section {
  font-size: 10px;
  font-weight: 600;
  letter-spacing: 1px;
  text-transform: uppercase;
  color: rgba(255,255,255,0.25);
  padding: 10px 8px 4px;
}

.nav-btn {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 9px 10px;
  border-radius: 8px;
  border: none;
  background: none;
  color: rgba(255,255,255,0.55);
  font-family: inherit;
  font-size: 13.5px;
  font-weight: 500;
  cursor: pointer;
  width: 100%;
  text-align: left;
  transition: all 0.15s;
  white-space: nowrap;
}

.nav-btn:hover {
  background: rgba(255,255,255,0.07);
  color: rgba(255,255,255,0.9);
}

.nav-btn.active {
  background: rgba(37,99,235,0.25);
  color: #fff;
}

.nav-btn svg {
  width: 16px; height: 16px;
  flex-shrink: 0;
  stroke: currentColor;
  fill: none;
  stroke-width: 1.8;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.sidebar-foot {
  padding: 12px 10px;
  border-top: 1px solid rgba(255,255,255,0.06);
  flex-shrink: 0;
}

.sidebar-stat {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 5px 8px;
}

.sidebar-stat-label {
  font-size: 11px;
  color: rgba(255,255,255,0.3);
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.sidebar-stat-val {
  font-size: 14px;
  font-weight: 700;
  color: #fff;
}

/* ── MOBILE TOGGLE ── */
.menu-toggle {
  display: none;
  position: fixed;
  top: 12px; left: 12px;
  z-index: 300;
  width: 38px; height: 38px;
  background: var(--sidebar);
  border: none;
  border-radius: 8px;
  align-items: center;
  justify-content: center;
  cursor: pointer;
}

.menu-toggle svg {
  width: 18px; height: 18px;
  stroke: #fff;
  fill: none;
  stroke-width: 2;
  stroke-linecap: round;
}

.overlay {
  display: none;
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.5);
  z-index: 150;
}

/* ── MAIN ── */
.main {
  margin-left: var(--sidebar-w);
  min-height: 100vh;
}

.topbar {
  position: sticky;
  top: 0;
  background: rgba(249,250,251,0.92);
  backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--border);
  height: 56px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 24px;
  z-index: 100;
}

.topbar-title {
  font-size: 15px;
  font-weight: 700;
  color: var(--ink);
}

.topbar-right {
  display: flex;
  align-items: center;
  gap: 8px;
}

.chip {
  display: inline-flex;
  align-items: center;
  padding: 4px 10px;
  border-radius: 100px;
  font-size: 12px;
  font-weight: 600;
}

.chip-blue  { background: #eff6ff; color: var(--accent); }
.chip-green { background: #ecfdf5; color: var(--green); }

.content { padding: 24px; max-width: 1200px; }

/* ── SECTIONS ── */
.section { display: none; animation: fadeIn 0.2s ease; }
.section.active { display: block; }

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(8px); }
  to   { opacity: 1; transform: translateY(0); }
}

/* ── PAGE HEADER ── */
.page-header { margin-bottom: 20px; }
.page-header h1 { font-size: 22px; font-weight: 700; color: var(--ink); margin-bottom: 3px; }
.page-header p  { font-size: 13px; color: var(--muted); }

/* ── STATS GRID ── */
.stats-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 14px;
  margin-bottom: 20px;
}

.stat-card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 18px 20px;
}

.stat-num {
  font-size: 32px;
  font-weight: 700;
  color: var(--ink);
  line-height: 1;
  margin-bottom: 4px;
}

.stat-lbl {
  font-size: 11px;
  font-weight: 600;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.6px;
}

/* ── CHARTS ── */
.charts-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px;
  margin-bottom: 20px;
}

.card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px;
}

.card-title {
  font-size: 13px;
  font-weight: 700;
  color: var(--ink);
  margin-bottom: 16px;
}

.bar-chart { display: flex; flex-direction: column; gap: 8px; }

.bar-row { display: flex; align-items: center; gap: 10px; }

.bar-lbl {
  font-size: 11px;
  color: var(--muted);
  width: 120px;
  flex-shrink: 0;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.bar-track {
  flex: 1;
  height: 6px;
  background: var(--bg);
  border-radius: 100px;
  overflow: hidden;
}

.bar-fill {
  height: 100%;
  border-radius: 100px;
  background: var(--accent);
  transition: width 0.5s ease;
}

.bar-val { font-size: 11px; font-weight: 600; color: var(--ink); width: 24px; text-align: right; flex-shrink: 0; }

.donut-wrap { display: flex; align-items: center; gap: 20px; }
.donut-legend { display: flex; flex-direction: column; gap: 8px; }
.legend-row { display: flex; align-items: center; gap: 8px; font-size: 12px; }
.legend-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.legend-name { color: var(--muted); flex: 1; }
.legend-val { font-weight: 600; color: var(--ink); }

/* ── SUBJECT CARDS ── */
.subjects-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 12px;
}

.subj-card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 16px;
  display: flex;
  align-items: center;
  gap: 12px;
  cursor: pointer;
  transition: all 0.15s;
}

.subj-card:hover {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(37,99,235,0.08);
}

.subj-icon {
  width: 38px; height: 38px;
  border-radius: 8px;
  background: #eff6ff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 18px;
  flex-shrink: 0;
}

.subj-name {
  font-size: 13px;
  font-weight: 600;
  color: var(--ink);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.subj-count { font-size: 12px; color: var(--muted); }

.subj-num {
  margin-left: auto;
  font-size: 18px;
  font-weight: 700;
  color: var(--accent);
  flex-shrink: 0;
}

/* ── SEARCH BOX ── */
.search-card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 20px;
  margin-bottom: 16px;
}

.field-row {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  align-items: flex-end;
}

.field {
  display: flex;
  flex-direction: column;
  gap: 5px;
}

.field label {
  font-size: 10px;
  font-weight: 700;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.7px;
}

.inp {
  height: 38px;
  border: 1.5px solid var(--border);
  border-radius: 8px;
  padding: 0 12px;
  font-family: inherit;
  font-size: 13px;
  color: var(--ink);
  background: var(--bg);
  outline: none;
  transition: border-color 0.15s;
}

.inp:focus { border-color: var(--accent); background: var(--white); }

.inp-lg { flex: 1; min-width: 200px; }
.inp-sm { width: 130px; }
.inp-xs { width: 100px; }

select.inp { cursor: pointer; }

.divider { height: 1px; background: var(--border); margin: 14px 0; }

.filter-row {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  align-items: flex-end;
}

/* ── BUTTONS ── */
.btn {
  height: 38px;
  padding: 0 18px;
  border-radius: 8px;
  border: none;
  font-family: inherit;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  transition: all 0.15s;
  white-space: nowrap;
}

.btn-primary { background: var(--accent); color: #fff; }
.btn-primary:hover { background: var(--accent-h); }

.btn-ghost {
  background: var(--bg);
  color: var(--ink2);
  border: 1.5px solid var(--border);
}

.btn-ghost:hover { border-color: var(--accent); color: var(--accent); }

/* ── RESULTS TABLE ── */
.results-card {
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  overflow: hidden;
}

.results-head {
  padding: 14px 20px;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
}

.results-title {
  font-size: 13px;
  font-weight: 700;
  color: var(--ink);
  display: flex;
  align-items: center;
  gap: 8px;
}

.count-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 22px;
  height: 22px;
  padding: 0 7px;
  background: var(--accent);
  color: #fff;
  border-radius: 100px;
  font-size: 11px;
  font-weight: 700;
}

.tbl-wrap { overflow-x: auto; }

table { width: 100%; border-collapse: collapse; font-size: 13px; }

thead th {
  background: var(--bg);
  padding: 10px 14px;
  text-align: left;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.7px;
  text-transform: uppercase;
  color: var(--muted);
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
}

tbody tr { border-bottom: 1px solid var(--border); transition: background 0.1s; }
tbody tr:last-child { border-bottom: none; }
tbody tr:hover { background: #f8faff; }

td { padding: 12px 14px; color: var(--ink2); vertical-align: middle; }

.td-q { max-width: 340px; line-height: 1.5; font-size: 13px; }

/* ── PILLS ── */
.pill {
  display: inline-flex;
  align-items: center;
  padding: 2px 9px;
  border-radius: 100px;
  font-size: 11px;
  font-weight: 600;
  white-space: nowrap;
}

.pill-blue   { background: #eff6ff; color: var(--accent); }
.pill-green  { background: #ecfdf5; color: var(--green); }
.pill-red    { background: #fef2f2; color: var(--red); }
.pill-amber  { background: #fffbeb; color: var(--amber); }
.pill-gray   { background: #f3f4f6; color: #6b7280; }
.pill-orange { background: #fff7ed; color: #ea580c; }

.marks-num {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 26px; height: 26px;
  border-radius: 50%;
  font-size: 12px;
  font-weight: 700;
  background: var(--bg);
  border: 1.5px solid var(--border);
  color: var(--ink);
}

.sim-wrap { display: flex; align-items: center; gap: 7px; }
.sim-bar { width: 48px; height: 4px; background: var(--bg); border-radius: 100px; overflow: hidden; }
.sim-fill { height: 100%; border-radius: 100px; }
.sim-txt { font-size: 11px; color: var(--muted); }

/* ── UPLOAD ── */
.upload-zone {
  border: 2px dashed var(--border);
  border-radius: var(--radius);
  padding: 36px;
  text-align: center;
  cursor: pointer;
  transition: all 0.15s;
  position: relative;
  background: var(--white);
}

.upload-zone:hover, .upload-zone.over { border-color: var(--accent); background: #f0f7ff; }

.upload-zone input {
  position: absolute; inset: 0;
  opacity: 0; cursor: pointer;
  width: 100%; height: 100%;
}

.upload-ico { font-size: 32px; margin-bottom: 10px; }
.upload-title { font-size: 14px; font-weight: 600; color: var(--ink); margin-bottom: 4px; }
.upload-sub { font-size: 12px; color: var(--muted); }
.upload-name { display: none; margin-top: 10px; font-size: 12px; font-weight: 600; color: var(--accent); background: #eff6ff; padding: 6px 12px; border-radius: 6px; }

/* ── EMPTY / LOADER ── */
.empty { text-align: center; padding: 40px; color: var(--muted); }
.empty-ico { font-size: 36px; margin-bottom: 10px; opacity: 0.35; }
.empty p { font-size: 13px; }

.loader { display: none; align-items: center; justify-content: center; padding: 40px; gap: 10px; flex-direction: column; }
.loader.on { display: flex; }
.spin { width: 28px; height: 28px; border: 2.5px solid var(--border); border-top-color: var(--accent); border-radius: 50%; animation: spin 0.7s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.spin-txt { font-size: 12px; color: var(--muted); }

/* ── NOTICE ── */
.notice {
  background: #eff6ff;
  border: 1px solid #bfdbfe;
  border-radius: 8px;
  padding: 12px 16px;
  font-size: 12.5px;
  color: var(--accent);
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

/* ── TOAST ── */
.toast-stack { position: fixed; bottom: 20px; right: 20px; z-index: 9999; display: flex; flex-direction: column; gap: 8px; pointer-events: none; }
.toast { padding: 12px 16px; border-radius: 10px; font-size: 13px; font-weight: 500; display: flex; align-items: center; gap: 8px; pointer-events: auto; box-shadow: 0 4px 20px rgba(0,0,0,0.12); animation: tin 0.25s ease; max-width: 300px; }
.toast-ok  { background: var(--ink); color: #fff; }
.toast-err { background: #fef2f2; color: var(--red); border: 1px solid #fecaca; }
.toast-inf { background: #eff6ff; color: var(--accent); border: 1px solid #bfdbfe; }
@keyframes tin { from { opacity: 0; transform: translateX(16px); } to { opacity: 1; transform: translateX(0); } }

mark { background: #fef9c3; color: #92400e; border-radius: 2px; padding: 0 1px; }

/* ── RESPONSIVE ── */
@media (max-width: 768px) {
  :root { --sidebar-w: 240px; }

  .menu-toggle { display: flex; }

  .sidebar {
    transform: translateX(-100%);
    transition: transform 0.25s ease;
  }

  .sidebar.open { transform: translateX(0); }
  .overlay.open { display: block; }

  .main { margin-left: 0; }

  .topbar { padding: 0 16px 0 58px; }

  .content { padding: 16px; }

  .stats-grid { grid-template-columns: repeat(2, 1fr); gap: 10px; }

  .charts-row { grid-template-columns: 1fr; }

  .field-row { flex-direction: column; }
  .inp-lg, .inp-sm, .inp-xs { width: 100%; }

  .filter-row { flex-direction: column; }
  .filter-row .field { width: 100%; }
  .filter-row .inp { width: 100%; }

  .btn { width: 100%; justify-content: center; }

  .subjects-grid { grid-template-columns: 1fr; }

  table { font-size: 12px; }
  td, thead th { padding: 10px 10px; }
  .td-q { max-width: 180px; }
}

@media (max-width: 480px) {
  .stats-grid { grid-template-columns: 1fr 1fr; }
  .stat-num { font-size: 24px; }
  .topbar-right .chip:last-child { display: none; }
}
</style>
</head>
<body>

<!-- Mobile toggle -->
<button class="menu-toggle" onclick="toggleSidebar()" aria-label="Menu">
  <svg viewBox="0 0 24 24"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
</button>

<div class="overlay" id="overlay" onclick="closeSidebar()"></div>

<!-- SIDEBAR -->
<aside class="sidebar" id="sidebar">
  <div class="logo">
    <div class="logo-icon">
      <svg viewBox="0 0 24 24">
        <path d="M12 2L2 7l10 5 10-5-10-5z"/>
        <path d="M2 17l10 5 10-5"/>
        <path d="M2 12l10 5 10-5"/>
      </svg>
    </div>
    <div class="logo-text">
      <div class="logo-name">ExamInsight</div>
      <div class="logo-tagline">Question Bank</div>
    </div>
  </div>

  <nav class="nav">
    <div class="nav-section">Main</div>
    <button class="nav-btn active" onclick="show('dashboard')" id="nav-dashboard">
      <svg viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
      Dashboard
    </button>
    <button class="nav-btn" onclick="show('search')" id="nav-search">
      <svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
      Manual Search
    </button>
    <button class="nav-btn" onclick="show('check')" id="nav-check">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><polyline points="9 15 11 17 15 13"/></svg>
      Check Paper
    </button>
    <div class="nav-section">Manage</div>
    <button class="nav-btn" onclick="show('insert')" id="nav-insert">
      <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
      Insert Questions
    </button>
    <button class="nav-btn" onclick="doDownload()">
      <svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Export Excel
    </button>
  </nav>

  <div class="sidebar-foot">
    <div class="sidebar-stat"><span class="sidebar-stat-label">Subjects</span><span class="sidebar-stat-val" id="ss-sub">—</span></div>
    <div class="sidebar-stat"><span class="sidebar-stat-label">Questions</span><span class="sidebar-stat-val" id="ss-q">—</span></div>
    <div class="sidebar-stat"><span class="sidebar-stat-label">Exams</span><span class="sidebar-stat-val" id="ss-e">—</span></div>
  </div>
</aside>

<!-- MAIN -->
<div class="main">
  <div class="topbar">
    <div class="topbar-title" id="page-title">Dashboard</div>
    <div class="topbar-right">
      <span class="chip chip-green" id="top-q">— Questions</span>
      <span class="chip chip-blue">B.E. Sem 6</span>
    </div>
  </div>

  <div class="content">

    <!-- DASHBOARD -->
    <div id="dashboard" class="section active">
      <div class="page-header">
        <h1>Overview</h1>
        <p>Question bank analytics across all subjects</p>
      </div>

      <div class="stats-grid">
        <div class="stat-card"><div class="stat-num" id="d-sub">—</div><div class="stat-lbl">Subjects</div></div>
        <div class="stat-card"><div class="stat-num" id="d-q">—</div><div class="stat-lbl">Questions</div></div>
        <div class="stat-card"><div class="stat-num" id="d-e">—</div><div class="stat-lbl">Exam Papers</div></div>
        <div class="stat-card"><div class="stat-num" id="d-y">—</div><div class="stat-lbl">Years</div></div>
      </div>

      <div class="charts-row">
        <div class="card">
          <div class="card-title">Questions per Subject</div>
          <div class="bar-chart" id="bar-chart"><div class="loader on"><div class="spin"></div></div></div>
        </div>
        <div class="card">
          <div class="card-title">Marks Distribution</div>
          <div class="donut-wrap" id="donut"><div class="loader on"><div class="spin"></div></div></div>
        </div>
      </div>

      <div class="page-header" style="margin-top:4px;">
        <h1 style="font-size:17px;">All Subjects</h1>
      </div>
      <div class="subjects-grid" id="subj-grid"><div class="loader on"><div class="spin"></div></div></div>
    </div>

    <!-- SEARCH -->
    <div id="search" class="section">
      <div class="page-header"><h1>Manual Search</h1><p>Search questions with filters and similarity matching</p></div>

      <div class="search-card">
        <div class="field-row">
          <div class="field" style="flex:2; min-width:200px;">
            <label>Search Question</label>
            <input class="inp inp-lg" id="sq" type="text" placeholder="e.g. What is exception handling?" onkeydown="if(event.key==='Enter')doSearch()">
          </div>
          <div class="field">
            <label>Subject</label>
            <select class="inp inp-sm" id="f-subject"><option value="">All Subjects</option></select>
          </div>
          <button class="btn btn-primary" onclick="doSearch()">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
            Search
          </button>
          <button class="btn btn-ghost" onclick="clearSearch()">Clear</button>
        </div>

        <div class="divider"></div>

        <div class="filter-row">
          <div class="field"><label>Marks</label>
            <select class="inp inp-xs" id="f-marks" style="height:34px;font-size:12px;">
              <option value="">Any</option><option value="1">1</option><option value="2">2</option><option value="4">4</option><option value="5">5</option><option value="6">6</option>
            </select>
          </div>
          <div class="field"><label>Exam Type</label>
            <select class="inp inp-xs" id="f-exam" style="height:34px;font-size:12px;">
              <option value="">All</option><option value="Summer">Summer</option><option value="Winter">Winter</option>
            </select>
          </div>
          <div class="field"><label>Year</label>
            <select class="inp inp-xs" id="f-year" style="height:34px;font-size:12px;">
              <option value="">Any</option><option value="2024">2024</option><option value="2025">2025</option>
            </select>
          </div>
          <div class="field"><label>Section</label>
            <select class="inp inp-xs" id="f-section" style="height:34px;font-size:12px;">
              <option value="">All</option><option value="A">A</option><option value="B">B</option>
            </select>
          </div>
          <div class="field"><label>Min Similarity %</label>
            <input type="number" class="inp inp-xs" id="f-sim" value="70" min="0" max="100" style="height:34px;font-size:12px;">
          </div>
        </div>
      </div>

      <div class="results-card">
        <div class="results-head">
          <div class="results-title">Results <span class="count-badge" id="s-count">0</span></div>
        </div>
        <div class="loader" id="s-loader"><div class="spin"></div><div class="spin-txt">Searching…</div></div>
        <div class="empty" id="s-empty"><div class="empty-ico">🔍</div><p>Type a question and press Search</p></div>
        <div class="tbl-wrap">
          <table id="s-table" style="display:none;">
            <thead><tr><th>#</th><th>Question</th><th>Subject</th><th>Q No.</th><th>Sec</th><th>Exam</th><th>Marks</th><th>Similarity</th></tr></thead>
            <tbody id="s-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- CHECK PAPER -->
    <div id="check" class="section">
      <div class="page-header"><h1>Check Paper</h1><p>Upload an exam paper to check which questions exist in the bank</p></div>
      <div class="notice">ℹ Questions with ≥70% match are marked <strong>Found</strong></div>

      <div class="card" style="margin-bottom:16px;">
        <div class="upload-zone" id="c-zone">
          <input type="file" id="c-file" accept=".pdf,.docx" onchange="onFile(this,'c-name')">
          <div class="upload-ico">📄</div>
          <div class="upload-title">Drop PDF or DOCX here</div>
          <div class="upload-sub">or click to browse</div>
          <div class="upload-name" id="c-name"></div>
        </div>
        <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap;">
          <button class="btn btn-primary" onclick="doCheck()">Analyse Paper</button>
          <button class="btn btn-ghost" onclick="clearCheck()">Clear</button>
        </div>
      </div>

      <div class="results-card">
        <div class="results-head">
          <div class="results-title">Results <span class="count-badge" id="c-count">0</span></div>
          <div id="c-summary" style="display:flex;gap:8px;"></div>
        </div>
        <div class="loader" id="c-loader"><div class="spin"></div><div class="spin-txt">Analysing…</div></div>
        <div class="empty" id="c-empty"><div class="empty-ico">📄</div><p>Upload a question paper to begin</p></div>
        <div class="tbl-wrap">
          <table id="c-table" style="display:none;">
            <thead><tr><th>#</th><th>Question</th><th>Status</th><th>Subject</th><th>Exam</th><th>Year</th><th>Marks</th><th>Similarity</th></tr></thead>
            <tbody id="c-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- INSERT -->
    <div id="insert" class="section">
      <div class="page-header"><h1>Insert Questions</h1><p>Extract questions from PDF/DOCX and add to the bank — duplicates skipped automatically</p></div>

      <div class="card" style="margin-bottom:16px;">
        <div class="field-row" style="margin-bottom:16px;">
          <div class="field" style="flex:1;">
            <label>Target Subject Sheet</label>
            <select class="inp" id="i-subject" style="width:100%;"><option value="">— Select Subject —</option></select>
          </div>
          <div class="field" style="flex:1;">
            <label>Or Create New Subject</label>
            <input class="inp" type="text" id="i-new" placeholder="New subject name…" style="width:100%;">
          </div>
        </div>
        <div class="upload-zone" id="i-zone">
          <input type="file" id="i-file" accept=".pdf,.docx" onchange="onFile(this,'i-name')">
          <div class="upload-ico">📂</div>
          <div class="upload-title">Drop PDF or DOCX to import</div>
          <div class="upload-sub">Questions extracted and de-duplicated automatically</div>
          <div class="upload-name" id="i-name"></div>
        </div>
        <div style="margin-top:14px;">
          <button class="btn btn-primary" onclick="doInsert()">Insert Questions</button>
        </div>
      </div>

      <div class="results-card" id="i-result-card" style="display:none;">
        <div class="results-head"><div class="results-title">Insert Report</div></div>
        <div style="padding:20px;" id="i-report"></div>
      </div>
    </div>

  </div>
</div>

<div class="toast-stack" id="toasts"></div>

<script>
// ── INIT ──
window.addEventListener('DOMContentLoaded', () => { loadStats(); loadDropdowns(); });

// ── NAV ──
function show(id) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  const nb = document.getElementById('nav-' + id);
  if (nb) nb.classList.add('active');
  const t = { dashboard:'Dashboard', search:'Manual Search', check:'Check Paper', insert:'Insert Questions' };
  document.getElementById('page-title').textContent = t[id] || '';
  closeSidebar();
}

// ── MOBILE SIDEBAR ──
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('overlay').classList.toggle('open');
}
function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('overlay').classList.remove('open');
}

// ── STATS ──
async function loadStats() {
  const d = await fetch('/stats').then(r => r.json());
  ['ss-sub','d-sub'].forEach(id => document.getElementById(id).textContent = d.subjects);
  ['ss-q','d-q'].forEach(id => document.getElementById(id).textContent = d.questions);
  ['ss-e','d-e'].forEach(id => document.getElementById(id).textContent = d.unique_exams);
  document.getElementById('d-y').textContent = d.years;
  document.getElementById('top-q').textContent = d.questions + ' Questions';
  buildBars(d.subject_counts);
  buildDonut(d.marks_dist);
  buildSubjCards(d.subject_counts);
}

// ── BAR CHART ──
const COLORS = ['#2563eb','#7c3aed','#059669','#d97706','#dc2626','#0891b2','#be185d'];

function buildBars(counts) {
  const el = document.getElementById('bar-chart');
  el.innerHTML = '';
  const entries = Object.entries(counts);
  const max = Math.max(...entries.map(e => e[1]));
  entries.forEach(([name, val], i) => {
    const pct = Math.round((val / max) * 100);
    const short = name.length > 20 ? name.slice(0, 18) + '…' : name;
    const row = document.createElement('div');
    row.className = 'bar-row';
    row.innerHTML = `<div class="bar-lbl" title="${name}">${short}</div><div class="bar-track"><div class="bar-fill" style="width:0%;background:${COLORS[i%COLORS.length]};" data-w="${pct}"></div></div><div class="bar-val">${val}</div>`;
    el.appendChild(row);
  });
  setTimeout(() => el.querySelectorAll('.bar-fill').forEach(b => b.style.width = b.dataset.w + '%'), 60);
}

// ── DONUT ──
function buildDonut(dist) {
  const el = document.getElementById('donut');
  const entries = Object.entries(dist).sort((a,b) => +a[0] - +b[0]);
  const total = entries.reduce((s,[,v]) => s+v, 0);
  const sz = 110, r = 42, cx = 55, cy = 55, c = 2*Math.PI*r;
  let svg = `<svg width="${sz}" height="${sz}" viewBox="0 0 ${sz} ${sz}" style="flex-shrink:0;"><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#f3f4f6" stroke-width="16"/>`;
  let off = 0;
  entries.forEach(([,v], i) => {
    const d = (v/total)*c;
    svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${COLORS[i%COLORS.length]}" stroke-width="16" stroke-dasharray="${d} ${c-d}" stroke-dashoffset="${c*0.25-off}" style="transform:rotate(-90deg);transform-origin:${cx}px ${cy}px;"/>`;
    off += d;
  });
  svg += `<text x="${cx}" y="${cy}" text-anchor="middle" dy="0.35em" font-family="Plus Jakarta Sans,sans-serif" font-weight="700" font-size="18" fill="#111827">${total}</text></svg>`;
  let leg = '<div class="donut-legend">';
  entries.forEach(([mark,count],i) => {
    leg += `<div class="legend-row"><div class="legend-dot" style="background:${COLORS[i%COLORS.length]};"></div><span class="legend-name">${mark}M</span><span class="legend-val">${count}</span></div>`;
  });
  el.innerHTML = svg + leg + '</div>';
}

// ── SUBJECT CARDS ──
const ICONS = ['📡','🤖','🌐','📊','🧮','🔬','⚡'];
function buildSubjCards(counts) {
  const grid = document.getElementById('subj-grid');
  grid.innerHTML = '';
  Object.entries(counts).forEach(([name, count], i) => {
    const c = document.createElement('div');
    c.className = 'subj-card';
    c.innerHTML = `<div class="subj-icon">${ICONS[i%ICONS.length]}</div><div style="flex:1;min-width:0;"><div class="subj-name">${name}</div><div class="subj-count">${count} questions</div></div><div class="subj-num">${count}</div>`;
    c.onclick = () => { document.getElementById('f-subject').value = name; document.getElementById('sq').value = ''; show('search'); doSearch(); };
    grid.appendChild(c);
  });
}

// ── DROPDOWNS ──
async function loadDropdowns() {
  const subjects = await fetch('/subjects').then(r => r.json());
  ['f-subject','i-subject'].forEach(id => {
    const el = document.getElementById(id);
    subjects.forEach(s => { const o = document.createElement('option'); o.value=s; o.textContent=s; el.appendChild(o); });
  });
}

// ── SEARCH ──
async function doSearch() {
  const q = document.getElementById('sq').value.trim();
  const payload = {
    q, subject: document.getElementById('f-subject').value,
    marks: document.getElementById('f-marks').value,
    exam:  document.getElementById('f-exam').value,
    year:  document.getElementById('f-year').value,
    section: document.getElementById('f-section').value,
    min_sim: parseInt(document.getElementById('f-sim').value)||0
  };
  document.getElementById('s-loader').classList.add('on');
  document.getElementById('s-table').style.display='none';
  document.getElementById('s-empty').style.display='none';
  const data = await fetch('/search',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}).then(r=>r.json());
  document.getElementById('s-loader').classList.remove('on');
  document.getElementById('s-count').textContent = data.length;
  if (!data.length) { document.getElementById('s-empty').style.display='block'; document.getElementById('s-empty').innerHTML='<div class="empty-ico">🔍</div><p>No questions matched</p>'; return; }
  document.getElementById('s-table').style.display='table';
  document.getElementById('s-tbody').innerHTML = data.map((r,i) => `<tr>
    <td style="color:var(--muted);font-size:11px;">${i+1}</td>
    <td class="td-q">${q?hlite(r.question,q):r.question}</td>
    <td><span class="pill pill-blue">${r.subject}</span></td>
    <td><span class="pill pill-gray">${r.q_no}</span></td>
    <td><span class="pill pill-orange">${r.section}</span></td>
    <td><span class="pill pill-amber">${r.exam_type} ${r.year}</span></td>
    <td><span class="marks-num">${r.marks}</span></td>
    <td>${simBar(r.similarity)}</td>
  </tr>`).join('');
}

function clearSearch() {
  ['sq','f-subject','f-marks','f-exam','f-year','f-section'].forEach(id => document.getElementById(id).value='');
  document.getElementById('f-sim').value='70';
  document.getElementById('s-count').textContent='0';
  document.getElementById('s-table').style.display='none';
  document.getElementById('s-empty').style.display='block';
  document.getElementById('s-empty').innerHTML='<div class="empty-ico">🔍</div><p>Type a question and press Search</p>';
}

// ── CHECK ──
async function doCheck() {
  const file = document.getElementById('c-file').files[0];
  if (!file) { toast('Please upload a PDF or DOCX file','err'); return; }
  document.getElementById('c-loader').classList.add('on');
  document.getElementById('c-table').style.display='none';
  document.getElementById('c-empty').style.display='none';
  const fd = new FormData(); fd.append('file',file);
  const data = await fetch('/multi',{method:'POST',body:fd}).then(r=>r.json());
  document.getElementById('c-loader').classList.remove('on');
  document.getElementById('c-count').textContent = data.length;
  if (!data.length) { document.getElementById('c-empty').style.display='block'; return; }
  const found = data.filter(d=>d.status==='Found').length;
  document.getElementById('c-summary').innerHTML=`<span class="chip chip-green">✓ ${found} Found</span><span class="chip" style="background:#fef2f2;color:var(--red);">✕ ${data.length-found} Not Found</span>`;
  document.getElementById('c-table').style.display='table';
  document.getElementById('c-tbody').innerHTML = data.map((r,i)=>`<tr>
    <td style="color:var(--muted);font-size:11px;">${i+1}</td>
    <td class="td-q">${r.question}</td>
    <td><span class="pill ${r.status==='Found'?'pill-green':'pill-red'}">${r.status==='Found'?'✓ Found':'✕ Not Found'}</span></td>
    <td>${r.subject?`<span class="pill pill-blue">${r.subject}</span>`:'—'}</td>
    <td>${r.exam_type?`<span class="pill pill-amber">${r.exam_type}</span>`:'—'}</td>
    <td style="color:var(--muted);font-size:11px;">${r.year||'—'}</td>
    <td>${r.marks?`<span class="marks-num">${r.marks}</span>`:'—'}</td>
    <td>${simBar(r.similarity)}</td>
  </tr>`).join('');
  toast(`Done: ${found} found, ${data.length-found} not found`,'ok');
}

function clearCheck() {
  document.getElementById('c-file').value='';
  document.getElementById('c-name').style.display='none';
  document.getElementById('c-count').textContent='0';
  document.getElementById('c-table').style.display='none';
  document.getElementById('c-empty').style.display='block';
  document.getElementById('c-empty').innerHTML='<div class="empty-ico">📄</div><p>Upload a question paper to begin</p>';
  document.getElementById('c-summary').innerHTML='';
}

// ── INSERT ──
async function doInsert() {
  const file = document.getElementById('i-file').files[0];
  let subject = document.getElementById('i-subject').value;
  const ns = document.getElementById('i-new').value.trim();
  if (ns) subject = ns;
  if (!file) { toast('Please upload a file','err'); return; }
  if (!subject) { toast('Please select or enter a subject','err'); return; }
  const fd = new FormData(); fd.append('file',file); fd.append('subject',subject);
  const data = await fetch('/insert',{method:'POST',body:fd}).then(r=>r.json());
  document.getElementById('i-result-card').style.display='block';
  document.getElementById('i-report').innerHTML=`
    <div style="display:flex;gap:14px;flex-wrap:wrap;">
      <div class="stat-card" style="flex:1;min-width:120px;"><div class="stat-num">${data.inserted}</div><div class="stat-lbl">Inserted</div></div>
      <div class="stat-card" style="flex:1;min-width:120px;"><div class="stat-num">${data.skipped}</div><div class="stat-lbl">Skipped</div></div>
      <div class="stat-card" style="flex:1;min-width:120px;"><div class="stat-num">${data.total}</div><div class="stat-lbl">Extracted</div></div>
    </div>
    <p style="margin-top:12px;font-size:12px;color:var(--muted);">Subject: <strong>${subject}</strong></p>`;
  toast(data.msg,'ok');
  loadStats();
}

// ── DOWNLOAD ──
function doDownload() { window.location.href='/download'; toast('Downloading…','inf'); }

// ── HELPERS ──
function onFile(input, nameId) {
  const el = document.getElementById(nameId);
  el.textContent = input.files[0] ? '📎 ' + input.files[0].name : '';
  el.style.display = input.files[0] ? 'block' : 'none';
}

function simBar(sim) {
  const p = Math.round(sim);
  const c = p>=90?'#059669':p>=70?'#2563eb':'#d97706';
  return `<div class="sim-wrap"><div class="sim-bar"><div class="sim-fill" style="width:${p}%;background:${c};"></div></div><span class="sim-txt">${p}%</span></div>`;
}

function hlite(text, q) {
  if (!q) return text;
  q.toLowerCase().split(/\s+/).filter(Boolean).forEach(w => {
    text = text.replace(new RegExp('('+w.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi'),'<mark>$1</mark>');
  });
  return text;
}

function toast(msg, type='ok') {
  const icons={ok:'✓',err:'✕',inf:'ℹ'};
  const cls={ok:'toast-ok',err:'toast-err',inf:'toast-inf'};
  const el = document.createElement('div');
  el.className = `toast ${cls[type]||'toast-ok'}`;
  el.innerHTML = `<span>${icons[type]}</span>${msg}`;
  document.getElementById('toasts').appendChild(el);
  setTimeout(()=>el.remove(), 3000);
}

// Drag & drop
['c-zone','i-zone'].forEach(id => {
  const z = document.getElementById(id); if (!z) return;
  z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('over'); });
  z.addEventListener('dragleave', () => z.classList.remove('over'));
  z.addEventListener('drop', e => {
    e.preventDefault(); z.classList.remove('over');
    const inp = z.querySelector('input[type=file]');
    if (inp && e.dataTransfer.files.length) {
      const dt = new DataTransfer(); dt.items.add(e.dataTransfer.files[0]);
      inp.files = dt.files;
      const m = inp.getAttribute('onchange')?.match(/'([^']+)'/);
      if (m) onFile(inp, m[1]);
    }
  });
});
</script>
</body>
</html>"""

# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────

@app.route("/")
def home():
    return render_template_string(HTML)

@app.route("/stats")
def stats():
    db = load_db()
    subject_counts, marks_dist, unique_exams, years = {}, {}, set(), set()
    for item in db:
        subject_counts[item["subject"]] = subject_counts.get(item["subject"], 0) + 1
        m = str(item["marks"])
        marks_dist[m] = marks_dist.get(m, 0) + 1
        if item["exam"]:  unique_exams.add(item["exam"])
        if item["year"]:  years.add(item["year"])
    return jsonify({"subjects": len(subject_counts), "questions": len(db),
                    "unique_exams": len(unique_exams), "years": len(years),
                    "subject_counts": subject_counts, "marks_dist": marks_dist})

@app.route("/subjects")
def subjects():
    sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    return jsonify(list(sheets.keys()))

@app.route("/search", methods=["POST"])
def search():
    body    = request.json
    q       = body.get("q","").strip()
    subject = body.get("subject","")
    marks   = body.get("marks","")
    exam    = body.get("exam","")
    year    = body.get("year","")
    section = body.get("section","")
    min_sim = float(body.get("min_sim",70)) / 100
    db, res = load_db(), []
    for item in db:
        if subject and item["subject"] != subject:        continue
        if marks   and str(item["marks"]) != str(marks): continue
        if exam    and item["exam_type"] != exam:         continue
        if year    and item["year"] != year:              continue
        if section and item["section"] != section:        continue
        sc = similarity(q, item["question"]) if q else 1.0
        if sc < min_sim: continue
        res.append({**item, "similarity": round(sc * 100, 2)})
    res.sort(key=lambda x: x["similarity"], reverse=True)
    return jsonify(res)

@app.route("/multi", methods=["POST"])
def multi():
    file = request.files["file"]
    qs   = extract_pdf(file) if file.filename.endswith(".pdf") else extract_docx(file)
    db, out = load_db(), []
    for q in qs:
        best, best_item = 0, None
        for item in db:
            sc = similarity(q, item["question"])
            if sc > best: best, best_item = sc, item
        if best >= 0.7 and best_item:
            out.append({"question":q,"status":"Found","subject":best_item["subject"],
                        "exam_type":best_item["exam_type"],"year":best_item["year"],
                        "marks":best_item["marks"],"similarity":round(best*100,2)})
        else:
            out.append({"question":q,"status":"Not Found","subject":"",
                        "exam_type":"","year":"","marks":"","similarity":round(best*100,2)})
    return jsonify(out)

@app.route("/insert", methods=["POST"])
def insert_questions():
    file, subject = request.files["file"], request.form["subject"]
    qs = extract_pdf(file) if file.filename.endswith(".pdf") else extract_docx(file)
    sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    if subject not in sheets:
        sheets[subject] = pd.DataFrame(columns=["Serial No","Subject","Exam","Section","Question No","Question","Marks"])
    df = sheets[subject]
    existing = df["Question"].astype(str).tolist() if "Question" in df.columns else []
    inserted = skipped = 0
    for q in qs:
        if any(similarity(q, ex) >= 0.7 for ex in existing):
            skipped += 1
        else:
            df = pd.concat([df, pd.DataFrame([{"Serial No":len(df)+1,"Subject":subject,"Exam":"Unknown",
                "Section":"A","Question No":f"Q{len(df)+1}","Question":q,"Marks":0}])], ignore_index=True)
            existing.append(q); inserted += 1
    sheets[subject] = df
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        for sname, sdf in sheets.items(): sdf.to_excel(writer, sheet_name=sname, index=False)
    return jsonify({"inserted":inserted,"skipped":skipped,"total":len(qs),
                    "msg":f"{inserted} inserted, {skipped} duplicates skipped"})

@app.route("/download")
def download():
    return send_file(EXCEL_FILE, as_attachment=True)

# ─────────────────────────────────────────────
# RUN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=False)
