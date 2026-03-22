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
# HTML TEMPLATE
# ─────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ExamInsight — Question Bank Intelligence</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=DM+Mono:wght@300;400;500&family=Instrument+Serif:ital@0;1&display=swap" rel="stylesheet">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --ink:#0a0a0f;--ink2:#1a1a28;--ink3:#2e2e45;--muted:#6b6b8a;
  --border:#e2e2ed;--surface:#f7f7fc;--white:#ffffff;
  --accent:#5b4cff;--accent2:#ff4c8b;--accent3:#00c9a7;
  --gold:#f5a623;--warn:#ff6b35;
  --radius:14px;--shadow:0 4px 24px rgba(10,10,15,0.08);
  --shadow-lg:0 12px 48px rgba(10,10,15,0.14);
}
html{scroll-behavior:smooth;}
body{font-family:'DM Mono',monospace;background:var(--surface);color:var(--ink);min-height:100vh;overflow-x:hidden;}

.sidebar{position:fixed;left:0;top:0;bottom:0;width:260px;background:var(--ink);display:flex;flex-direction:column;z-index:100;padding:0 0 24px;}
.logo{padding:32px 28px 24px;border-bottom:1px solid rgba(255,255,255,0.07);}
.logo-mark{font-family:'Syne',sans-serif;font-weight:800;font-size:22px;color:var(--white);letter-spacing:-0.5px;display:flex;align-items:center;gap:10px;}
.logo-icon{width:36px;height:36px;background:linear-gradient(135deg,var(--accent),var(--accent2));border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;}
.logo-sub{font-size:10px;color:var(--muted);letter-spacing:2px;text-transform:uppercase;margin-top:4px;font-family:'DM Mono',monospace;}
.nav{padding:20px 16px;flex:1;display:flex;flex-direction:column;gap:4px;}
.nav-label{font-size:9px;letter-spacing:2.5px;text-transform:uppercase;color:rgba(255,255,255,0.25);padding:8px 12px 4px;font-family:'DM Mono',monospace;}
.nav-item{display:flex;align-items:center;gap:12px;padding:11px 14px;border-radius:10px;cursor:pointer;color:rgba(255,255,255,0.5);font-size:13px;font-family:'DM Mono',monospace;transition:all 0.2s;border:none;background:none;width:100%;text-align:left;position:relative;}
.nav-item:hover{background:rgba(255,255,255,0.06);color:rgba(255,255,255,0.85);}
.nav-item.active{background:rgba(91,76,255,0.18);color:#fff;}
.nav-item.active::before{content:'';position:absolute;left:0;top:50%;transform:translateY(-50%);width:3px;height:20px;background:var(--accent);border-radius:0 3px 3px 0;}
.nav-icon{font-size:16px;width:20px;text-align:center;}
.sidebar-stats{margin:0 16px;padding:16px;background:rgba(255,255,255,0.04);border-radius:var(--radius);border:1px solid rgba(255,255,255,0.06);}
.sidebar-stats-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;}
.sidebar-stats-row:last-child{margin-bottom:0;}
.ss-label{font-size:10px;color:rgba(255,255,255,0.35);text-transform:uppercase;letter-spacing:1px;}
.ss-value{font-family:'Syne',sans-serif;font-weight:700;font-size:18px;color:var(--white);}

.main{margin-left:260px;min-height:100vh;display:flex;flex-direction:column;}
.topbar{position:sticky;top:0;z-index:50;background:rgba(247,247,252,0.85);backdrop-filter:blur(12px);border-bottom:1px solid var(--border);padding:0 40px;height:64px;display:flex;align-items:center;justify-content:space-between;}
.page-title{font-family:'Syne',sans-serif;font-weight:700;font-size:18px;color:var(--ink);}
.topbar-right{display:flex;align-items:center;gap:12px;}
.badge{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:100px;font-size:11px;font-weight:500;letter-spacing:0.5px;}
.badge-accent{background:rgba(91,76,255,0.1);color:var(--accent);}
.badge-green{background:rgba(0,201,167,0.1);color:var(--accent3);}
.content{padding:40px;flex:1;}

.section{display:none;animation:fadeUp 0.35s ease;}
.section.active{display:block;}
@keyframes fadeUp{from{opacity:0;transform:translateY(16px);}to{opacity:1;transform:translateY(0);}}

.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:20px;margin-bottom:40px;}
.stat-card{background:var(--white);border-radius:var(--radius);padding:28px;border:1px solid var(--border);box-shadow:var(--shadow);position:relative;overflow:hidden;}
.stat-card::after{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.stat-card.c1::after{background:linear-gradient(90deg,var(--accent),var(--accent2));}
.stat-card.c2::after{background:linear-gradient(90deg,var(--accent3),#00a8ff);}
.stat-card.c3::after{background:linear-gradient(90deg,var(--gold),var(--warn));}
.stat-card.c4::after{background:linear-gradient(90deg,var(--accent2),var(--warn));}
.stat-number{font-family:'Syne',sans-serif;font-weight:800;font-size:42px;color:var(--ink);line-height:1;margin-bottom:6px;}
.stat-label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:1.5px;}
.stat-icon{position:absolute;right:24px;top:24px;font-size:28px;opacity:0.12;}

.subjects-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px;margin-bottom:40px;}
.subject-card{background:var(--white);border-radius:var(--radius);padding:22px 24px;border:1px solid var(--border);display:flex;align-items:center;gap:16px;transition:all 0.2s;cursor:pointer;}
.subject-card:hover{border-color:var(--accent);box-shadow:0 0 0 3px rgba(91,76,255,0.08),var(--shadow);transform:translateY(-1px);}
.subject-dot{width:44px;height:44px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;}
.subject-info{flex:1;min-width:0;}
.subject-name{font-family:'Syne',sans-serif;font-weight:600;font-size:14px;color:var(--ink);margin-bottom:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.subject-meta{font-size:11px;color:var(--muted);}
.subject-count{font-family:'Syne',sans-serif;font-weight:700;font-size:20px;color:var(--accent);}

.section-heading{margin-bottom:28px;}
.section-heading h2{font-family:'Syne',sans-serif;font-weight:800;font-size:28px;color:var(--ink);margin-bottom:6px;}
.section-heading p{font-size:13px;color:var(--muted);line-height:1.6;}

.search-box{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);padding:28px;margin-bottom:24px;}
.input-row{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap;}
.input-group{display:flex;flex-direction:column;gap:6px;flex:1;min-width:200px;}
.input-label{font-size:10px;letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);}
.input-field{height:44px;border:1.5px solid var(--border);border-radius:10px;padding:0 14px;font-family:'DM Mono',monospace;font-size:13px;color:var(--ink);background:var(--surface);transition:all 0.2s;outline:none;width:100%;}
.input-field:focus{border-color:var(--accent);background:var(--white);box-shadow:0 0 0 3px rgba(91,76,255,0.1);}
select.input-field{cursor:pointer;}
.btn{height:44px;padding:0 24px;border-radius:10px;border:none;cursor:pointer;font-family:'Syne',sans-serif;font-weight:600;font-size:13px;transition:all 0.2s;display:inline-flex;align-items:center;gap:8px;white-space:nowrap;}
.btn-primary{background:var(--accent);color:var(--white);}
.btn-primary:hover{background:#4a3ce8;transform:translateY(-1px);box-shadow:0 4px 16px rgba(91,76,255,0.35);}
.btn-secondary{background:var(--surface);color:var(--ink);border:1.5px solid var(--border);}
.btn-secondary:hover{border-color:var(--accent);color:var(--accent);}
.filters-row{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;padding-top:14px;border-top:1px solid var(--border);}

.results-wrap{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);overflow:hidden;}
.results-header{padding:18px 24px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;}
.results-title{font-family:'Syne',sans-serif;font-weight:700;font-size:14px;color:var(--ink);display:flex;align-items:center;gap:10px;}
.result-count{display:inline-flex;align-items:center;justify-content:center;min-width:24px;height:24px;padding:0 8px;background:var(--accent);color:var(--white);border-radius:100px;font-size:11px;font-weight:700;}
.table-wrap{overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:12.5px;}
thead th{background:var(--surface);padding:12px 16px;text-align:left;font-size:10px;letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);font-weight:500;border-bottom:1px solid var(--border);white-space:nowrap;}
tbody tr{border-bottom:1px solid var(--border);transition:background 0.15s;}
tbody tr:last-child{border-bottom:none;}
tbody tr:hover{background:rgba(91,76,255,0.03);}
td{padding:14px 16px;color:var(--ink2);vertical-align:middle;}
.td-question{max-width:380px;line-height:1.5;}
.pill{display:inline-flex;align-items:center;padding:3px 10px;border-radius:100px;font-size:10.5px;font-weight:500;white-space:nowrap;}
.pill-found{background:rgba(0,201,167,0.12);color:#00a389;}
.pill-notfound{background:rgba(255,76,139,0.12);color:#d93a7a;}
.pill-subject{background:rgba(91,76,255,0.1);color:var(--accent);}
.pill-exam{background:rgba(245,166,35,0.12);color:#c4820a;}
.pill-section{background:rgba(255,107,53,0.1);color:var(--warn);}
.marks-badge{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:50%;font-weight:700;font-size:12px;background:var(--surface);border:1.5px solid var(--border);color:var(--ink);}

.upload-zone{border:2px dashed var(--border);border-radius:var(--radius);padding:48px;text-align:center;cursor:pointer;transition:all 0.2s;background:var(--white);position:relative;}
.upload-zone:hover,.upload-zone.drag-over{border-color:var(--accent);background:rgba(91,76,255,0.03);}
.upload-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
.upload-icon{font-size:40px;margin-bottom:14px;}
.upload-title{font-family:'Syne',sans-serif;font-weight:700;font-size:16px;color:var(--ink);margin-bottom:6px;}
.upload-sub{font-size:12px;color:var(--muted);}
.upload-file-name{display:none;margin-top:14px;padding:10px 16px;background:rgba(91,76,255,0.08);border-radius:8px;font-size:12px;color:var(--accent);font-weight:500;}

.sim-bar-wrap{display:flex;align-items:center;gap:8px;}
.sim-bar{flex:1;height:4px;background:var(--surface);border-radius:100px;overflow:hidden;max-width:60px;}
.sim-bar-fill{height:100%;border-radius:100px;}
.sim-pct{font-size:11px;color:var(--muted);white-space:nowrap;}

.insert-card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);padding:32px;margin-bottom:24px;}
.insert-card h3{font-family:'Syne',sans-serif;font-weight:700;font-size:16px;color:var(--ink);margin-bottom:20px;padding-bottom:16px;border-bottom:1px solid var(--border);}

.toast-wrap{position:fixed;bottom:28px;right:28px;z-index:9999;display:flex;flex-direction:column;gap:10px;pointer-events:none;}
.toast{padding:14px 20px;border-radius:12px;font-size:13px;font-weight:500;display:flex;align-items:center;gap:10px;pointer-events:auto;box-shadow:var(--shadow-lg);animation:toastIn 0.3s cubic-bezier(0.16,1,0.3,1);max-width:320px;}
.toast-success{background:var(--ink);color:var(--white);}
.toast-error{background:#fff0f4;color:var(--accent2);border:1px solid rgba(255,76,139,0.2);}
.toast-info{background:rgba(91,76,255,0.12);color:var(--accent);border:1px solid rgba(91,76,255,0.2);}
@keyframes toastIn{from{opacity:0;transform:translateX(24px) scale(0.95);}to{opacity:1;transform:translateX(0) scale(1);}}

.loader-wrap{display:none;align-items:center;justify-content:center;padding:48px;gap:14px;flex-direction:column;}
.loader-wrap.active{display:flex;}
.spinner{width:36px;height:36px;border:3px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin 0.75s linear infinite;}
@keyframes spin{to{transform:rotate(360deg);}}
.loader-text{font-size:12px;color:var(--muted);}

.empty-state{text-align:center;padding:56px 24px;color:var(--muted);}
.empty-state .empty-icon{font-size:44px;margin-bottom:14px;opacity:0.4;}
.empty-state p{font-family:'Instrument Serif',serif;font-style:italic;font-size:16px;}

::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:100px;}
::-webkit-scrollbar-thumb:hover{background:var(--muted);}
mark{background:rgba(91,76,255,0.15);color:var(--accent);border-radius:3px;padding:0 2px;}

.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:28px;}
.chart-card{background:var(--white);border-radius:var(--radius);border:1px solid var(--border);box-shadow:var(--shadow);padding:24px;}
.chart-card h3{font-family:'Syne',sans-serif;font-weight:700;font-size:14px;color:var(--ink);margin-bottom:20px;}
.bar-chart{display:flex;flex-direction:column;gap:10px;}
.bar-row{display:flex;align-items:center;gap:10px;}
.bar-label{font-size:11px;color:var(--muted);width:110px;flex-shrink:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.bar-track{flex:1;height:8px;background:var(--surface);border-radius:100px;overflow:hidden;}
.bar-fill{height:100%;border-radius:100px;transition:width 0.6s cubic-bezier(0.34,1.56,0.64,1);}
.bar-val{font-size:11px;font-weight:600;color:var(--ink);width:28px;text-align:right;flex-shrink:0;}
.donut-wrap{display:flex;align-items:center;gap:24px;}
.donut-legend{display:flex;flex-direction:column;gap:10px;}
.legend-item{display:flex;align-items:center;gap:8px;font-size:12px;}
.legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;}
.legend-name{color:var(--muted);}
.legend-val{font-weight:600;color:var(--ink);margin-left:auto;padding-left:8px;}
.q-no-badge{display:inline-flex;align-items:center;padding:2px 8px;background:var(--surface);border:1px solid var(--border);border-radius:6px;font-size:10.5px;color:var(--muted);white-space:nowrap;}
.notice{background:rgba(91,76,255,0.07);border:1px solid rgba(91,76,255,0.15);border-radius:10px;padding:14px 18px;font-size:12.5px;color:var(--accent);margin-bottom:20px;display:flex;align-items:center;gap:10px;}
</style>
</head>
<body>

<aside class="sidebar">
  <div class="logo">
    <div class="logo-mark"><div class="logo-icon">🎓</div>ExamInsight</div>
    <div class="logo-sub">Question Bank Intelligence</div>
  </div>
  <nav class="nav">
    <div class="nav-label">Main</div>
    <button class="nav-item active" onclick="show('dashboard')" id="nav-dashboard"><span class="nav-icon">◈</span> Dashboard</button>
    <button class="nav-item" onclick="show('search')" id="nav-search"><span class="nav-icon">⌕</span> Manual Search</button>
    <button class="nav-item" onclick="show('check')" id="nav-check"><span class="nav-icon">◎</span> Check Paper</button>
    <div class="nav-label" style="margin-top:8px;">Manage</div>
    <button class="nav-item" onclick="show('insert')" id="nav-insert"><span class="nav-icon">⊕</span> Insert Questions</button>
    <button class="nav-item" onclick="downloadExcel()"><span class="nav-icon">↓</span> Export Excel</button>
  </nav>
  <div class="sidebar-stats">
    <div class="sidebar-stats-row"><span class="ss-label">Subjects</span><span class="ss-value" id="ss-subjects">—</span></div>
    <div class="sidebar-stats-row"><span class="ss-label">Questions</span><span class="ss-value" id="ss-questions">—</span></div>
    <div class="sidebar-stats-row"><span class="ss-label">Exams</span><span class="ss-value" id="ss-exams">—</span></div>
  </div>
</aside>

<div class="main">
  <div class="topbar">
    <div class="page-title" id="page-title">Dashboard</div>
    <div class="topbar-right">
      <span class="badge badge-green" id="topbar-count">Loading...</span>
      <span class="badge badge-accent">B.E. Sem 6</span>
    </div>
  </div>

  <div class="content">

    <!-- DASHBOARD -->
    <div id="dashboard" class="section active">
      <div class="section-heading"><h2>Question Bank Overview</h2><p>Analytics across all subjects and exam papers</p></div>
      <div class="stats-grid">
        <div class="stat-card c1"><div class="stat-icon">📚</div><div class="stat-number" id="d-subjects">—</div><div class="stat-label">Total Subjects</div></div>
        <div class="stat-card c2"><div class="stat-icon">❓</div><div class="stat-number" id="d-questions">—</div><div class="stat-label">Total Questions</div></div>
        <div class="stat-card c3"><div class="stat-icon">📝</div><div class="stat-number" id="d-exams">—</div><div class="stat-label">Exam Papers</div></div>
        <div class="stat-card c4"><div class="stat-icon">⭐</div><div class="stat-number" id="d-years">—</div><div class="stat-label">Academic Years</div></div>
      </div>
      <div class="chart-row">
        <div class="chart-card"><h3>Questions per Subject</h3><div class="bar-chart" id="subject-bars"><div class="loader-wrap active"><div class="spinner"></div></div></div></div>
        <div class="chart-card"><h3>Marks Distribution</h3><div class="donut-wrap" id="marks-donut"><div class="loader-wrap active"><div class="spinner"></div></div></div></div>
      </div>
      <div class="section-heading" style="margin-top:8px;"><h2 style="font-size:20px;">All Subjects</h2></div>
      <div class="subjects-grid" id="subjects-grid"><div class="loader-wrap active"><div class="spinner"></div></div></div>
    </div>

    <!-- SEARCH -->
    <div id="search" class="section">
      <div class="section-heading"><h2>Manual Search</h2><p>Search any question across all subjects using 70% similarity matching</p></div>
      <div class="search-box">
        <div class="input-row">
          <div class="input-group" style="flex:2;">
            <label class="input-label">Search Question</label>
            <input class="input-field" id="search-q" type="text" placeholder="e.g. What is exception handling?" onkeydown="if(event.key==='Enter') doSearch()">
          </div>
          <div class="input-group">
            <label class="input-label">Subject</label>
            <select class="input-field" id="filter-subject"><option value="">All Subjects</option></select>
          </div>
          <button class="btn btn-primary" onclick="doSearch()">⌕ Search</button>
          <button class="btn btn-secondary" onclick="clearSearch()">✕ Clear</button>
        </div>
        <div class="filters-row">
          <div class="input-group" style="min-width:130px;flex:unset;">
            <label class="input-label">Marks</label>
            <select class="input-field" id="filter-marks" style="height:36px;font-size:12px;">
              <option value="">Any</option><option value="1">1 Mark</option><option value="2">2 Marks</option><option value="4">4 Marks</option><option value="5">5 Marks</option><option value="6">6 Marks</option>
            </select>
          </div>
          <div class="input-group" style="min-width:130px;flex:unset;">
            <label class="input-label">Exam Type</label>
            <select class="input-field" id="filter-exam" style="height:36px;font-size:12px;">
              <option value="">All</option><option value="Summer">Summer</option><option value="Winter">Winter</option>
            </select>
          </div>
          <div class="input-group" style="min-width:120px;flex:unset;">
            <label class="input-label">Year</label>
            <select class="input-field" id="filter-year" style="height:36px;font-size:12px;">
              <option value="">Any</option><option value="2024">2024</option><option value="2025">2025</option>
            </select>
          </div>
          <div class="input-group" style="min-width:120px;flex:unset;">
            <label class="input-label">Section</label>
            <select class="input-field" id="filter-section" style="height:36px;font-size:12px;">
              <option value="">All</option><option value="A">Section A</option><option value="B">Section B</option>
            </select>
          </div>
          <div class="input-group" style="min-width:150px;flex:unset;">
            <label class="input-label">Min Similarity %</label>
            <input type="number" class="input-field" id="filter-sim" value="70" min="0" max="100" style="height:36px;font-size:12px;">
          </div>
        </div>
      </div>
      <div class="results-wrap">
        <div class="results-header">
          <div class="results-title">Results <span class="result-count" id="search-count">0</span></div>
        </div>
        <div class="table-wrap">
          <div class="loader-wrap" id="search-loader"><div class="spinner"></div><div class="loader-text">Searching...</div></div>
          <div class="empty-state" id="search-empty"><div class="empty-icon">🔍</div><p>Type a question above and press Search</p></div>
          <table id="search-table" style="display:none;">
            <thead><tr><th>#</th><th>Question</th><th>Subject</th><th>Q No.</th><th>Section</th><th>Exam</th><th>Year</th><th>Marks</th><th>Similarity</th></tr></thead>
            <tbody id="search-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- CHECK PAPER -->
    <div id="check" class="section">
      <div class="section-heading"><h2>Check Question Paper</h2><p>Upload a PDF or DOCX exam paper to see which questions exist in the bank</p></div>
      <div class="notice">ℹ️ Questions matching ≥70% similarity are marked <strong>Found</strong>. Others are <strong>Not Found</strong>.</div>
      <div class="insert-card">
        <h3>Upload Question Paper</h3>
        <div class="upload-zone" id="check-zone">
          <input type="file" id="check-file" accept=".pdf,.docx" onchange="onFileChange(this,'check-fname')">
          <div class="upload-icon">📄</div>
          <div class="upload-title">Drop your PDF or DOCX here</div>
          <div class="upload-sub">or click to browse</div>
          <div class="upload-file-name" id="check-fname"></div>
        </div>
        <div style="margin-top:16px;display:flex;gap:10px;">
          <button class="btn btn-primary" onclick="doCheck()">◎ Analyse Paper</button>
          <button class="btn btn-secondary" onclick="clearCheck()">✕ Clear</button>
        </div>
      </div>
      <div class="results-wrap">
        <div class="results-header">
          <div class="results-title">Analysis Results <span class="result-count" id="check-count">0</span></div>
          <div style="display:flex;gap:10px;" id="check-summary"></div>
        </div>
        <div class="table-wrap">
          <div class="loader-wrap" id="check-loader"><div class="spinner"></div><div class="loader-text">Analysing paper...</div></div>
          <div class="empty-state" id="check-empty"><div class="empty-icon">📄</div><p>Upload a question paper to begin analysis</p></div>
          <table id="check-table" style="display:none;">
            <thead><tr><th>#</th><th>Question from Paper</th><th>Status</th><th>Matched Subject</th><th>Exam</th><th>Year</th><th>Marks</th><th>Similarity</th></tr></thead>
            <tbody id="check-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- INSERT -->
    <div id="insert" class="section">
      <div class="section-heading"><h2>Insert Questions</h2><p>Extract questions from PDF/DOCX and add to question bank — duplicates auto-skipped</p></div>
      <div class="insert-card">
        <h3>Add to Question Bank</h3>
        <div class="input-row" style="margin-bottom:20px;">
          <div class="input-group">
            <label class="input-label">Target Subject Sheet</label>
            <select class="input-field" id="insert-subject"><option value="">— Select Subject —</option></select>
          </div>
          <div class="input-group">
            <label class="input-label">Or Create New Subject</label>
            <input class="input-field" type="text" id="insert-new-subject" placeholder="Type new subject name...">
          </div>
        </div>
        <div class="upload-zone" id="insert-zone">
          <input type="file" id="insert-file" accept=".pdf,.docx" onchange="onFileChange(this,'insert-fname')">
          <div class="upload-icon">📂</div>
          <div class="upload-title">Drop PDF or DOCX to import questions</div>
          <div class="upload-sub">Questions are extracted and de-duplicated automatically</div>
          <div class="upload-file-name" id="insert-fname"></div>
        </div>
        <div style="margin-top:16px;"><button class="btn btn-primary" onclick="doInsert()">⊕ Insert Questions</button></div>
      </div>
      <div class="results-wrap" id="insert-results" style="display:none;">
        <div class="results-header"><div class="results-title">Insert Report</div></div>
        <div style="padding:28px;" id="insert-report"></div>
      </div>
    </div>

  </div>
</div>

<div class="toast-wrap" id="toast-wrap"></div>

<script>
window.addEventListener('DOMContentLoaded', () => { loadStats(); loadSubjectDropdowns(); });

function show(id) {
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  const navEl = document.getElementById('nav-' + id);
  if (navEl) navEl.classList.add('active');
  const titles = { dashboard:'Dashboard', search:'Manual Search', check:'Check Paper', insert:'Insert Questions' };
  document.getElementById('page-title').textContent = titles[id] || '';
}

async function loadStats() {
  const d = await fetch('/stats').then(r => r.json());
  document.getElementById('ss-subjects').textContent  = d.subjects;
  document.getElementById('ss-questions').textContent = d.questions;
  document.getElementById('ss-exams').textContent     = d.unique_exams;
  document.getElementById('d-subjects').textContent   = d.subjects;
  document.getElementById('d-questions').textContent  = d.questions;
  document.getElementById('d-exams').textContent      = d.unique_exams;
  document.getElementById('d-years').textContent      = d.years;
  document.getElementById('topbar-count').textContent = d.questions + ' Questions';
  buildSubjectCards(d.subject_counts);
  buildBarChart(d.subject_counts);
  buildDonut(d.marks_dist);
}

const subjectColors = [
  ['#5b4cff','#f0eeff','📡'],['#ff4c8b','#fff0f6','🤖'],['#00c9a7','#edfaf7','🌐'],
  ['#f5a623','#fff8ec','📊'],['#00a8ff','#e8f6ff','🧮'],['#ff6b35','#fff2ec','🔬'],['#7c3aed','#f3e8ff','⚡'],
];

function buildSubjectCards(counts) {
  const grid = document.getElementById('subjects-grid');
  grid.innerHTML = '';
  Object.entries(counts).forEach(([name, count], i) => {
    const [clr, bg, icon] = subjectColors[i % subjectColors.length];
    const card = document.createElement('div');
    card.className = 'subject-card';
    card.innerHTML = `<div class="subject-dot" style="background:${bg};color:${clr};">${icon}</div><div class="subject-info"><div class="subject-name">${name}</div><div class="subject-meta">${count} questions</div></div><div class="subject-count" style="color:${clr};">${count}</div>`;
    card.onclick = () => { document.getElementById('filter-subject').value = name; document.getElementById('search-q').value = ''; show('search'); doSearch(); };
    grid.appendChild(card);
  });
}

function buildBarChart(counts) {
  const wrap = document.getElementById('subject-bars');
  wrap.innerHTML = '';
  const entries = Object.entries(counts);
  const max = Math.max(...entries.map(e => e[1]));
  entries.forEach(([name, val], i) => {
    const [clr] = subjectColors[i % subjectColors.length];
    const pct = Math.round((val / max) * 100);
    const row = document.createElement('div');
    row.className = 'bar-row';
    row.innerHTML = `<div class="bar-label" title="${name}">${name.length>18?name.slice(0,16)+'…':name}</div><div class="bar-track"><div class="bar-fill" style="width:0%;background:${clr};" data-w="${pct}"></div></div><div class="bar-val">${val}</div>`;
    wrap.appendChild(row);
  });
  setTimeout(() => wrap.querySelectorAll('.bar-fill').forEach(b => b.style.width = b.dataset.w + '%'), 80);
}

function buildDonut(dist) {
  const wrap = document.getElementById('marks-donut');
  const entries = Object.entries(dist).sort((a,b) => Number(a[0])-Number(b[0]));
  const total = entries.reduce((s,[,v]) => s+v, 0);
  const colors = ['#5b4cff','#ff4c8b','#00c9a7','#f5a623','#00a8ff','#ff6b35'];
  const size=130, r=50, cx=65, cy=65, circ=2*Math.PI*r;
  let svg = `<svg class="donut-svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#f0f0f6" stroke-width="18"/>`;
  let offset = 0;
  entries.forEach(([mark, count], i) => {
    const dash = (count/total)*circ;
    svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${colors[i%colors.length]}" stroke-width="18" stroke-dasharray="${dash} ${circ-dash}" stroke-dashoffset="${circ*0.25-offset}" style="transform:rotate(-90deg);transform-origin:${cx}px ${cy}px;"/>`;
    offset += dash;
  });
  svg += `<text x="${cx}" y="${cy}" text-anchor="middle" dy="0.35em" font-family="Syne,sans-serif" font-weight="800" font-size="22" fill="#0a0a0f">${total}</text></svg>`;
  let legend = '<div class="donut-legend">';
  entries.forEach(([mark, count], i) => {
    legend += `<div class="legend-item"><div class="legend-dot" style="background:${colors[i%colors.length]};"></div><span class="legend-name">${mark} Mark${mark>1?'s':''}</span><span class="legend-val">${count}</span></div>`;
  });
  wrap.innerHTML = svg + legend + '</div>';
}

async function loadSubjectDropdowns() {
  const subjects = await fetch('/subjects').then(r => r.json());
  ['filter-subject','insert-subject'].forEach(id => {
    const el = document.getElementById(id);
    subjects.forEach(s => { const o = document.createElement('option'); o.value=s; o.textContent=s; el.appendChild(o); });
  });
}

async function doSearch() {
  const q       = document.getElementById('search-q').value.trim();
  const subject = document.getElementById('filter-subject').value;
  const marks   = document.getElementById('filter-marks').value;
  const exam    = document.getElementById('filter-exam').value;
  const year    = document.getElementById('filter-year').value;
  const section = document.getElementById('filter-section').value;
  const minSim  = parseInt(document.getElementById('filter-sim').value)||0;
  document.getElementById('search-loader').classList.add('active');
  document.getElementById('search-table').style.display='none';
  document.getElementById('search-empty').style.display='none';
  const data = await fetch('/search',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({q,subject,marks,exam,year,section,min_sim:minSim})}).then(r=>r.json());
  document.getElementById('search-loader').classList.remove('active');
  document.getElementById('search-count').textContent = data.length;
  if (!data.length) { document.getElementById('search-empty').style.display='block'; document.getElementById('search-empty').innerHTML='<div class="empty-icon">🔍</div><p>No questions matched your query</p>'; return; }
  document.getElementById('search-table').style.display='table';
  document.getElementById('search-tbody').innerHTML = data.map((r,i) => `<tr>
    <td style="color:var(--muted);font-size:11px;">${i+1}</td>
    <td class="td-question">${q?highlight(r.question,q):r.question}</td>
    <td><span class="pill pill-subject">${r.subject}</span></td>
    <td><span class="q-no-badge">${r.q_no}</span></td>
    <td><span class="pill pill-section">${r.section}</span></td>
    <td><span class="pill pill-exam">${r.exam_type} ${r.year}</span></td>
    <td style="color:var(--muted);font-size:11px;">${r.year}</td>
    <td><div class="marks-badge">${r.marks}</div></td>
    <td>${simBar(r.similarity)}</td>
  </tr>`).join('');
}

function clearSearch() {
  ['search-q','filter-subject','filter-marks','filter-exam','filter-year','filter-section'].forEach(id => document.getElementById(id).value='');
  document.getElementById('filter-sim').value='70';
  document.getElementById('search-count').textContent='0';
  document.getElementById('search-table').style.display='none';
  document.getElementById('search-empty').style.display='block';
  document.getElementById('search-empty').innerHTML='<div class="empty-icon">🔍</div><p>Type a question above and press Search</p>';
}

async function doCheck() {
  const file = document.getElementById('check-file').files[0];
  if (!file) { toast('Please upload a PDF or DOCX file','error'); return; }
  document.getElementById('check-loader').classList.add('active');
  document.getElementById('check-table').style.display='none';
  document.getElementById('check-empty').style.display='none';
  const fd = new FormData(); fd.append('file',file);
  const data = await fetch('/multi',{method:'POST',body:fd}).then(r=>r.json());
  document.getElementById('check-loader').classList.remove('active');
  document.getElementById('check-count').textContent=data.length;
  if (!data.length) { document.getElementById('check-empty').style.display='block'; document.getElementById('check-empty').innerHTML='<div class="empty-icon">📄</div><p>No questions could be extracted</p>'; return; }
  const found=data.filter(d=>d.status==='Found').length;
  document.getElementById('check-summary').innerHTML=`<span class="badge badge-green">✓ Found: ${found}</span><span class="badge" style="background:rgba(255,76,139,0.1);color:var(--accent2);">✕ Not Found: ${data.length-found}</span>`;
  document.getElementById('check-table').style.display='table';
  document.getElementById('check-tbody').innerHTML=data.map((r,i)=>`<tr>
    <td style="color:var(--muted);font-size:11px;">${i+1}</td>
    <td class="td-question">${r.question}</td>
    <td><span class="pill ${r.status==='Found'?'pill-found':'pill-notfound'}">${r.status==='Found'?'✓ Found':'✕ Not Found'}</span></td>
    <td>${r.subject?`<span class="pill pill-subject">${r.subject}</span>`:'—'}</td>
    <td>${r.exam_type?`<span class="pill pill-exam">${r.exam_type}</span>`:'—'}</td>
    <td style="color:var(--muted);font-size:11px;">${r.year||'—'}</td>
    <td>${r.marks?`<div class="marks-badge">${r.marks}</div>`:'—'}</td>
    <td>${r.similarity?simBar(r.similarity):'—'}</td>
  </tr>`).join('');
  toast(`Analysis complete: ${found} found, ${data.length-found} not found`,'success');
}

function clearCheck() {
  document.getElementById('check-file').value='';
  document.getElementById('check-fname').style.display='none';
  document.getElementById('check-count').textContent='0';
  document.getElementById('check-table').style.display='none';
  document.getElementById('check-empty').style.display='block';
  document.getElementById('check-empty').innerHTML='<div class="empty-icon">📄</div><p>Upload a question paper to begin analysis</p>';
}

async function doInsert() {
  const file = document.getElementById('insert-file').files[0];
  let subject = document.getElementById('insert-subject').value;
  const newSub = document.getElementById('insert-new-subject').value.trim();
  if (newSub) subject = newSub;
  if (!file) { toast('Please upload a PDF or DOCX file','error'); return; }
  if (!subject) { toast('Please select or enter a subject name','error'); return; }
  const fd = new FormData(); fd.append('file',file); fd.append('subject',subject);
  const data = await fetch('/insert',{method:'POST',body:fd}).then(r=>r.json());
  document.getElementById('insert-results').style.display='block';
  document.getElementById('insert-report').innerHTML=`
    <div style="display:flex;gap:20px;flex-wrap:wrap;">
      <div class="stat-card c2" style="flex:1;min-width:140px;padding:20px;"><div class="stat-number" style="font-size:32px;">${data.inserted}</div><div class="stat-label">Inserted</div></div>
      <div class="stat-card c3" style="flex:1;min-width:140px;padding:20px;"><div class="stat-number" style="font-size:32px;">${data.skipped}</div><div class="stat-label">Skipped</div></div>
      <div class="stat-card c1" style="flex:1;min-width:140px;padding:20px;"><div class="stat-number" style="font-size:32px;">${data.total}</div><div class="stat-label">Extracted</div></div>
    </div>
    <p style="margin-top:16px;font-size:12px;color:var(--muted);">Subject: <strong>${subject}</strong> — ${new Date().toLocaleTimeString()}</p>`;
  toast(data.msg,'success');
  loadStats();
}

function downloadExcel() { window.location.href='/download'; toast('Downloading Excel file...','info'); }

function onFileChange(input, fnameId) {
  const el = document.getElementById(fnameId);
  el.textContent = input.files[0] ? '📎 ' + input.files[0].name : '';
  el.style.display = input.files[0] ? 'block' : 'none';
}

function simBar(sim) {
  const pct = Math.round(sim);
  const clr = pct>=90?'#00c9a7':pct>=70?'#5b4cff':'#f5a623';
  return `<div class="sim-bar-wrap"><div class="sim-bar"><div class="sim-bar-fill" style="width:${pct}%;background:${clr};"></div></div><span class="sim-pct">${pct}%</span></div>`;
}

function highlight(text, query) {
  if (!query) return text;
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  let result = text;
  words.forEach(w => { result = result.replace(new RegExp('('+w.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi'),'<mark>$1</mark>'); });
  return result;
}

function toast(msg, type='success') {
  const icons={success:'✓',error:'✕',info:'ℹ'};
  const wrap=document.getElementById('toast-wrap');
  const el=document.createElement('div');
  el.className=`toast toast-${type}`;
  el.innerHTML=`<span>${icons[type]||'●'}</span>${msg}`;
  wrap.appendChild(el);
  setTimeout(()=>el.remove(),3500);
}

['check-zone','insert-zone'].forEach(id => {
  const zone = document.getElementById(id);
  if (!zone) return;
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    const input = zone.querySelector('input[type=file]');
    if (input && e.dataTransfer.files.length) {
      const dt = new DataTransfer();
      dt.items.add(e.dataTransfer.files[0]);
      input.files = dt.files;
      const m = input.getAttribute('onchange')?.match(/'([^']+)'/);
      if (m) onFileChange(input, m[1]);
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
        if item["exam"]:   unique_exams.add(item["exam"])
        if item["year"]:   years.add(item["year"])
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
                    "msg":f"{inserted} questions inserted, {skipped} duplicates skipped"})

@app.route("/download")
def download():
    return send_file(EXCEL_FILE, as_attachment=True)

# ─────────────────────────────────────────────
# RUN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import os
    from pyngrok import ngrok
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        public_url = ngrok.connect(5000)
        print("\n✅  ExamInsight is running!")
        print(f"👉  Local  → http://127.0.0.1:5000")
        print(f"🌍  Public → {public_url}\n")
    app.run(debug=True)
