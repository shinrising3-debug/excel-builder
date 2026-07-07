# 📊 Excel Builder

**Build any formatted Excel file directly from GitHub. No install, no coding.**

Two ways to use — pick whichever fits:

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Cost](https://img.shields.io/badge/Cost-Free-brightgreen)

---

## 📋 Option 1: Simple Form (No AI)

A step-by-step form that asks you exactly what you need:

1. Go to **[Actions](../../actions)** → click **"📊 Build Excel (Simple Form)"**
2. Click **"Run workflow"** and fill in:

| Field | What to enter | Example |
|---|---|---|
| **Title** | Name of your spreadsheet | `Gate / Barrier Outage Tracker` |
| **Status filter?** | Pick Yes or No | `Yes` |
| **Status options** | Your statuses separated by `\|` | `Pending \| Escalated \| Fixed \| Manual In/Out` |
| **Columns** | All your columns, comma-separated | `Door / Barrier, Time Occurred, Identified By, Notes` |
| **Column types** | Type per column (same order) | `text, text, text, long` |
| **Rows** | How many blank rows | `200` |
| **File name** | Output file name | `Outage_Tracker` |

3. Click **Run workflow** → wait 20 sec → download `.xlsx` from **Artifacts**

### Column Types

| Type | What it does |
|---|---|
| `text` | Standard column (default if left blank) |
| `date` | Formatted as `mm/dd/yyyy` |
| `number` | Center-aligned number |
| `currency` | Formatted as `$#,##0.00` |
| `long` | Extra-wide column for notes/details |

**Status filter** adds a dropdown as the last column. When you pick a status, the **entire row** changes background color automatically.

---

## 🤖 Option 2: AI Mode (Free — Gemini)

Just describe what you want in plain English:

1. Go to **[Actions](../../actions)** → click **"🤖 Build Excel (AI — Free)"**
2. Click **"Run workflow"**
3. Type your description, for example:

> Employee attendance tracker with name, department, date, clock in time, clock out time, total hours, overtime hours, and a status dropdown with Present, Absent, Late, Half Day. Include sample data for 3 employees.

4. Click **Run workflow** → wait 30 sec → download from **Artifacts**

Powered by Google Gemini 2.0 Flash (free tier — no cost, no limits for normal use).

---

## 🎨 What Every File Gets

Regardless of which option you use, every generated file includes:

- Dark navy headers with gold accent underline
- Alternating row shading (white / off-white)
- Thin cell borders on every cell
- Aptos font, black body text
- Sequential `#` column
- Frozen header row
- Dropdown arrows on status/choice columns
- Full-row conditional coloring based on status
- Landscape print layout, fit to page
- Gridlines hidden for clean look

---

## 💻 Run Locally (Optional)

Download `excel_builder.py` + `Excel_Builder.bat` → double-click the `.bat`:

- **[1] AI Mode** — describe it, Claude builds it (needs Anthropic API key)
- **[2] Manual Mode** — step-by-step column builder (no API needed)

---

## 📁 Files

| File | Purpose |
|---|---|
| `build_simple.py` | Simple form builder (GitHub Actions) |
| `build_ai.py` | Gemini AI builder (GitHub Actions) |
| `generate_ci.py` | Advanced column-syntax builder |
| `excel_builder.py` | Local CLI tool (both modes) |
| `Excel_Builder.bat` | Windows launcher for local tool |

---

## 🔑 Setup (Repo Owner Only)

The **Simple Form** workflow needs zero secrets — it's pure Python.

The **AI workflow** needs a free Gemini key:
1. Get a key at [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. Go to repo **Settings → Secrets → Actions → New secret**
3. Name: `GEMINI_API_KEY` — paste your key

---

MIT License · Built by [@shinrising3-debug](https://github.com/shinrising3-debug)
