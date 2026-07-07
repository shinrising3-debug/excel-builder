# 📊 Excel Builder

**Describe any spreadsheet in plain English → get a fully formatted Excel file.**

A standalone Python CLI tool that generates professional Excel spreadsheets from natural language descriptions using Claude AI, or step-by-step without any API.

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac%20%7C%20Linux-lightgrey)

---

## ✨ What It Does

**AI Mode** — Type something like:

> *"Employee attendance tracker with name, date, clock in, clock out, total hours worked, overtime, and a status dropdown with Present, Absent, Late, Half Day"*

...and get a fully formatted `.xlsx` file with headers, dropdowns, borders, conditional formatting, alternating rows — the whole thing. Ready to use.

**Manual Mode** — No API key? No internet? Build any spreadsheet step by step:
1. Name your columns
2. Pick types (Text, Number, Currency, Date, Dropdown, Long Text)
3. Set dropdown options
4. Choose how many rows to pre-format
5. Done.

---

## 🚀 Quick Start

### Windows
1. Download `excel_builder.py` and `Excel_Builder.bat`
2. Put both in the same folder
3. Double-click `Excel_Builder.bat`

### Mac / Linux
```bash
python excel_builder.py
```

**First run** auto-installs dependencies (`openpyxl`, `anthropic`).

---

## 🔑 API Key Setup (for AI Mode)

AI Mode uses Claude (Anthropic) to generate spreadsheets. You need an API key from [console.anthropic.com](https://console.anthropic.com).

**Option A:** Run the tool → pick `[3] Set API key` → paste your key. Saved locally to `.api_key`.

**Option B:** Set environment variable:
```bash
export ANTHROPIC_API_KEY=sk-ant-...
```

Manual Mode works without any API key.

---

## 📋 Features

| Feature | AI Mode | Manual Mode |
|---|---|---|
| Natural language input | ✅ | ❌ |
| Column-by-column builder | ❌ | ✅ |
| Professional formatting | ✅ | ✅ |
| Dropdowns / data validation | ✅ | ✅ |
| Conditional formatting | ✅ | ❌ |
| Date / currency formatting | ✅ | ✅ |
| Alternating row colors | ✅ | ✅ |
| Frozen headers | ✅ | ✅ |
| Pre-formatted empty rows | ✅ | ✅ |
| Requires internet | ✅ | ❌ |
| Requires API key | ✅ | ❌ |

---

## 🎨 Output Styling

Every generated file includes:
- Dark navy headers (`#1B2A4A`) with gold accent underline
- Alternating row shading (white / off-white)
- Thin cell borders throughout
- Aptos font, black body text
- Frozen header row
- Landscape print layout, fit to page
- Gridlines hidden for clean look

---

## 💡 Example Prompts (AI Mode)

```
Monthly budget tracker with income categories, expense categories,
amounts, running balance, and a type dropdown (Fixed, Variable, One-time)
```

```
Equipment maintenance log with asset ID, equipment name, location,
last service date, next due date, technician, cost, and status
dropdown (Scheduled, In Progress, Completed, Overdue)
```

```
Hiring pipeline tracker with candidate name, position, applied date,
resume link, interview stage dropdown (Applied, Phone Screen,
Technical, Onsite, Offer, Hired, Rejected), interviewer, and notes
```

```
Inventory tracker for a warehouse — SKU, item name, category,
quantity on hand, reorder point, supplier, unit cost, total value
formula, and a status column that shows Low Stock / In Stock / Overstocked
```

---

## 🛠 Requirements

- Python 3.8+
- `openpyxl` (auto-installed)
- `anthropic` (auto-installed, only needed for AI Mode)

---

## 📁 File Structure

```
excel_builder.py          # Main script
Excel_Builder.bat         # Windows launcher
.api_key                  # Your API key (auto-created, gitignored)
_last_generated_code.py   # Debug: last AI-generated code
```

---

## 📝 License

MIT — use it however you want.

---

Built with [Claude](https://claude.ai) by [@shinrising3-debug](https://github.com/shinrising3-debug)
