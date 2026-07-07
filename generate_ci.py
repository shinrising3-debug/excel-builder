"""
CI/CD Excel Generator — called by GitHub Actions.
Reads DESCRIPTION and FILE_NAME from environment variables.
Outputs a formatted .xlsx in the current directory.
"""

import os, sys, re, traceback

import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

SYSTEM_PROMPT = """You are an Excel spreadsheet generator. The user will describe a spreadsheet they need.

You must respond with ONLY valid Python code that uses openpyxl to create the spreadsheet.
No explanation, no markdown, no ```python blocks — JUST the raw Python code.

RULES:
1. Always save to the path stored in variable OUTPUT_PATH (it will be defined before your code runs)
2. Import everything you need at the top
3. Make it look PROFESSIONAL and PRESENTABLE:
   - Turn off gridlines: ws.sheet_view.showGridLines = False
   - Use 'Aptos' font throughout
   - Dark header row (background: 1B2A4A, font: white, bold)
   - Gold accent line under headers (bottom border: D4AF37, medium)
   - Cell borders on ALL data cells (thin, color CED4DA)
   - Alternating row colors (white FFFFFF and off-white F8F9FA)
   - Proper column widths (auto-size based on content type)
   - Freeze the header row
   - All body text must be BLACK (color: 212529)
   - Wrap text and center-align short fields, left-align long text
   - Row heights: headers 30, data rows 28
   - Add a title row at top (merged, font size 16, bold)
   - Add a subtle subtitle row below title
   - Add a 4px gold accent bar as row 1
   - Data starts after the header row
4. Add data validation dropdowns where appropriate (showDropDown=False to show the arrow)
5. Add conditional formatting with BACKGROUND color only (no font color change) where it makes sense
   - Use PatternFill with start_color and end_color with FF prefix for conditional formatting
   - Use FormulaRule with stopIfTrue=True
   - Apply conditional formatting to the ENTIRE row based on status/category column
6. Pre-format at least 200 empty rows with borders, fonts, alignment, and alternating fill
7. Set landscape orientation and fit to page width
8. Include any sample data the user mentioned
9. If the user mentions specific columns, use EXACTLY those column names
10. For date columns use number_format = "mm/dd/yyyy"
11. For currency columns use number_format = '$#,##0.00'
12. NEVER use Excel Tables (ws.add_table) — they cause rendering issues
13. Add a # column as the first column for sequential entry numbering

The output file path will be: OUTPUT_PATH (pre-defined variable, do not redefine it)
Your code will be exec'd in a Python environment with openpyxl already imported.
OUTPUT_PATH will already be defined as a variable before your code runs.
"""


def sanitize_filename(name):
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = name.strip().replace(' ', '_')
    if not name.endswith('.xlsx'):
        name += '.xlsx'
    return name


def main():
    description = os.environ.get("DESCRIPTION", "").strip()
    file_name = sanitize_filename(os.environ.get("FILE_NAME", "My_Spreadsheet"))
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()

    if not description:
        print("❌ No description provided. Set DESCRIPTION env var.")
        sys.exit(1)

    if not api_key:
        print("❌ No API key found. Add ANTHROPIC_API_KEY as a repository secret.")
        sys.exit(1)

    output_path = os.path.join(os.getcwd(), file_name)
    print(f"📊 Generating: {file_name}")
    print(f"📝 Description: {description[:200]}...")
    print()

    client = anthropic.Anthropic(api_key=api_key)

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8000,
            system=SYSTEM_PROMPT,
            messages=[
                {"role": "user", "content": description}
            ]
        )

        code = response.content[0].text

        # Strip markdown fences if present
        code = re.sub(r'^```python\s*\n?', '', code, flags=re.MULTILINE)
        code = re.sub(r'^```\s*$', '', code, flags=re.MULTILINE)
        code = code.strip()

        # Save generated code for debugging
        with open("_generated_code.py", "w", encoding="utf-8") as f:
            f.write(code)
        print("💾 Generated code saved to _generated_code.py")

        # Execute
        exec_globals = {
            "OUTPUT_PATH": output_path,
            "__builtins__": __builtins__,
        }
        exec(code, exec_globals)

        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print()
            print(f"✅ SUCCESS — {file_name} ({size:,} bytes)")
        else:
            print("❌ Code ran but no file was created.")
            sys.exit(1)

    except Exception as e:
        print(f"❌ Error: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
