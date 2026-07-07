"""
AI Excel Builder — powered by Groq (free, fast, no credit card).
Uses Llama 3.3 70B via Groq's OpenAI-compatible API.
"""

import os, sys, re, traceback
from openai import OpenAI

SYSTEM_PROMPT = """You are an Excel spreadsheet generator. The user will describe a spreadsheet.

Respond with ONLY valid Python code using openpyxl. No explanation, no markdown fences, no backticks. JUST raw Python.

RULES:
1. Save to the path in variable OUTPUT_PATH (pre-defined, do not redefine it)
2. Import everything you need at the top
3. Professional formatting:
   - ws.sheet_view.showGridLines = False
   - Font: Aptos throughout, body text color 212529 (black)
   - Header row: background 1B2A4A, white bold text, gold bottom border D4AF37 medium
   - Cell borders: thin, color CED4DA on ALL data cells
   - Alternating rows: white FFFFFF / off-white F8F9FA
   - Row 1: 4px gold accent bar (fill entire row with D4AF37)
   - Row 2: merged title, font size 16, bold, color 1B2A4A
   - Row 3: merged subtitle, font size 10, color 6C757D
   - Freeze the header row
   - Set proper column widths
4. Add data validation dropdowns (showDropDown=False to show arrow)
5. Conditional formatting: BACKGROUND only, no font color change, entire ROW
   - PatternFill(start_color="FF...", end_color="FF...", fill_type="solid")
   - FormulaRule with stopIfTrue=True
6. Pre-format 200 empty rows with borders, fonts, alignment, alternating fill
7. Add a # column as column A
8. Date columns: number_format = "mm/dd/yyyy"
9. Currency columns: number_format = '$#,##0.00'
10. NEVER use ws.add_table()
11. Sanitize sheet title: remove characters / \\ : ? * [ ]
12. Landscape orientation, fit to page width

OUTPUT_PATH is pre-defined before your code runs. Do not redefine it."""


def main():
    description = os.environ.get("DESCRIPTION", "").strip()
    filename = os.environ.get("FILE_NAME", "My_Spreadsheet").strip()
    api_key = os.environ.get("GROQ_API_KEY", "").strip()

    if not description:
        print("❌ No description provided.")
        sys.exit(1)

    if not api_key:
        print("❌ No Groq API key. Add GROQ_API_KEY as a repository secret.")
        print("   Get a free key at: https://console.groq.com (no credit card)")
        sys.exit(1)

    if not filename.endswith('.xlsx'):
        filename = re.sub(r'[<>:"/\\|?*]', '', filename).replace(' ', '_') + '.xlsx'

    output_path = os.path.join(os.getcwd(), filename)

    print(f"📊 Generating: {filename}")
    print(f"📝 Description: {description[:300]}")
    print(f"🤖 Using: Groq — Llama 3.3 70B (free)")
    print()

    client = OpenAI(
        api_key=api_key,
        base_url="https://api.groq.com/openai/v1"
    )

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": description}
            ],
            temperature=0.2,
            max_tokens=8000,
        )

        code = response.choices[0].message.content

        # Strip markdown fences
        code = re.sub(r'^```python\s*\n?', '', code, flags=re.MULTILINE)
        code = re.sub(r'^```\s*$', '', code, flags=re.MULTILINE)
        code = code.strip()

        # Save for debugging
        with open("_generated_code.py", "w", encoding="utf-8") as f:
            f.write(code)
        print("💾 Generated code saved to _generated_code.py\n")

        # Execute
        exec(code, {"OUTPUT_PATH": output_path, "__builtins__": __builtins__})

        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print(f"\n✅ Done — {filename} ({size:,} bytes)")
        else:
            print("❌ Code ran but no file created.")
            print("Generated code:")
            print(code[:2000])
            sys.exit(1)

    except Exception as e:
        print(f"❌ Error: {e}")
        traceback.print_exc()
        if os.path.exists("_generated_code.py"):
            print("\nGenerated code saved to _generated_code.py")
        sys.exit(1)


if __name__ == "__main__":
    main()
