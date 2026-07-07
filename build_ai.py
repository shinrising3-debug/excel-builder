"""
AI Excel Builder — powered by Google Gemini (free tier).
Describe what you need in plain English, Gemini generates the code, we run it.
"""

import os, sys, re, traceback

SYSTEM_PROMPT = """You are an Excel spreadsheet generator. The user will describe a spreadsheet.

Respond with ONLY valid Python code using openpyxl. No explanation, no markdown fences, no backticks.

RULES:
1. Save to the path in variable OUTPUT_PATH (pre-defined, do not redefine)
2. Import everything you need at the top
3. Professional formatting:
   - ws.sheet_view.showGridLines = False
   - Font: Aptos throughout, body text color 212529 (black)
   - Header row: background 1B2A4A, white bold text, gold bottom border (D4AF37, medium)
   - Cell borders: thin, color CED4DA on ALL data cells
   - Alternating rows: white FFFFFF / off-white F8F9FA
   - Row 1: 4px gold accent bar (fill entire row with D4AF37)
   - Row 2: merged title (font size 16, bold, color 1B2A4A)
   - Row 3: merged subtitle (font size 10, color 6C757D)
   - Freeze the header row
   - Auto column widths based on content
4. Add data validation dropdowns where useful (showDropDown=False)
5. Conditional formatting: BACKGROUND only (no font color changes), apply to ENTIRE ROW
   - Use PatternFill(start_color="FF...", end_color="FF...", fill_type="solid")
   - Use FormulaRule with stopIfTrue=True
6. Pre-format 200+ empty rows with borders, fonts, alignment, alternating fill
7. Add a # column as column A
8. Date columns: number_format = "mm/dd/yyyy"
9. Currency columns: number_format = '$#,##0.00'
10. NEVER use ws.add_table() — it breaks rendering
11. Landscape, fit to page width

OUTPUT_PATH is pre-defined. Just write the code."""


def main():
    description = os.environ.get("DESCRIPTION", "").strip()
    filename = os.environ.get("FILE_NAME", "My_Spreadsheet").strip()
    api_key = os.environ.get("GEMINI_API_KEY", "").strip()

    if not description:
        print("❌ No description provided.")
        sys.exit(1)

    if not api_key:
        print("❌ No Gemini API key. Add GEMINI_API_KEY as a repository secret.")
        print("   Get a free key at: https://aistudio.google.com/apikey")
        sys.exit(1)

    if not filename.endswith('.xlsx'):
        filename = re.sub(r'[<>:"/\\|?*]', '', filename).replace(' ', '_') + '.xlsx'

    output_path = os.path.join(os.getcwd(), filename)

    print(f"📊 Generating: {filename}")
    print(f"📝 Description: {description[:300]}")
    print(f"🤖 Using: Google Gemini (free tier)")
    print()

    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    try:
        response = model.generate_content(
            f"{SYSTEM_PROMPT}\n\nUser request:\n{description}",
            generation_config=genai.GenerationConfig(
                temperature=0.2,
                max_output_tokens=8000,
            )
        )

        code = response.text

        # Strip markdown fences
        code = re.sub(r'^```python\s*\n?', '', code, flags=re.MULTILINE)
        code = re.sub(r'^```\s*$', '', code, flags=re.MULTILINE)
        code = code.strip()

        # Save for debugging
        with open("_generated_code.py", "w", encoding="utf-8") as f:
            f.write(code)
        print("💾 Generated code saved to _generated_code.py")

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
            print("\nGenerated code saved to _generated_code.py for debugging")
        sys.exit(1)


if __name__ == "__main__":
    main()
