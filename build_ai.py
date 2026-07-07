"""
AI Excel Builder — Groq generates a JSON spec, our proven builder creates the file.
No code execution of AI output. Much more reliable.
"""

import os, sys, re, json, traceback
from openai import OpenAI
from generate_ci import build_excel, parse_columns

SYSTEM_PROMPT = """You convert spreadsheet descriptions into a JSON specification.

RESPOND WITH ONLY A JSON OBJECT. No markdown, no backticks, no explanation.

JSON format:
{
  "title": "Sheet Title Here",
  "columns": "Name, Date (date), Amount (currency), Status (dropdown: Option1|Option2|Option3), Notes (long)",
  "sample_data": "John Smith|2024-01-15|1500.00|Option1|First entry\\nJane Doe|2024-01-16|2300.00|Option2|Second entry"
}

Column type syntax in the "columns" string:
- Plain text (default): just the name, e.g. "Employee Name"
- Date formatted: add (date), e.g. "Hire Date (date)"
- Number: add (number), e.g. "Quantity (number)"
- Currency: add (currency), e.g. "Total Cost (currency)"
- Dropdown: add (dropdown: Opt1|Opt2|Opt3), e.g. "Status (dropdown: Pending|Fixed|Escalated)"
- Wide text column: add (long), e.g. "Notes (long)"

Sample data rules:
- Columns separated by |
- Rows separated by \\n (newline)
- Match the column order (excluding the # column which is auto-added)
- If the user didn't mention sample data, include 2-3 realistic example rows anyway

RESPOND WITH ONLY THE JSON OBJECT. Nothing else."""


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
            temperature=0.1,
            max_tokens=4000,
        )

        raw = response.choices[0].message.content.strip()

        # Strip markdown fences if present
        raw = re.sub(r'^```(?:json)?\s*\n?', '', raw, flags=re.MULTILINE)
        raw = re.sub(r'^```\s*$', '', raw, flags=re.MULTILINE)
        raw = raw.strip()

        print(f"📄 AI response:\n{raw[:500]}\n")

        spec = json.loads(raw)
        title = spec.get("title", "My Spreadsheet")
        columns_raw = spec.get("columns", "")
        sample_data = spec.get("sample_data", "")

        if not columns_raw:
            print("❌ AI returned no columns.")
            sys.exit(1)

        columns = parse_columns(columns_raw)
        if not columns:
            print("❌ Could not parse columns from AI response.")
            sys.exit(1)

        print(f"📋 Title: {title}")
        print(f"📐 Columns ({len(columns)}):")
        for col in columns:
            opts = f" → {', '.join(col['options'])}" if col.get('options') else ""
            print(f"   • {col['name']} ({col['type']}){opts}")
        print()

        build_excel(title, columns, 200, output_path, sample_data or None)

        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print(f"\n✅ Done — {filename} ({size:,} bytes)")
        else:
            print("❌ No file created.")
            sys.exit(1)

    except json.JSONDecodeError as e:
        print(f"❌ AI didn't return valid JSON: {e}")
        print(f"Raw response:\n{raw[:1000]}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
