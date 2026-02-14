import os
import sys
import json
import pandas as pd


def clean_record(record: dict) -> dict:
    cleaned = {}

    for key, value in record.items():
        # Skip NaN / None
        if pd.isna(value):
            continue

        # Clean strings
        if isinstance(value, str):
            value = value.strip()
            if value == "":
                continue

        cleaned[key] = value

    return cleaned


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    xlsx_path = os.path.join(base_dir, "brawl_vault_full_export.xlsx")
    json_path = os.path.join(base_dir, "brawl_data.json")

    if not os.path.exists(xlsx_path):
        print(f"ERROR: 'brawl_vault_full_export.xlsx' not found in:\n{base_dir}")
        sys.exit(1)

    try:
        df = pd.read_excel(xlsx_path, engine="openpyxl")

        records = df.to_dict(orient="records")

        cleaned_records = [clean_record(record) for record in records]

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(cleaned_records, f, indent=2, ensure_ascii=False)

        print(f"Successfully created:\n{json_path}")

    except Exception as e:
        print("An error occurred while converting the file:")
        print(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
