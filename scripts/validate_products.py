from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[1]
DATA_FILE = ROOT / "data" / "products.jsonl"

REQUIRED_FIELDS = {
    "id",
    "status",
    "name",
    "keyword",
    "category",
    "card_path",
    "mandatory_attributes",
    "optional_attributes",
    "synonyms",
    "updated_at",
}

ALLOWED_STATUSES = {"draft", "review", "approved", "deprecated"}


def main() -> int:
    if not DATA_FILE.exists():
        print(f"ERROR: file not found: {DATA_FILE}")
        return 1

    errors = []
    seen_ids = set()

    with DATA_FILE.open("r", encoding="utf-8") as f:
        for line_no, line in enumerate(f, start=1):
            line = line.strip()
            if not line:
                continue

            try:
                item = json.loads(line)
            except json.JSONDecodeError as exc:
                errors.append(f"line {line_no}: invalid JSON: {exc}")
                continue

            missing = REQUIRED_FIELDS - set(item)
            if missing:
                errors.append(f"line {line_no}: missing fields: {sorted(missing)}")

            extra = set(item) - REQUIRED_FIELDS
            if extra:
                errors.append(f"line {line_no}: extra fields: {sorted(extra)}")

            item_id = item.get("id")
            if item_id in seen_ids:
                errors.append(f"line {line_no}: duplicate id: {item_id}")
            seen_ids.add(item_id)

            if item.get("status") not in ALLOWED_STATUSES:
                errors.append(f"line {line_no}: invalid status: {item.get('status')}")

            card_path = item.get("card_path", "")
            full_card_path = ROOT / card_path
            if not full_card_path.exists():
                errors.append(f"line {line_no}: card file not found: {card_path}")

            for list_field in ["mandatory_attributes", "optional_attributes", "synonyms"]:
                if not isinstance(item.get(list_field), list):
                    errors.append(f"line {line_no}: field must be a list: {list_field}")

    if errors:
        print("Validation failed:")
        for err in errors:
            print(f"- {err}")
        return 1

    print("Validation passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
