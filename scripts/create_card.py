from pathlib import Path
import re
import sys
from datetime import date

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "product_card_template.md"
CARDS_DIR = ROOT / "cards"


def slugify(text: str) -> str:
    # Простая транслитерация для частых русских букв.
    table = {
        "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
        "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
        "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
        "ф": "f", "х": "h", "ц": "c", "ч": "ch", "ш": "sh", "щ": "sch",
        "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
    }
    value = "".join(table.get(ch.lower(), ch.lower()) for ch in text)
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "new_product"


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python scripts/create_card.py \"Название продукта\"")
        return 1

    name = " ".join(sys.argv[1:]).strip()
    slug = slugify(name)
    path = CARDS_DIR / f"{slug}.md"

    if path.exists():
        print(f"ERROR: card already exists: {path}")
        return 1

    content = TEMPLATE.read_text(encoding="utf-8")
    content = content.replace('id: ""', f'id: "prd_{slug}"')
    content = content.replace('name: ""', f'name: "{name}"')
    content = content.replace('keyword: ""', f'keyword: "{name}"')
    content = content.replace('updated_at: ""', f'updated_at: "{date.today().isoformat()}"')
    content = content.replace("# <Наименование продукта>", f"# {name}")

    path.write_text(content, encoding="utf-8")
    print(f"Created: {path.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
