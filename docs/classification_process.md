# Deterministic Classification Process

Этот процесс нужен для повторяемой классификации материалов по карточкам классов без участия модели на этапе принятия решения.

## Что делает скрипт

Скрипт `scripts/classify_materials.py`:

1. читает Excel-классификатор;
2. загружает карточки классов и правила скоринга;
3. читает список материалов из CSV/XLSX;
4. считает Python-score по каждому классу;
5. возвращает top-N кандидатов и решение:
   - `auto_accept` — класс можно принимать автоматически;
   - `review` — нужна ручная проверка;
   - `low_confidence` — низкая уверенность;
   - `no_class` — подходящий класс не найден.

## Ожидаемые листы классификатора

Обязательные:

- `Class_Cards`
- `Terms`
- `Attributes`
- `Stop_Terms`

Дополнительные листы могут быть в файле, но текущая версия скрипта использует в основном перечисленные выше.

## Минимальные колонки

### Class_Cards

- `class_code`
- `class_name`
- `class_path`
- `main_product_name`
- `search_text`

### Terms

- `class_code`
- `class_name`
- `term`
- `term_type`
- `match_type`
- `weight`
- `is_required`
- `comment`

### Attributes

- `class_code`
- `class_name`
- `attribute_name`
- `aliases`
- `unit`
- `is_required`
- `weight`
- `value_pattern`
- `normalization_rule`
- `comment`

### Stop_Terms

- `class_code`
- `class_name`
- `stop_term`
- `match_type`
- `penalty`
- `comment`

## Входной файл материалов

Поддерживаются `.csv` и `.xlsx`.

Скрипт ищет колонку с наименованием по одному из названий:

- `material_name`
- `name`
- `clean_name`
- `наименование`
- `материал`
- `номенклатура`

Колонка ID ищется по названиям:

- `material_id`
- `id`
- `код`
- `номер`
- `n`

Если ID не найден, будет использован порядковый номер строки.

## Как запустить

```bash
pip install -r requirements.txt

python scripts/classify_materials.py \
  --classifier classifier.xlsx \
  --input examples/materials_input.csv \
  --output output/classification_result.xlsx
```

## Параметры

```bash
--top-k 3
```

Сколько кандидатов выводить.

```bash
--accept-threshold 0.55
```

Порог для автоматического принятия класса.

```bash
--review-threshold 0.35
```

Порог для ручной проверки.

```bash
--min-gap 0.10
```

Минимальный разрыв между первым и вторым кандидатом. Если разрыв меньше, случай считается спорным.

## Логика скоринга

Скрипт учитывает:

1. совпадения по `Terms`;
2. обязательные термины;
3. совпадения по атрибутам и regex-паттернам из `Attributes`;
4. антипризнаки из `Stop_Terms`;
5. запасной сигнал по `class_name` и `main_product_name`.

Итоговый score нормируется в диапазон от `0` до `1`.

## Важный принцип

Модель может помогать заполнять карточки, термины, атрибуты и стоп-слова, но финальная классификация этим скриптом выполняется детерминированно.
Одинаковый входной файл + одинаковый классификатор = одинаковый результат.
