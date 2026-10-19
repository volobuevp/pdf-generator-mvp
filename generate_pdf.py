# -*- coding: utf-8 -*-
"""
Генератор PDF: заявка на перемещение ТМЦ.
Читает данные (CSV/JSON) из /data, подставляет в HTML-шаблон из /templates,
сохраняет PDF в /output и открывает его в системной программе.

Автоматически присваивает номер заявки и дату отгрузки (день генерации).
"""

import csv
import json
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

from xhtml2pdf import pisa

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

COMPANY = "ООО «Союзинструмент»"  # юрлицо по умолчанию (пример)

# Где ищем шрифт с кириллицей (Windows / macOS / свой файл в fonts/)
FONT_CANDIDATES = [
    BASE_DIR / "fonts" / "DejaVuSans.ttf",
    Path("C:/Windows/Fonts/arial.ttf"),
    Path("C:/Windows/Fonts/segoeui.ttf"),
    Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
    Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
]


def find_font():
    for candidate in FONT_CANDIDATES:
        if candidate.exists():
            return candidate.as_posix()
    sys.exit("Ошибка: не найден шрифт с кириллицей. Положите DejaVuSans.ttf в папку fonts/.")


def read_text(path):
    """Читает файл с автодетектом кодировки (Excel сохраняет CSV в CP1251)."""
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    sys.exit(f"Ошибка: не удалось распознать кодировку файла {path.name}.")


def parse_csv(path):
    text = read_text(path)
    lines = [ln for ln in text.splitlines() if ln.strip()]
    delim = ";" if lines and ";" in lines[0] else ","
    rows = list(csv.reader(lines, delimiter=delim))

    meta, items = {}, []
    header_idx = None
    for i, row in enumerate(rows):
        cells = [c.strip() for c in row]
        if any("товар" in c.lower() for c in cells if c):
            header_idx = i
            break
        if len(cells) >= 2 and cells[0]:
            meta[cells[0].lower()] = cells[1]

    if header_idx is None:
        sys.exit(f"Ошибка: в {path.name} не найдена строка заголовка с колонкой «товар».")

    for row in rows[header_idx + 1:]:
        cells = [c.strip() for c in row]
        if len(cells) < 3 or not cells[0]:
            continue
        try:
            qty, price = float(cells[1].replace(",", ".")), float(cells[2].replace(",", "."))
        except ValueError:
            continue
        items.append({"product": cells[0], "qty": qty, "price": price})

    if not items:
        sys.exit(f"Ошибка: в {path.name} не найдено ни одной строки с товаром.")
    return meta, items


def parse_json(path):
    data = json.loads(read_text(path))
    items_raw = data.get("товары") or data.get("items") or []
    items = [
        {
            "product": str(it.get("товар") or it.get("product")),
            "qty": float(it.get("количество") or it.get("qty")),
            "price": float(it.get("цена за шт") or it.get("цена") or it.get("price")),
        }
        for it in items_raw
    ]
    if not items:
        sys.exit(f"Ошибка: в {path.name} нет списка товаров (ключ «товары»/«items»).")
    meta = {
        k.lower(): v
        for k, v in data.items()
        if isinstance(v, str) and k not in ("товары", "items")
    }
    return meta, items


def load_data(path):
    if path.suffix.lower() == ".json":
        return parse_json(path)
    return parse_csv(path)


def next_doc_number():
    """Номер заявки = максимум существующих в output/ + 1."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    numbers = [
        int(p.stem.split("_")[1])
        for p in OUTPUT_DIR.glob("Перемещение_*_*.pdf")
        if len(p.stem.split("_")) == 3 and p.stem.split("_")[1].isdigit()
    ]
    return max(numbers, default=0) + 1


def choose(options, title):
    print(f"\n{title}:")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    while True:
        answer = input("Выберите пункт (номер): ").strip()
        if answer.isdigit() and 1 <= int(answer) <= len(options):
            return options[int(answer) - 1]
        print("Некорректный ввод, попробуйте ещё раз.")


def render_html(template_path, meta, items, doc_number, font_path):
    html = read_text(template_path)
    today = date.today().strftime("%d.%m.%Y")
    receipt_date = meta.get("дата приёмки") or meta.get("дата_приёмки") or ""

    total = sum(it["qty"] * it["price"] for it in items)
    fmt = lambda n: f"{n:,.2f}".replace(",", " ").replace(".", ",")

    rows_html = "".join(
        f"<tr><td>{it['product']}</td><td>{it['qty']:.10g}</td>"
        f"<td>{fmt(it['price'])}</td><td>{fmt(it['qty'] * it['price'])}</td></tr>"
        for it in items
    )

    placeholders = {
        "font": font_path,
        "company": COMPANY,
        "number": f"{doc_number:03d}",
        "ship_date": today,
        "receipt_date": receipt_date,
        "rows": rows_html,
        "total": fmt(total),
    }
    for key, value in placeholders.items():
        html = html.replace("{{ " + key + " }}", str(value)).replace("{{" + key + "}}", str(value))
    return html, total


def save_pdf(html, pdf_path):
    with open(pdf_path, "wb") as result:
        status = pisa.CreatePDF(html, dest=result, encoding="utf-8")
    if status.err:
        sys.exit(f"Ошибка при генерации PDF (ошибок: {status.err}).")


def open_pdf(pdf_path):
    if sys.platform == "win32":
        os.startfile(pdf_path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", pdf_path], check=True)
    else:
        subprocess.run(["xdg-open", pdf_path], check=True)


def main():
    if not DATA_DIR.exists():
        sys.exit("Нет папки data/ — положите туда файлы с данными (CSV или JSON).")

    data_files = sorted([*DATA_DIR.glob("*.csv"), *DATA_DIR.glob("*.json")])
    templates = sorted(TEMPLATES_DIR.glob("*.html"))
    if not data_files:
        sys.exit("В папке data/ нет CSV/JSON файлов.")
    if not templates:
        sys.exit("В папке templates/ нет HTML-шаблонов.")

    print("=" * 50)
    print("Генератор PDF — заявки на перемещение ТМЦ")
    print("=" * 50)

    data_path = choose(data_files, "Доступные файлы с данными")
    template_path = choose(templates, "Доступные шаблоны")

    meta, items = load_data(data_path)
    doc_number = next_doc_number()

    html, total = render_html(template_path, meta, items, doc_number, find_font())

    OUTPUT_DIR.mkdir(exist_ok=True)
    pdf_path = OUTPUT_DIR / f"Перемещение_{doc_number:03d}_{date.today():%Y-%m-%d}.pdf"
    save_pdf(html, pdf_path)

    print(f"\nГотово: заявка № {doc_number:03d}")
    print(f"Товарных позиций: {len(items)}, сумма итого: {total:,.2f}".replace(",", " "))
    print(f"Файл: {pdf_path}")
    open_pdf(pdf_path)


if __name__ == "__main__":
    main()