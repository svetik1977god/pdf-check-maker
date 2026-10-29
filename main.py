#!/usr/bin/env python3
"""PDF receipt maker.

Console tool that reads CSV / JSON data from ./data, HTML templates from
./templates, asks the user to pick a data file, a template and a receipt
(by invoice id), fills the template with the data and writes a PDF into
./output with WeasyPrint. The finished PDF is opened in the system viewer.

Works on Windows and macOS. Cyrillic is covered by the bundled DejaVu Sans
(see ./fonts); any other .ttf/.otf dropped into ./fonts is picked up too.
"""

from __future__ import annotations

import csv
import html as html_lib
import io
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
FONTS_DIR = BASE_DIR / "fonts"

MENU_WIDTH = 66

# --------------------------------------------------------------------------
# Native libraries (WeasyPrint needs Pango / GObject at runtime)
# --------------------------------------------------------------------------


def register_windows_dll_dirs() -> None:
    """Expose the GTK runtime folder so ctypes/cffi can load Pango DLLs.

    Python 3.8+ no longer searches PATH for extension dependencies, so the
    GTK directory has to be registered explicitly before importing WeasyPrint.
    """
    if sys.platform != "win32" or not hasattr(os, "add_dll_directory"):
        return

    candidates: list[Path] = []
    gtk_basedir = os.environ.get("GTK_BASEDIR", "")
    if gtk_basedir:
        candidates.append(Path(gtk_basedir) / "bin")
    candidates += [
        Path(r"C:\Program Files\GTK3-Runtime Win64\bin"),
        Path(r"C:\Program Files (x86)\GTK3-Runtime Win64\bin"),
        Path(r"C:\gtk\bin"),
    ]
    candidates += [Path(item) for item in os.environ.get("PATH", "").split(os.pathsep) if item]

    seen: set[str] = set()
    for directory in candidates:
        try:
            resolved = str(directory)
        except OSError:
            continue
        if resolved.lower() in seen:
            continue
        seen.add(resolved.lower())
        if (Path(resolved) / "libgobject-2.0-0.dll").exists():
            try:
                os.add_dll_directory(resolved)
            except OSError:
                pass


register_windows_dll_dirs()

try:
    from weasyprint import HTML  # type: ignore
except Exception as exc:  # ImportError or OSError when Pango is missing
    HTML = None  # type: ignore[assignment]
    WEASYPRINT_ERROR: Exception | None = exc
else:
    WEASYPRINT_ERROR = None

DEPENDENCY_HELP = """
  WeasyPrint не установлен корректно.

  1) pip install -r requirements.txt

  2) Нативная часть (Pango/GTK) обязательна:
     • Windows : установите GTK3 Runtime — https://github.com/tschoonj/
                 GTK-for-Windows-Runtime-Environment-Installer
                 (можно задать путь через переменную GTK_BASEDIR)
     • macOS   : brew install pango
     • Linux   : sudo apt install libpango-1.0-0 libpangocairo-1.0-0 \
                       libcairo2 fonts-dejavu-core

  Текущая ошибка: {error}
"""

# --------------------------------------------------------------------------
# Field aliases (keys are compared in normalized form)
# --------------------------------------------------------------------------


def norm_key(value: str) -> str:
    """Lower-case a header name and drop everything but letters and digits."""
    return re.sub(r"[^0-9a-z\u0400-\u04ff]+", "", str(value).lower())


# Ordered by priority: the first alias found in the data wins.
ID_KEYS = (
    "invoiceid", "invoice", "invoicenumber", "checkid", "receiptid", "nomercheka",
    "nomerdokumenta", "doknom", "number", "id",
    "чек", "номерчека", "номер", "счёт", "счет", "документ", "код",
)
DATE_KEYS = frozenset(
    {"date", "data", "checkdate", "datetime", "createdat", "created", "datacheka",
     "дата", "датачека", "число", "выдан"}
)
CUSTOMER_KEYS = frozenset(
    {"customer", "customername", "client", "buyer", "pokupatel", "klient",
     "покупатель", "клиент", "заказчик", "фио", "физлицо"}
)
SELLER_KEYS = frozenset(
    {"seller", "company", "shop", "organizaciya", "prodavets", "postavschik",
     "продавец", "компания", "организация", "поставщик", "продавцы"}
)
CURRENCY_KEYS = frozenset({"currency", "curr", "valuta", "валюта"})
ITEM_NAME_KEYS = frozenset(
    {"item", "description", "name", "product", "service", "tovar", "usluga", "naimenovanie",
     "poziciya", "opisanie", "товар", "услуга", "наименование", "позиция", "описание", "работы"}
)
ITEM_QTY_KEYS = frozenset(
    {"qty", "quantity", "count", "colvo", "kolichestvo", "amountof",
     "колво", "количество", "шт", "объём", "объем"}
)
ITEM_PRICE_KEYS = frozenset(
    {"price", "cost", "unitprice", "tsena", "rate", "цена", "тариф", "расценка", "ставка"}
)
ITEM_SUM_KEYS = frozenset(
    {"sum", "amount", "total", "summa", "itogo", "value", "стоимость", "сумма", "итого"}
)
TOTAL_KEYS = ITEM_SUM_KEYS

CURRENCY_SIGNS = {"RUB": "\u20bd", "RUR": "\u20bd", "USD": "$", "EUR": "\u20ac", "GBP": "\u00a3", "KZT": "\u20b8"}

FACE_RULE = (
    '@font-face {{ font-family: "{family}"; src: url("{src}"); '
    "font-weight: {weight}; font-style: {style}; }}"
)
WEIGHT_TOKENS = {
    "thin": "100", "hairline": "100", "extralight": "200", "ultralight": "200", "light": "300",
    "regular": "400", "normal": "400", "book": "400", "medium": "500", "semibold": "600",
    "demibold": "600", "bold": "700", "extrabold": "800", "ultrabold": "800",
    "black": "900", "heavy": "900",
}
STYLE_TOKENS = {"italic": "italic", "oblique": "oblique"}
KNOWN_FAMILIES = {
    "dejavusans": "DejaVu Sans",
    "dejavusansmono": "DejaVu Sans Mono",
    "roboto": "Roboto",
    "robotomono": "Roboto Mono",
    "arial": "Arial",
    "inter": "Inter",
    "montserrat": "Montserrat",
    "opensans": "Open Sans",
    "lato": "Lato",
}

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
UNSAFE_NAME_RE = re.compile(r"[^\w.\-]+", re.UNICODE)


# --------------------------------------------------------------------------
# Small value helpers
# --------------------------------------------------------------------------


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        if value != value:  # NaN
            return True
    except (TypeError, ValueError):
        return True  # pandas.NA and friends
    return False


def unwrap(value: Any) -> Any:
    """Turn numpy scalars into plain Python values."""
    if hasattr(value, "item") and not isinstance(value, (str, bytes, list, tuple, dict)):
        try:
            return value.item()
        except (ValueError, AttributeError):
            return value
    return value


def to_text(value: Any) -> str:
    value = unwrap(value)
    if is_blank(value):
        return ""
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}".rstrip("0").rstrip(".")
    if isinstance(value, (list, tuple)):
        return ", ".join(to_text(item) for item in value)
    if isinstance(value, dict):
        return "; ".join(f"{key}: {to_text(val)}" for key, val in value.items())
    return str(value).strip()


def to_number(value: Any) -> float | None:
    value = unwrap(value)
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if is_blank(value) else float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def money(value: float | None, currency: str = "") -> str:
    """Format a number the Russian way: 1 234,56 ₽."""
    if value is None:
        return "—"
    code = str(currency).upper().strip()
    sign = CURRENCY_SIGNS.get(code, "")
    text = f"{value:,.2f}".replace(",", "\u00a0").replace(".", ",")
    if sign:
        return f"{text}\u00a0{sign}"
    if code:
        return f"{text}\u00a0{code}"
    return text


def format_date(value: Any) -> str:
    text = to_text(value).split("T")[0].split(" ")[0]
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d", "%d-%m-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return text


def pick_key(mapping: dict[str, Any], aliases: frozenset[str]) -> str | None:
    """First key of the record whose normalized name is listed in aliases."""
    for key in mapping:
        if norm_key(key) in aliases:
            return key
    return None


def pick_value(mapping: dict[str, Any], aliases: frozenset[str]) -> Any:
    key = pick_key(mapping, aliases)
    return mapping.get(key) if key else None


def escape(value: Any) -> str:
    return html_lib.escape(to_text(value), quote=False)


def clip(text: str, width: int) -> str:
    text = text or "—"
    return text if len(text) <= width else text[: width - 1] + "…"


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------


@dataclass
class Item:
    name: str
    qty: float
    price: float | None
    amount: float | None


@dataclass
class Invoice:
    invoice_id: str
    fields: dict[str, Any]
    items: list[Item]
    total: float | None
    currency: str
    source: Path

    def field(self, aliases: frozenset[str]) -> str:
        return to_text(pick_value(self.fields, aliases))

    @property
    def date_text(self) -> str:
        return format_date(pick_value(self.fields, DATE_KEYS))

    @property
    def customer_text(self) -> str:
        return self.field(CUSTOMER_KEYS)


# --------------------------------------------------------------------------
# Reading data files
# --------------------------------------------------------------------------


def find_files(directory: Path, extensions: tuple[str, ...]) -> list[Path]:
    if not directory.is_dir():
        return []
    files = [p for p in directory.rglob("*") if p.is_file() and p.suffix.lower() in extensions]
    return sorted(files, key=lambda p: str(p.relative_to(directory)).lower())


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def detect_delimiter(text: str) -> str:
    header = text.splitlines()[0] if text.splitlines() else ""
    counts = {sep: header.count(sep) for sep in (";", ",", "\t", "|")}
    best = max(counts, key=lambda sep: counts[sep])
    return best if counts[best] else ","


def csv_records(path: Path) -> list[dict[str, Any]]:
    """Parse CSV with pandas when available, otherwise with the stdlib."""
    text = read_text(path)
    delimiter = detect_delimiter(text)
    try:
        import pandas as pd
    except ImportError:
        pass
    else:
        try:
            frame = pd.read_csv(io.StringIO(text), sep=delimiter, skipinitialspace=True)
        except Exception:
            frame = None
        if frame is not None and not frame.empty:
            return [
                {str(key).strip(): unwrap(value) for key, value in row.items()}
                for row in frame.to_dict(orient="records")
            ]
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter, skipinitialspace=True)
    return [{(key or "").strip(): value for key, value in row.items()} for row in reader]


def json_records(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(read_text(path))
    if isinstance(payload, list):
        return [entry for entry in payload if isinstance(entry, dict)]
    if isinstance(payload, dict):
        for key, value in payload.items():
            if isinstance(value, list) and any(isinstance(entry, dict) for entry in value):
                return [entry for entry in value if isinstance(entry, dict)]
        return [payload]
    raise ValueError("в файле нет записей-объектов")


def load_records(path: Path) -> list[dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        records = csv_records(path)
    elif suffix == ".json":
        records = json_records(path)
    else:
        raise ValueError(f"неподдерживаемое расширение {suffix}")
    return [record for record in records if isinstance(record, dict) and record]


def detect_id_key(records: list[dict[str, Any]]) -> str | None:
    available: dict[str, str] = {}
    for record in records[:10]:
        for key in record:
            available.setdefault(norm_key(key), key)
    for alias in ID_KEYS:
        if alias in available:
            return available[alias]
    return None


def nested_items(record: dict[str, Any]) -> list[dict[str, Any]] | None:
    """Return a list of line items embedded in a record, if there is one."""
    for key, value in record.items():
        if isinstance(value, list) and value and all(isinstance(entry, dict) for entry in value):
            if any(pick_key(entry, ITEM_NAME_KEYS | ITEM_PRICE_KEYS) for entry in value):
                return value
        if isinstance(value, str):
            text = value.strip()
            if text.startswith("[") and norm_key(key) in {"items", "tovary", "pozicii", "uslugi", "positions"}:
                try:
                    parsed = json.loads(text)
                except json.JSONDecodeError:
                    continue
                if isinstance(parsed, list) and all(isinstance(entry, dict) for entry in parsed):
                    return parsed or None
    return None


def make_item(entry: dict[str, Any]) -> Item | None:
    name_key = pick_key(entry, ITEM_NAME_KEYS)
    qty_key = pick_key(entry, ITEM_QTY_KEYS)
    price_key = pick_key(entry, ITEM_PRICE_KEYS)
    sum_key = pick_key(entry, ITEM_SUM_KEYS)
    if not (name_key or qty_key or price_key or sum_key):
        return None

    name = to_text(entry.get(name_key)) if name_key else ""
    qty = to_number(entry.get(qty_key)) if qty_key else None
    price = to_number(entry.get(price_key)) if price_key else None
    amount = to_number(entry.get(sum_key)) if sum_key else None
    if amount is None and price is not None:
        amount = price * (qty if qty is not None else 1)
    if not name and price is None and amount is None:
        return None
    return Item(name=name or "—", qty=qty if qty is not None else 1.0, price=price, amount=amount)


def build_invoice(invoice_id: str, rows: list[dict[str, Any]], source: Path) -> Invoice:
    first = rows[0]
    row_item_keys = {
        key
        for key in (
            pick_key(first, ITEM_NAME_KEYS),
            pick_key(first, ITEM_QTY_KEYS),
            pick_key(first, ITEM_PRICE_KEYS),
            pick_key(first, ITEM_SUM_KEYS),
        )
        if key
    }

    embedded = nested_items(first)
    items: list[Item] = []
    if embedded:
        items = [item for item in (make_item(entry) for entry in embedded) if item]
    elif row_item_keys:
        items = [item for item in (make_item(row) for row in rows) if item]

    fields = {
        key: value
        for key, value in first.items()
        if key not in row_item_keys and not isinstance(value, (list, dict))
    }

    if items and any(item.amount is not None for item in items):
        total = sum(item.amount or 0.0 for item in items)
    else:
        total = to_number(pick_value(first, TOTAL_KEYS))

    currency = to_text(pick_value(first, CURRENCY_KEYS)) or "RUB"
    return Invoice(
        invoice_id=invoice_id,
        fields=fields,
        items=items,
        total=total,
        currency=currency,
        source=source,
    )


def load_invoices(path: Path) -> list[Invoice]:
    """Read a data file and group its records by invoice id."""
    records = load_records(path)
    if not records:
        raise ValueError("файл не содержит данных")

    id_key = detect_id_key(records)
    order: list[str] = []
    grouped: dict[str, list[dict[str, Any]]] = {}
    for position, record in enumerate(records, start=1):
        key = to_text(record.get(id_key)) if id_key else ""
        key = key or f"запись {position}"
        if key not in grouped:
            grouped[key] = []
            order.append(key)
        grouped[key].append(record)

    return [build_invoice(key, grouped[key], path) for key in order]


# --------------------------------------------------------------------------
# Templates and rendering
# --------------------------------------------------------------------------


def font_family_of(stem: str) -> tuple[str, str, str]:
    """Guess family / weight / style from a font file name."""
    normalized = re.sub(r"[^a-z0-9]", "", stem.lower())
    weight = "400"
    for token, value in sorted(WEIGHT_TOKENS.items(), key=lambda pair: -len(pair[0])):
        if normalized.endswith(token):
            weight = value
            break
    style = "normal"
    for token, value in STYLE_TOKENS.items():
        if normalized.endswith(token):
            style = value
            break
    for prefix, family in KNOWN_FAMILIES.items():
        if normalized.startswith(prefix):
            return family, weight, style
    words = [part for part in re.split(r"[-_\s]+", stem) if part and part.lower() not in WEIGHT_TOKENS
             and part.lower() not in STYLE_TOKENS]
    return " ".join(words) or stem, weight, style


def build_font_css() -> tuple[str, str]:
    """Build @font-face rules from ./fonts and return the CSS plus the family name.

    We use relative paths (from BASE_DIR) because WeasyPrint resolves URLs
    against base_url, which is set to BASE_DIR. This avoids the unreliable
    file:// URL handling on some platforms.
    """
    rules: list[str] = []
    families: list[str] = []
    if FONTS_DIR.is_dir():
        for font in sorted(FONTS_DIR.iterdir()):
            if not font.is_file() or font.suffix.lower() not in {".ttf", ".otf", ".woff", ".woff2"}:
                continue
            family, weight, style = font_family_of(font.stem)
            # forward slashes: CSS treats backslashes as escape characters
            rel = font.relative_to(BASE_DIR).as_posix()
            rules.append(
                FACE_RULE.format(
                    family=family, src=rel, weight=weight, style=style
                )
            )
            if family not in families:
                families.append(family)

    stack = ", ".join(f'"{family}"' for family in families)
    fallback = '"DejaVu Sans", "Roboto", "Arial", sans-serif'
    css = "\n".join(rules + [f'html, body {{ font-family: {stack + ", " if stack else ""}{fallback}; }}'])
    # Empty family means "no font files in ./fonts, rely on system fonts".
    return css, (families[0] if families else "")


def inject_css(page: str, css: str) -> str:
    tag = f"\n  <style>\n{css}\n  </style>"
    match = re.search(r"<head[^>]*>", page, re.IGNORECASE)
    if match:
        return page[: match.end()] + tag + page[match.end():]
    match = re.search(r"<html[^>]*>", page, re.IGNORECASE)
    if match:
        return page[: match.end()] + f'\n<head><meta charset="utf-8">{tag}</head>' + page[match.end():]
    return tag.lstrip() + page


def items_rows(items: list[Item], currency: str) -> str:
    rows = []
    for position, item in enumerate(items, start=1):
        qty = to_text(item.qty)
        rows.append(
            "<tr>"
            f'<td class="c">{position}</td>'
            f'<td>{html_lib.escape(item.name, quote=False)}</td>'
            f'<td class="r">{qty}</td>'
            f'<td class="r">{money(item.price, currency)}</td>'
            f'<td class="r">{money(item.amount, currency)}</td>'
            "</tr>"
        )
    return "\n        ".join(rows) if rows else '<tr><td class="c">—</td><td>нет позиций</td><td class="r">—</td><td class="r">—</td><td class="r">—</td></tr>'


def build_context(invoice: Invoice, template: Path) -> dict[str, str]:
    context: dict[str, str] = {}
    for key, value in invoice.fields.items():
        context[str(key)] = escape(value)

    context["invoice_id"] = html_lib.escape(invoice.invoice_id, quote=False)
    context["date"] = escape(pick_value(invoice.fields, DATE_KEYS))
    context["date_fmt"] = escape(invoice.date_text)
    context["customer"] = escape(invoice.customer_text)
    context["seller"] = escape(invoice.field(SELLER_KEYS)) or context.get("seller", "")
    context["items_rows"] = items_rows(invoice.items, invoice.currency)
    context["items_count"] = str(len(invoice.items))
    context["total"] = money(invoice.total, invoice.currency)
    context["total_fmt"] = context["total"]
    context["total_raw"] = f"{invoice.total:.2f}" if invoice.total is not None else ""
    context["currency"] = escape(invoice.currency)
    context["currency_sign"] = CURRENCY_SIGNS.get(invoice.currency.upper(), invoice.currency)
    context["source_file"] = escape(invoice.source.name)
    context["template_name"] = escape(template.stem)
    context["template_file"] = escape(template.name)
    context["generated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
    return context


def fill_template(text: str, context: dict[str, str]) -> tuple[str, list[str]]:
    missing: list[str] = []

    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        if key in context:
            return context[key]
        if key not in missing:
            missing.append(key)
        return ""

    return PLACEHOLDER_RE.sub(replace, text), missing


# --------------------------------------------------------------------------
# Console menu
# --------------------------------------------------------------------------


class Quit(Exception):
    """The user chose to leave the menu."""


def rule(char: str = "─") -> str:
    return char * MENU_WIDTH


def section(title: str, hint: str = "") -> None:
    print()
    print(rule("═"))
    print(f"  {title}")
    if hint:
        print(f"  {hint}")
    print(rule("═"))


def ask_choice(count: int, label: str) -> int:
    while True:
        try:
            raw = input(f"  {label} [1–{count} | 0 — выход]: ").strip().lower()
        except EOFError:
            raise Quit
        if raw in {"0", "q", "quit", "exit", "выход"}:
            raise Quit
        if raw.isdigit() and 1 <= int(raw) <= count:
            return int(raw)
        print(f"  ⚠  Нужно число от 1 до {count} (или 0 для выхода). Вы ввели: «{raw}»")


# --------------------------------------------------------------------------
# PDF output
# --------------------------------------------------------------------------


def safe_name(text: str) -> str:
    cleaned = UNSAFE_NAME_RE.sub("_", text).strip("._")
    return cleaned[:60] or "receipt"


def open_pdf(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
        print("  ▶ PDF открыт в системной программе просмотра.")
    except Exception as exc:
        print(f"  ⚠  Не удалось открыть файл автоматически: {exc}")
        print(f"     Откройте его вручную: {path}")


def write_pdf(body_html: str, invoice: Invoice, template: Path) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    target = OUTPUT_DIR / f"{safe_name(invoice.invoice_id)}__{template.stem}__{stamp}.pdf"
    css, family = build_font_css()
    document = inject_css(body_html, css)
    HTML(string=document, base_url=str(BASE_DIR)).write_pdf(str(target))
    print(f"  ✎ Шрифт для кириллицы: {family or 'системный (DejaVu Sans / Arial)'}")
    return target


# --------------------------------------------------------------------------
# Interaction
# --------------------------------------------------------------------------


def choose_data_file() -> tuple[Path, list[Invoice]]:
    files = find_files(DATA_DIR, (".csv", ".json"))
    if not files:
        section("Файлы с данными не найдены")
        print(f"  Положите CSV или JSON в папку: {DATA_DIR}")
        raise Quit

    section("Шаг 1 из 3 — файл с данными", f"папка {rel(DATA_DIR)}")
    invoices_by_file: dict[Path, list[Invoice]] = {}
    readable: list[Path] = []
    entries: list[str] = []
    broken: list[tuple[Path, str]] = []
    for path in files:
        try:
            invoices = load_invoices(path)
        except Exception as exc:
            broken.append((path, str(exc)))
            continue
        invoices_by_file[path] = invoices
        readable.append(path)
        positions = sum(len(invoice.items) for invoice in invoices)
        entries.append(
            f"{rel(path)}   [{path.suffix.lower().lstrip('.')}]   "
            f"чеков: {len(invoices)}, позиций: {positions}"
        )

    if not entries:
        print("  ⚠  Ни один файл не удалось прочитать.")
        for path, error in broken:
            print(f"     {rel(path)} → {error}")
        raise Quit

    for number, entry in enumerate(entries, start=1):
        print(f"  {number:>2}) {entry}")
    for path, error in broken:
        print(f"  --  {rel(path)} — пропущен: {error}")

    chosen = readable[ask_choice(len(readable), "Файл") - 1]
    return chosen, invoices_by_file[chosen]


def choose_template() -> Path:
    templates = find_files(TEMPLATES_DIR, (".html", ".htm"))
    if not templates:
        section("Шаблон не найден")
        print(f"  Положите HTML-шаблон в папку: {TEMPLATES_DIR}")
        raise Quit

    section("Шаг 2 из 3 — HTML-шаблон", f'папка {rel(TEMPLATES_DIR)}')
    for number, path in enumerate(templates, start=1):
        title = template_title(path) or "без заголовка"
        print(f"  {number:>2}) {rel(path)}   —   {title}")

    return templates[ask_choice(len(templates), "Шаблон") - 1]


def template_title(path: Path) -> str:
    try:
        text = read_text(path)
    except OSError:
        return ""
    match = re.search(r"<title>(.*?)</title>", text, re.IGNORECASE | re.DOTALL)
    if not match:
        return ""
    title = re.sub(r"\{\{[^}]*\}\}", "", match.group(1))
    return html_lib.unescape(re.sub(r"\s+", " ", title)).strip()[:60]


def choose_invoice(invoices: list[Invoice], source: Path) -> Invoice:
    section(f"Шаг 3 из 3 — чек из файла {source.name}", f"всего чеков: {len(invoices)}")
    print(f"  {'№':>2}) {'Идентификатор':<15}  {'Дата':<10}  {'Покупатель':<28}  {'Поз.':>3}   Итого")
    print("  " + "─" * (MENU_WIDTH - 2))
    for number, invoice in enumerate(invoices, start=1):
        print(
            f"  {number:>2}) {clip(invoice.invoice_id, 15):<15}  "
            f"{clip(invoice.date_text, 10):<10}  "
            f"{clip(invoice.customer_text, 28):<28}  "
            f"{len(invoice.items):>3}   {money(invoice.total, invoice.currency)}"
        )

    return invoices[ask_choice(len(invoices), "Чек") - 1]


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(BASE_DIR))
    except ValueError:
        return str(path)


def generate(invoice: Invoice, template: Path) -> None:
    section(f"Генерируем PDF: чек {invoice.invoice_id} → {template.name}")
    context = build_context(invoice, template)
    page, missing = fill_template(read_text(template), context)
    if missing:
        print(f"  ⚠  В данных нет полей для: {', '.join(missing)} — вставлено пусто")
    if not invoice.items:
        print("  ⚠  В чеке не найдены позиции (колонки item/qty/price или массив items)")

    target = write_pdf(page, invoice, template)
    size = target.stat().st_size
    print(f"  ✔ Файл сохранён: {rel(target)}  ({size / 1024:.1f} КБ)")
    open_pdf(target)


def main() -> int:
    try:
        # Never crash on an unusual console code page: replace unencodable chars.
        sys.stdout.reconfigure(errors="replace")
        sys.stderr.reconfigure(errors="replace")
    except (AttributeError, OSError, ValueError):
        pass

    print()
    print(rule("═"))
    print("  PDF-ЧЕК-МЕЙКЕР")
    print("  данные → HTML-шаблон → PDF")
    print(rule("═"))

    if HTML is None:
        print(DEPENDENCY_HELP.format(error=WEASYPRINT_ERROR))
        return 1

    try:
        while True:
            source, invoices = choose_data_file()
            template = choose_template()
            invoice = choose_invoice(invoices, source)
            generate(invoice, template)
            print()
            try:
                again = input("  Сгенерировать ещё один чек? [y/N]: ").strip().lower()
            except EOFError:
                again = ""
            if again not in {"y", "ye", "yes", "д", "да"}:
                print("\n  До встречи!\n")
                return 0
    except Quit:
        print("\n  Остановлено. Файлы не изменены.\n")
        return 0
    except KeyboardInterrupt:
        print("\n\n  Прервано пользователем.\n")
        return 130
    except Exception as exc:
        print(f"\n  ✖ Ошибка: {type(exc).__name__}: {exc}\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
