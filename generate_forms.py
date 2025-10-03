#!/usr/bin/env python3
"""General DOCX filler from CSV rows.

This tool reads a CSV file whose first row contains the column labels. Each
subsequent row is converted into a dictionary and used to fill a DOCX
template. The template must contain placeholders wrapped in double brackets,
for example ``[[Nombre]]`` or ``[[RUT]]``. Matching between CSV labels and
placeholders is case-insensitive and ignores leading ``#`` symbols as well as
repeated whitespace.

Example::

    python generate_forms.py estudiantes.csv plantilla.docx --name-column Nombres

The command above will generate one document per row inside ``./salida`` (by
default) and name each file using the ``Nombres`` column.
"""

from __future__ import annotations

import argparse
import csv
import re
import unicodedata
import sys
from pathlib import Path
from typing import Dict, Iterable, Iterator, Optional

from docx import Document

PLACEHOLDER_PATTERN = re.compile(r"\[\[\s*([^\]]+?)\s*\]\]", re.IGNORECASE)
NAME_TEMPLATE_PATTERN = re.compile(r"\$([^$\[\]]+?)(?:\[(\d+)\])?")


def normalize_label(label: Optional[str]) -> str:
    """Return a normalized version of a CSV or placeholder label."""

    if not label:
        return ""
    clean = label.strip()
    if clean.startswith("#"):
        clean = clean[1:]
    clean = re.sub(r"\s+", " ", clean)
    return clean.casefold()


def strip_diacritics(text: str) -> str:
    """Return ``text`` without diacritical marks."""

    normalized = unicodedata.normalize("NFKD", text)
    return "".join(char for char in normalized if not unicodedata.combining(char))


def sanitize_filename(raw: str) -> str:
    """Return a filesystem-friendly version of ``raw``."""

    value = strip_diacritics(raw.strip())
    value = re.sub(r"\s+", "_", value)
    value = re.sub(r"[^A-Za-z0-9._-]", "", value)
    return value


def iter_paragraphs(element) -> Iterator:
    """Yield all paragraphs contained in ``element`` (document, cell, etc.)."""

    if hasattr(element, "paragraphs"):
        for paragraph in element.paragraphs:
            yield paragraph
    if hasattr(element, "tables"):
        for table in element.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from iter_paragraphs(cell)


def replace_placeholders(text: str, replacements: Dict[str, str]) -> str:
    """Replace ``[[...]]`` placeholders in ``text`` using ``replacements``."""

    def repl(match: re.Match[str]) -> str:
        raw_key = match.group(1)
        key = normalize_label(raw_key)
        if key in replacements:
            return replacements[key]
        return match.group(0)

    return PLACEHOLDER_PATTERN.sub(repl, text)


def apply_replacements(doc: Document, replacements: Dict[str, str]) -> None:
    """Modify ``doc`` in-place, replacing placeholders for every paragraph."""

    for paragraph in iter_paragraphs(doc):
        original = paragraph.text
        updated = replace_placeholders(original, replacements)
        if updated != original:
            paragraph.text = updated


def clean_cell_value(value: Optional[str]) -> str:
    """Return a normalized string for a CSV cell value."""

    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return text


def load_rows(csv_path: Path, encoding: str) -> Iterable[Dict[str, str]]:
    """Yield dictionaries for each non-empty row in ``csv_path``."""

    try:
        with csv_path.open(newline="", encoding=encoding) as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames is None:
                raise ValueError("El CSV no contiene encabezados")
            for row in reader:
                if row is None:
                    continue
                # Skip completely empty rows
                if all((value is None or str(value).strip() == "") for value in row.values()):
                    continue
                yield {key: clean_cell_value(value) for key, value in row.items() if key}
    except UnicodeDecodeError as exc:
        raise UnicodeDecodeError(
            exc.encoding or encoding,
            exc.object,
            exc.start,
            exc.end,
            f"No se pudo leer el CSV con la codificación {encoding!r}"
        ) from exc


def row_to_replacements(row: Dict[str, str]) -> Dict[str, str]:
    """Return a mapping of normalized label -> value for a CSV row."""

    replacements: Dict[str, str] = {}
    for label, value in row.items():
        key = normalize_label(label)
        replacements[key] = value
    return replacements


def resolve_name(row: Dict[str, str], replacements: Dict[str, str], *, name_column: Optional[str], index: int) -> str:
    """Return a base filename for the row."""

    if name_column:
        target = ""
        if NAME_TEMPLATE_PATTERN.search(name_column):
            target = evaluate_name_template(name_column, replacements)
        else:
            target = replacements.get(normalize_label(name_column), "")
        target = target.strip()
        if target:
            sanitized = sanitize_filename(target)
            if sanitized:
                return sanitized

    # fallback to first non-empty column value
    for value in row.values():
        if value and value.strip():
            sanitized = sanitize_filename(value)
            if sanitized:
                return sanitized

    return f"row_{index:03d}"


def evaluate_name_template(template: str, replacements: Dict[str, str]) -> str:
    """Return a filename string generated from ``template``.

    ``template`` may contain expressions of the form ``$Campo`` or
    ``$Campo[0]``. The former inserts the entire value stored in the
    corresponding column, while the latter first splits the value on
    whitespace and then selects the element at the requested index. Column
    labels follow the same normalization rules used for placeholders.
    Missing columns or indexes yield empty strings.
    """

    def repl(match: re.Match[str]) -> str:
        label, index_str = match.groups()
        key = normalize_label(label)
        value = replacements.get(key, "")
        if not value:
            return ""
        if index_str is not None:
            try:
                index = int(index_str)
            except ValueError:
                return ""
            parts = value.split()
            if 0 <= index < len(parts):
                return parts[index]
            return ""
        return value

    return NAME_TEMPLATE_PATTERN.sub(repl, template)


def main() -> None:
    parser = argparse.ArgumentParser(description="Rellena un DOCX por cada fila del CSV usando marcadores [[...]].")
    parser.add_argument("csv", help="Ruta al CSV que contiene los datos")
    parser.add_argument("template", help="Ruta al DOCX con marcadores del tipo [[Campo]]")
    parser.add_argument("--outdir", default="salida", help="Directorio donde se guardarán los archivos generados")
    parser.add_argument("--name-column", help="Nombre de la columna para el nombre del archivo de salida")
    parser.add_argument(
        "--encoding",
        default="utf-8-sig",
        help="Codificación usada para leer el CSV (por defecto utf-8-sig)",
    )
    args = parser.parse_args()

    csv_path = Path(args.csv)
    template_path = Path(args.template)
    outdir = Path(args.outdir)

    if not csv_path.exists():
        print(f"[ERROR] No existe el archivo CSV: {csv_path}")
        sys.exit(1)
    if not template_path.exists():
        print(f"[ERROR] No existe la plantilla DOCX: {template_path}")
        sys.exit(1)

    outdir.mkdir(parents=True, exist_ok=True)

    try:
        rows = list(load_rows(csv_path, args.encoding))
    except (UnicodeDecodeError, ValueError) as exc:
        print(f"[ERROR] {exc}")
        sys.exit(1)
    if not rows:
        print("[WARN] El CSV no contiene filas con datos. Nada que generar.")
        return

    for index, row in enumerate(rows, start=1):
        replacements = row_to_replacements(row)
        document = Document(str(template_path))
        apply_replacements(document, replacements)

        base_name = resolve_name(row, replacements, name_column=args.name_column, index=index)
        output_path = outdir / f"{base_name}.docx"

        # Avoid accidental overwrite by appending a counter if needed
        counter = 1
        final_path = output_path
        while final_path.exists():
            final_path = outdir / f"{base_name}_{counter}.docx"
            counter += 1

        document.save(str(final_path))
        print(f"[OK] Generado: {final_path}")


if __name__ == "__main__":
    main()

