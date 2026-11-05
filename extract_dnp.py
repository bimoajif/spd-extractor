#!/usr/bin/env python3
"""
Extract Nomor ST, Nama, and Total Biaya Riil from Indonesian DNP travel-cost PDFs.

Install:
    python -m pip install pdfplumber pandas

Usage:
    python extract_dnp.py "DNP ST-486.pdf"
    python extract_dnp.py ./folder_pdf --output ./hasil
    python extract_dnp.py "DNP ST-486.pdf" --json hasil.json --csv hasil.csv
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import pdfplumber
import pandas as pd


COLUMNS = ["Nomor ST", "Nama", "Total Biaya Riil"]


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def extract_nomor_st(text: str) -> str:
    match = re.search(
        r"Nomor\s*ST\s*:\s*([A-Z0-9./_-]+)",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        # Some documents show the ST number in the title rather than its label.
        match = re.search(
            r"\bST-\d+/[A-Z0-9./_-]+",
            text,
            flags=re.IGNORECASE,
        )
        return match.group(0).strip() if match else ""
    return match.group(1).strip()


def clean_name(cell: str) -> str:
    """Remove line breaks, NIP, and rank from the Nama/NIP/Gol table cell."""
    lines = [normalize_space(x) for x in (cell or "").splitlines()]
    lines = [x for x in lines if x]

    # Remove NIP lines (typically 15 digits, sometimes split across lines).
    cleaned = []
    for line in lines:
        if re.fullmatch(r"\d{10,20}", line):
            continue
        # Remove a rank-only line such as "003 / III/a".
        if re.fullmatch(r"\d{1,4}\s*/\s*[IVX]+/[A-Za-z]+", line):
            continue
        # Remove a NIP followed by other text, if present.
        line = re.sub(r"\b\d{10,20}\b", "", line).strip()
        line = re.sub(r"\s*/\s*$", "", line).strip()
        if line:
            cleaned.append(line)

    name = normalize_space(" ".join(cleaned))
    name = re.sub(r"\s*/\s*", " ", name).strip()
    return name


def parse_rupiah(value: str) -> int | None:
    """Convert values such as 'Rp400.000' into integer 400000."""
    if not value:
        return None
    match = re.search(r"Rp\s*([\d.]+)", value, flags=re.IGNORECASE)
    if not match:
        return None
    digits = re.sub(r"\D", "", match.group(1))
    return int(digits) if digits else None


def extract_from_pdf(pdf_path: Path) -> list[dict]:
    records: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        nomor_st = extract_nomor_st(full_text)

        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []

            for table in tables:
                if not table:
                    continue

                # Locate the table header and relevant columns.
                header_index = None
                name_col = amount_col = number_col = None

                for i, row in enumerate(table):
                    header = [normalize_space(str(cell or "")).lower() for cell in row]
                    for j, cell in enumerate(header):
                        if "nama" in cell and "nip" in cell:
                            name_col = j
                        if "total biaya" in cell and "riil" in cell:
                            amount_col = j
                        if cell == "no" or cell.startswith("no "):
                            number_col = j
                    if name_col is not None and amount_col is not None:
                        header_index = i
                        break

                if header_index is None:
                    continue

                for row in table[header_index + 1:]:
                    if not row:
                        continue

                    # Skip the summary/total row; only numbered people are wanted.
                    row_number = str(row[number_col] or "").strip() if number_col is not None and number_col < len(row) else ""
                    if not row_number.isdigit():
                        continue

                    name_cell = str(row[name_col] or "") if name_col < len(row) else ""
                    amount_cell = str(row[amount_col] or "") if amount_col < len(row) else ""

                    name = clean_name(name_cell)
                    amount = parse_rupiah(amount_cell)

                    if name and amount is not None:
                        records.append({
                            "Nomor ST": nomor_st,
                            "Nama": name,
                            "Total Biaya Riil": amount,
                        })

    return records


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract Nomor ST, Nama, and Total Biaya Riil from DNP PDF files."
    )
    parser.add_argument("input", help="PDF file or folder containing PDF files")
    parser.add_argument("--output", default="hasil_ekstraksi", help="Output folder (default: hasil_ekstraksi)")
    parser.add_argument("--json", dest="json_file", help="Optional explicit JSON output path")
    parser.add_argument("--csv", dest="csv_file", help="Optional explicit CSV output path")
    parser.add_argument("--xlsx", dest="xlsx_file", help="Optional explicit XLSX output path")
    args = parser.parse_args()

    source = Path(args.input)
    if source.is_file() and source.suffix.lower() == ".pdf":
        pdf_files = [source]
    elif source.is_dir():
        pdf_files = sorted(source.glob("*.pdf"))
    else:
        raise SystemExit(f"Input tidak ditemukan atau bukan PDF/folder: {source}")

    if not pdf_files:
        raise SystemExit("Tidak ada file PDF yang ditemukan.")

    all_records = []
    for pdf_file in pdf_files:
        extracted = extract_from_pdf(pdf_file)
        print(f"{pdf_file.name}: {len(extracted)} orang diekstrak")
        all_records.extend(extracted)

    df = pd.DataFrame(all_records, columns=COLUMNS)

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    json_path = Path(args.json_file) if args.json_file else output_dir / "hasil_ekstraksi.json"
    csv_path = Path(args.csv_file) if args.csv_file else output_dir / "hasil_ekstraksi.csv"
    xlsx_path = Path(args.xlsx_file) if args.xlsx_file else output_dir / "hasil_extraksi.xlsx"

    json_path.parent.mkdir(parents=True, exist_ok=True)
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    df.to_json(json_path, orient="records", force_ascii=False, indent=2)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    df.to_excel(xlsx_path, index=False)

    print("\nHasil:")
    print(df.to_string(index=False))
    print(f"\nJSON: {json_path}")
    print(f"CSV : {csv_path}")
    print(f"XLSX : {xlsx_path}")


if __name__ == "__main__":
    main()
