import re
from io import BytesIO
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="DNP PDF Extractor", page_icon="📄", layout="wide")

COLS = ["Nomor ST", "Nama", "Total Biaya Riil"]

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 1

def norm(x):
    return re.sub(r"\s+", " ", str(x or "")).strip()

def get_st(text):
    for pattern in (r"Nomor\s*ST\s*:\s*([A-Z0-9./_-]+)", r"\b(ST-\d+/[A-Z0-9./_-]+)\b"):
        m = re.search(pattern, text or "", re.I)
        if m:
            return m.group(1)
    return ""

def rupiah(x):
    m = re.search(r"Rp\s*([\d.]+)", str(x or ""), re.I)
    if not m:
        return None
    digits = re.sub(r"\D", "", m.group(1))
    return int(digits) if digits else None

def person_name(cell):
    result = []
    for line in str(cell or "").splitlines():
        line = norm(line)
        line = re.sub(r"\b\d{10,20}\b", "", line)
        line = re.sub(r"\b\d{1,4}\s*/\s*[IVX]+/[A-Za-z]+\b", "", line)
        line = re.sub(r"\s*/\s*$", "", line).strip()
        if line:
            result.append(line)
    return norm(" ".join(result)).strip(" /")

def columns(table):
    for i, row in enumerate(table):
        headers = [norm(x).lower() for x in row]
        name = amount = number = None
        for j, h in enumerate(headers):
            if "nama" in h and "nip" in h: name = j
            if "total biaya" in h and "riil" in h: amount = j
            if h == "no" or h.startswith("no "): number = j
        if name is not None and amount is not None:
            return i, name, amount, number
    return None, None, None, None

def extract_pdf(data, filename):
    found, notes = [], []
    try:
        with pdfplumber.open(BytesIO(data)) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            st_number = get_st(text)
            if not st_number:
                notes.append(f"{filename}: Nomor ST tidak ditemukan.")
            for page_no, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables() or []
                if not tables:
                    try:
                        tables = page.extract_tables(table_settings={
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "intersection_tolerance": 5,
                            "snap_tolerance": 3,
                            "join_tolerance": 3,
                        }) or []
                    except Exception:
                        tables = []
                before = len(found)
                for table in tables:
                    if not table: continue
                    hi, ni, ai, oi = columns(table)
                    if hi is None: continue
                    for row in table[hi+1:]:
                        if not row: continue
                        no = norm(row[oi]) if oi is not None and oi < len(row) else ""
                        if not no.isdigit(): continue
                        name = person_name(row[ni] if ni < len(row) else "")
                        amount = rupiah(row[ai] if ai < len(row) else "")
                        if name and amount is not None:
                            found.append({"Nomor ST": st_number, "Nama": name, "Total Biaya Riil": amount})
                if len(found) == before and (page.extract_text() or "").strip():
                    notes.append(f"{filename}, halaman {page_no}: tabel tidak dikenali; periksa PDF atau gunakan OCR jika berupa scan.")
    except Exception as e:
        notes.append(f"{filename}: gagal membaca PDF ({e}).")
    return found, notes

def excel_bytes(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Hasil Ekstraksi")
        ws = writer.sheets["Hasil Ekstraksi"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in {"A": 28, "B": 36, "C": 24}.items():
            ws.column_dimensions[col].width = width
        for cell in ws["C"][1:]:
            cell.number_format = '"Rp" #,##0'
    return out.getvalue()


### UI Starts Here ###

st.title("📄 DNP PDF Extractor")
st.caption("Ekstrak Nomor ST, Nama, dan Total Biaya Riil dari satu atau banyak dokumen DNP.")
with st.sidebar:
    st.header("Cara penggunaan")
    st.markdown("1. Upload satu atau beberapa PDF.\n2. Klik **Proses PDF**.\n3. Periksa/koreksi hasil.\n4. Download CSV atau Excel.")
    st.info("Versi ini untuk PDF digital yang memiliki teks dan tabel. PDF scan/foto memerlukan OCR.")

files = st.file_uploader("Pilih file PDF", type=["pdf"], accept_multiple_files=True, key=st.session_state["uploader_key"],)
if files:
    st.success(f"{len(files)} file dipilih.")
    left, right = st.columns(2)
    if left.button("Hapus semua file", use_container_width=True):
        records, notes = [], []
        st.session_state["processed"] = False
        st.session_state["uploader_key"] += 1
        st.rerun()
    if right.button("⚙️ Proses PDF", type="primary", use_container_width=True):
        records, notes = [], []
        progress = st.progress(0)
        for i, file in enumerate(files):
            rows, warnings = extract_pdf(file.getvalue(), file.name)
            records.extend(rows)
            notes.extend(warnings)
            progress.progress((i + 1) / len(files))
        st.session_state["records"] = records
        st.session_state["notes"] = notes
        st.session_state["processed"] = True

if st.session_state.get("processed"):
    records = st.session_state.get("records", [])
    notes = st.session_state.get("notes", [])
    if notes:
        with st.expander(f"Catatan proses ({len(notes)})", expanded=True):
            for note in notes: st.warning(note)
    if not records:
        st.error("Tidak ada data yang berhasil diekstrak. Periksa struktur tabel PDF.")
    else:
        df = pd.DataFrame(records, columns=COLS)
        total = int(df["Total Biaya Riil"].sum())
        a, b, c = st.columns(3)
        a.metric("Jumlah orang", f"{len(df):,}")
        b.metric("Jumlah Nomor ST", f"{df['Nomor ST'].nunique():,}")
        c.metric("Total Biaya Riil", "Rp" + f"{total:,.0f}".replace(",", "."))
        st.subheader("Preview dan koreksi data")
        edited = st.data_editor(df, use_container_width=True, hide_index=True, num_rows="dynamic")
        left, right = st.columns(2)
        with left:
            st.download_button("⬇️ Download CSV", edited.to_csv(index=False).encode("utf-8-sig"),
                "hasil_ekstraksi_dnp.csv", "text/csv", use_container_width=True)
        with right:
            st.download_button("⬇️ Download Excel (.xlsx)", excel_bytes(edited),
                "hasil_ekstraksi_dnp.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
else:
    st.markdown("### Data yang diekstrak:\n- **Nomor ST** — nomor Surat Tugas dari header PDF.\n- **Nama** — nama pegawai tanpa NIP/pangkat.\n- **Total Biaya Riil** — angka Rupiah yang bisa dihitung di Excel.")
