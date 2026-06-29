"""
Rekap Bukti Potong & Faktur Pajak
Mendukung: PPh 21 (BP21), PPh 23 (BPPU), PPN Keluaran, PPN Masukan
"""

import os
import re
import pandas as pd
from glob import glob
from pypdf import PdfReader

# ========================== KONFIGURASI ==========================
FOLDER_PDF = r"E:\Faktur Pajak All"
OUTPUT_EXCEL = "rekap_faktur.xlsx"
DEBUG_MODE = False   # True = tampilkan detail per file
# ================================================================

BULAN_MAP = {
    'januari': '01', 'februari': '02', 'maret': '03', 'april': '04',
    'mei': '05', 'juni': '06', 'juli': '07', 'agustus': '08',
    'september': '09', 'oktober': '10', 'november': '11', 'desember': '12'
}


# ────────────────────────────────────────────────────────────────
# UTILITAS
# ────────────────────────────────────────────────────────────────

def bersihkan_angka(teks: str):
    """
    Konversi string angka Indonesia ke integer.
    Menangani dua format:
      - Titik = pemisah ribuan saja : '23.333.333'   -> 23333333
      - Titik = ribuan, koma = desimal: '2.643.352.479,00' -> 2643352479
    """
    if not teks:
        return None
    teks = str(teks).strip()
    # Jika ada koma, anggap koma adalah pemisah desimal -> buang bagian desimal
    if ',' in teks:
        teks = teks.split(',')[0]
    # Hapus semua karakter bukan digit
    bersih = re.sub(r'[^\d]', '', teks)
    return int(bersih) if bersih else None


def cari(pattern: str, teks: str, group: int = 1, flags=re.IGNORECASE | re.DOTALL):
    """re.search shortcut -> string bersih atau None."""
    m = re.search(pattern, teks, flags)
    if m:
        return re.sub(r'\s+', ' ', m.group(group)).strip()
    return None


def baca_pdf(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    return "\n".join(
        page.extract_text() or "" for page in reader.pages
    )


def deteksi_jenis(teks: str, nama_file: str) -> str:
    t = teks.upper()
    f = nama_file.upper()
    if 'BP21' in t or 'PAJAK PENGHASILAN PASAL 21' in t:
        return 'PPh21'
    if 'BPPU' in t or 'PPH UNIFIKASI' in t or 'PASAL 23' in t or 'PASAL 4' in t:
        return 'PPh23'
    if 'FAKTUR PAJAK' in t or 'KODE DAN NOMOR SERI FAKTUR' in t:
        if 'MASUKAN' in f or 'MASUK' in f:
            return 'PPN_Masukan'
        return 'PPN_Keluaran'   # default jika tidak ada kata "masukan"
    return 'UNKNOWN'


def tanggal_ke_masa(tanggal_str: str) -> str | None:
    """'20 Januari 2025' -> '01-2025'"""
    if not tanggal_str:
        return None
    for nama, num in BULAN_MAP.items():
        if nama in tanggal_str.lower():
            tahun = re.search(r'(\d{4})', tanggal_str)
            if tahun:
                return f"{num}-{tahun.group(1)}"
    return None


# ────────────────────────────────────────────────────────────────
# PARSER PPh 21  (BP21)
# Tabel: PENGHASILAN BRUTO | DPP(%) | TARIF(%) | PPh DIPOTONG
# ────────────────────────────────────────────────────────────────

def parse_pph21(teks: str) -> dict:
    d = {}

    m = re.search(r'(\w{5,15})\s+(\d{2}-\d{4})\s+(?:TIDAK FINAL|FINAL)', teks)
    if m:
        d['nomor']      = m.group(1)
        d['masa_pajak'] = m.group(2)

    d['npwp_penerima'] = cari(r'A\.1\s+NIK/NPWP\s*:\s*(\d+)', teks)
    d['nama_penerima'] = cari(r'A\.2\s+Nama\s*:\s*(.+?)(?=\s*A\.3)', teks)

    d['kode_objek']  = cari(r'(\d{2}-\d{3}-\d{2})', teks)
    d['objek_pajak'] = cari(
        r'\d{2}-\d{3}-\d{2}\s+(.+?)(?=\s*\d{1,3}(?:\.\d{3})+\s+\d+\s+\d+)', teks
    )

    # [Bruto] [DPP%] [Tarif%] [PPh]  — sebelum B.8
    m = re.search(
        r'(\d{1,3}(?:\.\d{3})+)\s+(\d{1,3})\s+(\d{1,2})\s+(\d{1,3}(?:\.\d{3})*)\s+B\.8',
        teks
    )
    if m:
        d['penghasilan_bruto'] = bersihkan_angka(m.group(1))
        d['dpp_persen']        = int(m.group(2))
        d['tarif_persen']      = int(m.group(3))
        d['pph_dipotong']      = bersihkan_angka(m.group(4))
        if d.get('penghasilan_bruto') and d.get('dpp_persen'):
            d['dpp'] = int(d['penghasilan_bruto'] * d['dpp_persen'] / 100)

    d['tanggal_dokumen'] = cari(r'Tanggal\s+Dokumen\s*:\s*(.+?)(?=\s*B\.9)', teks)
    d['npwp_pemotong']   = cari(r'C\.1\s+NPWP/NIK\s*:\s*(\d+)', teks)
    d['nama_pemotong']   = cari(r'C\.3\s+Nama\s+Pemotong\s*:\s*(.+?)(?=\s*C\.4)', teks)
    d['tanggal']         = cari(r'C\.4\s+Tanggal\s*:\s*(.+?)(?=\s*C\.5)', teks)
    d['masa_pajak']      = d.get('masa_pajak') or tanggal_ke_masa(d.get('tanggal', ''))

    return d


# ────────────────────────────────────────────────────────────────
# PARSER PPh 23 / UNIFIKASI  (BPPU)
# Tabel: DPP(Rp) | TARIF(%) | PAJAK PENGHASILAN(Rp)
# ────────────────────────────────────────────────────────────────

def parse_pph23(teks: str) -> dict:
    d = {}

    m = re.search(r'(\w{5,15})\s+(\d{2}-\d{4})\s+(?:TIDAK FINAL|FINAL)', teks)
    if m:
        d['nomor']      = m.group(1)
        d['masa_pajak'] = m.group(2)

    d['jenis_pph']     = cari(r'B\.2\s+Jenis\s+PPh\s*:\s*(.+?)(?=\s*KODE|\s*B\.3)', teks)
    d['npwp_penerima'] = cari(r'A\.1\s+NPWP\s*/\s*NIK\s*:\s*(\d+)', teks)
    d['nama_penerima'] = cari(r'A\.2\s+NAMA\s*:\s*(.+?)(?=\s*A\.3)', teks)
    d['kode_objek']    = cari(r'(\d{2}-\d{3}-\d{2})', teks)
    d['objek_pajak']   = cari(
        r'\d{2}-\d{3}-\d{2}\s+(.+?)(?=\s*\d{1,3}(?:\.\d{3})+\s+\d+\s+\d{1,3}(?:\.\d{3}))', teks
    )

    # [DPP] [Tarif%] [PPh]  — sebelum B.8
    m = re.search(
        r'(\d{1,3}(?:\.\d{3})+)\s+(\d{1,3})\s+(\d{1,3}(?:\.\d{3})*)\s+B\.8',
        teks
    )
    if m:
        d['dpp']          = bersihkan_angka(m.group(1))
        d['tarif_persen'] = int(m.group(2))
        d['pph_dipotong'] = bersihkan_angka(m.group(3))

    d['tanggal_dokumen'] = cari(r'Tanggal\s*:\s*(.+?)(?=\s*B\.9)', teks)
    d['npwp_pemotong']   = cari(r'C\.1\s+NPWP\s*/\s*NIK\s*:\s*(\d+)', teks)
    d['nama_pemotong']   = cari(r'C\.3\s+NAMA\s+PEMOTONG[^:]*:\s*(.+?)(?=\s*C\.4)', teks)
    d['tanggal']         = cari(r'C\.4\s+TANGGAL\s*:\s*(.+?)(?=\s*C\.5)', teks)
    d['masa_pajak']      = d.get('masa_pajak') or tanggal_ke_masa(d.get('tanggal', ''))

    return d


# ────────────────────────────────────────────────────────────────
# PARSER PPN (Faktur Pajak e-Faktur)
# Berlaku untuk PPN Keluaran maupun PPN Masukan
# ────────────────────────────────────────────────────────────────

def parse_ppn(teks: str) -> dict:
    d = {}

    # ── Nomor Faktur (17 digit, tanpa titik/strip) ────────────
    d['nomor'] = cari(r'Kode\s+dan\s+Nomor\s+Seri\s+Faktur\s+Pajak\s*:\s*(\d{15,17})', teks)

    # ── Tanggal & Masa Pajak ──────────────────────────────────
    # Format: "KOTA ..., 20 Januari\n2025"  atau  "09 Mei 2025"
    m = re.search(
        r',\s*(\d{1,2}\s+(?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|'
        r'September|Oktober|November|Desember)[\s\n]+\d{4})',
        teks, re.IGNORECASE
    )
    if m:
        tanggal_raw = re.sub(r'\s+', ' ', m.group(1)).strip()
        d['tanggal']   = tanggal_raw
        d['masa_pajak'] = tanggal_ke_masa(tanggal_raw)

    # ── Penjual (Pengusaha Kena Pajak / PKP) ─────────────────
    blok_pkp = re.search(
        r'Pengusaha\s+Kena\s+Pajak\s*:(.*?)(?=Pembeli\s+Barang)',
        teks, re.DOTALL | re.IGNORECASE
    )
    if blok_pkp:
        b = blok_pkp.group(1)
        d['nama_penjual'] = cari(r'Nama\s*:\s*(.+?)(?=\s*Alamat|\s*NPWP|\s*#)', b)
        d['npwp_penjual'] = cari(r'NPWP\s*:\s*(\d+)', b)

    # ── Pembeli ───────────────────────────────────────────────
    blok_pembeli = re.search(
        r'Pembeli\s+Barang\s+Kena\s+Pajak[^:]*:(.*?)(?=No\.\s|Email\s*:|NIK\s*:)',
        teks, re.DOTALL | re.IGNORECASE
    )
    if blok_pembeli:
        b = blok_pembeli.group(1)
        d['nama_pembeli'] = cari(r'Nama\s*:\s*(.+?)(?=\s*Alamat|\s*NPWP|\s*#)', b)
        d['npwp_pembeli'] = cari(r'NPWP\s*:\s*(\d+)', b)

    # ── Nama Barang / Jasa (baris pertama dalam tabel) ────────
    # Setelah "1 [kode 6 digit]" muncul nama barang di baris berikutnya
    m = re.search(r'\b1\s+\d{6}\s*\n(.+?)(?=\n)', teks)
    if m:
        d['nama_barang_jasa'] = m.group(1).strip()

    # ── DPP ───────────────────────────────────────────────────
    # Format: "Dasar Pengenaan Pajak 2.643.352.479,00"
    d['dpp'] = bersihkan_angka(
        cari(r'Dasar\s+Pengenaan\s+Pajak\s+([\d.,]+)', teks)
    )

    # ── PPN ───────────────────────────────────────────────────
    # Format: "Jumlah PPN (Pajak Pertambahan Nilai) 317.202.298,00"
    d['pph_dipotong'] = bersihkan_angka(
        cari(r'Jumlah\s+PPN\s*\([^)]+\)\s+([\d.,]+)', teks)
    )

    # ── Nomor Referensi ───────────────────────────────────────
    d['referensi'] = cari(r'\(Referensi\s*:\s*(.+?)\)', teks)

    return d


# ────────────────────────────────────────────────────────────────
# MAIN EKSTRAKSI
# ────────────────────────────────────────────────────────────────

def ekstrak_data_faktur(pdf_path: str) -> dict:
    base = {
        # Umum
        'file_name':         os.path.basename(pdf_path),
        'jenis_dokumen':     None,
        'nomor':             None,
        'masa_pajak':        None,
        'tanggal':           None,
        'tanggal_dokumen':   None,
        # PPh 21 / 23
        'npwp_penerima':     None,
        'nama_penerima':     None,
        'npwp_pemotong':     None,
        'nama_pemotong':     None,
        'kode_objek':        None,
        'objek_pajak':       None,
        'jenis_pph':         None,
        'penghasilan_bruto': None,
        'dpp_persen':        None,
        'tarif_persen':      None,
        # PPN (kolom baru)
        'nama_penjual':      None,
        'npwp_penjual':      None,
        'nama_pembeli':      None,
        'npwp_pembeli':      None,
        'nama_barang_jasa':  None,
        'referensi':         None,
        # Nilai (semua jenis)
        'dpp':               None,
        'pph_dipotong':      None,
    }

    try:
        teks = baca_pdf(pdf_path)

        if not teks.strip():
            base['jenis_dokumen'] = 'SCAN_KOSONG'
            print(f"  WARNING PDF scan (tidak ada teks): {base['file_name']}")
            return base

        jenis = deteksi_jenis(teks, base['file_name'])
        base['jenis_dokumen'] = jenis

        if DEBUG_MODE:
            print(f"\n{'='*60}\nFILE: {base['file_name']} [{jenis}]")
            print(teks[:500])

        if jenis == 'PPh21':
            hasil = parse_pph21(teks)
        elif jenis == 'PPh23':
            hasil = parse_pph23(teks)
        elif jenis in ('PPN_Keluaran', 'PPN_Masukan'):
            hasil = parse_ppn(teks)
        else:
            print(f"  ? Jenis tidak dikenali: {base['file_name']}")
            hasil = {}

        base.update(hasil)

        if DEBUG_MODE:
            for k, v in base.items():
                if v and k != 'file_name':
                    print(f"  {k:<22}: {v}")

    except Exception as e:
        print(f"  ERROR {os.path.basename(pdf_path)}: {e}")

    return base


# ────────────────────────────────────────────────────────────────
# MAIN
# ────────────────────────────────────────────────────────────────

KOLOM_OUTPUT = [
    # Identitas dokumen
    'file_name', 'jenis_dokumen', 'nomor', 'masa_pajak', 'tanggal', 'tanggal_dokumen',
    # PPh 21 / 23 — penerima & pemotong
    'npwp_penerima', 'nama_penerima',
    'npwp_pemotong', 'nama_pemotong',
    # PPh — detail objek
    'kode_objek', 'objek_pajak', 'jenis_pph',
    'penghasilan_bruto', 'dpp_persen', 'tarif_persen',
    # PPN — penjual & pembeli (kolom baru)
    'npwp_penjual', 'nama_penjual',
    'npwp_pembeli', 'nama_pembeli',
    'nama_barang_jasa', 'referensi',
    # Nilai (semua jenis)
    'dpp', 'pph_dipotong',
]


def main():
    files = glob(os.path.join(FOLDER_PDF, "*.pdf"))
    if not files:
        print(f"Tidak ada PDF di: {FOLDER_PDF}")
        return

    print(f"Ditemukan {len(files)} file PDF. Memproses...\n")

    hasil = []
    for i, fp in enumerate(files, 1):
        print(f"  ({i:>3}/{len(files)}) {os.path.basename(fp)}")
        hasil.append(ekstrak_data_faktur(fp))

    df = pd.DataFrame(hasil, columns=KOLOM_OUTPUT)

    # ── Ringkasan per jenis ──────────────────────────────────
    print("\nRingkasan per jenis dokumen:")
    CEK = {
        'PPh21':        ['nomor', 'masa_pajak', 'nama_penerima', 'penghasilan_bruto', 'dpp', 'pph_dipotong'],
        'PPh23':        ['nomor', 'masa_pajak', 'nama_penerima', 'dpp', 'tarif_persen', 'pph_dipotong'],
        'PPN_Keluaran': ['nomor', 'masa_pajak', 'tanggal', 'nama_penjual', 'nama_pembeli', 'dpp', 'pph_dipotong'],
        'PPN_Masukan':  ['nomor', 'masa_pajak', 'tanggal', 'nama_penjual', 'nama_pembeli', 'dpp', 'pph_dipotong'],
    }
    for jenis, grup in df.groupby('jenis_dokumen'):
        print(f"\n  [{jenis}] -- {len(grup)} file")
        for col in CEK.get(jenis, ['nomor', 'dpp']):
            ok = grup[col].notna().sum()
            bar = '#' * ok + '-' * (len(grup) - ok)
            print(f"    {col:<22}: {ok:>3}/{len(grup)} [{bar}]")

    # ── Simpan ke Excel dengan sheet per jenis ───────────────
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        # Sheet gabungan semua
        df.to_excel(writer, index=False, sheet_name='Semua')

        # Sheet per jenis dokumen
        for jenis, grup in df.groupby('jenis_dokumen'):
            # Hapus kolom kosong sepenuhnya agar lebih rapi
            grup_bersih = grup.dropna(axis=1, how='all')
            sheet_name  = jenis[:31]   # Excel max 31 karakter
            grup_bersih.to_excel(writer, index=False, sheet_name=sheet_name)

    print(f"\nSelesai! Disimpan ke: {OUTPUT_EXCEL}")
    print(f"Sheet: 'Semua' + {df['jenis_dokumen'].nunique()} sheet per jenis\n")

    cols_preview = ['file_name', 'jenis_dokumen', 'nomor', 'masa_pajak', 'dpp', 'pph_dipotong']
    print("Pratinjau 5 data pertama:")
    print(df[cols_preview].head().to_string(index=False))


if __name__ == "__main__":
    main()