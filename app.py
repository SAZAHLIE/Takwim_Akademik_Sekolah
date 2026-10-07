import io
import os
import re
import base64
from datetime import date, datetime, timedelta, timezone
import holidays
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.protection import Protection
import pandas as pd
import pypdf
import streamlit as st
import streamlit.components.v1 as components

# Pustaka ReportLab untuk penjanaan PDF Cetakan Kemas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------------------------------------------
# KONFIGURASI LAMAN WEB STREAMLIT
# -------------------------------------------------------------
st.set_page_config(
    page_title="Sistem Kalender Akademik & eRPH",
    layout="wide",
    page_icon="🏫",
)

# -------------------------------------------------------------
# SKRIP SEKATAN KLIK KANAN & VIEW PAGE SOURCE
# -------------------------------------------------------------
st_security_script = """
<script>
    (function() {
        const targetWindow = window.parent || window;

        targetWindow.addEventListener('contextmenu', function(e) {
            e.preventDefault();
            e.stopPropagation();
            return false;
        }, true);

        targetWindow.addEventListener('keydown', function(e) {
            const isCmdOrCtrl = e.ctrlKey || e.metaKey;
            
            if (e.keyCode === 123 || e.key === 'F12') {
                e.preventDefault();
                e.stopPropagation();
                return false;
            }
            
            if (isCmdOrCtrl) {
                const key = e.key.toLowerCase();
                if (key === 'u' || key === 's') {
                    e.preventDefault();
                    e.stopPropagation();
                    return false;
                }
                if (e.shiftKey && (key === 'i' || key === 'j' || key === 'c')) {
                    e.preventDefault();
                    e.stopPropagation();
                    return false;
                }
            }
        }, true);
    })();
</script>
"""
components.html(st_security_script, height=0, width=0)

# -------------------------------------------------------------
# CARIAN LOGO UNTUK KAD HAKCIPTA
# -------------------------------------------------------------
possible_logos = ["logo.png", "logo.jpg", "logo.jpeg", "SEKOLAH.png", "SEKOLAH.jpg"]
logo_path = None

for f in possible_logos:
    if os.path.exists(f):
        logo_path = f
        break

if not logo_path:
    for f in os.listdir("."):
        if f.lower().endswith((".png", ".jpg", ".jpeg")):
            logo_path = f
            break

logo_html = ""
if logo_path and os.path.exists(logo_path):
    try:
        with open(logo_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        ext = logo_path.split(".")[-1].lower()
        mime_type = "image/png" if ext == "png" else "image/jpeg"
        logo_html = f'<img src="data:{mime_type};base64,{encoded_string}" style="max-height: 65px; max-width: 80px; object-fit: contain;" />'
    except Exception:
        logo_html = ""

# -------------------------------------------------------------
# HEADER & HAKCIPTA EKSLUSIF
# -------------------------------------------------------------
col_title, col_copyright = st.columns([3.2, 1.3])

with col_title:
    st.title("🏫 Sistem Kalender Akademik & eRPH")
    st.markdown(
        "Aplikasi pintar khusus untuk **Prasekolah, Sekolah Rendah, Sekolah Menengah, dan Sekolah Agama Negeri (SAN)**. "
        "Sistem ini membaca maklumat rasmi daripada fail PDF Kalendar Persekolahan KPM bagi menjana Kalender Akademik secara automatik."
    )

with col_copyright:
    st.markdown(
        f'''
        <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; padding: 12px 16px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.03);">
            <div style="display: flex; justify-content: space-between; align-items: center; gap: 12px;">
                <div style="text-align: left; display: flex; flex-direction: column; gap: 4px;">
                    <p style="margin: 0; font-size: 12px; font-weight: 600; color: #1f4e78;">🔒 Hakcipta Terpelihara</p>
                    <p style="margin: 0; font-size: 14px; font-weight: 700; color: #0d6efd;">© Sazahlie S.</p>
                    <div style="margin-top: 2px;">
                        <a href="mailto:sazahlie.sapar@moe.edu.my" style="display: inline-block; background-color: #0d6efd; color: white; padding: 4px 10px; text-decoration: none; border-radius: 6px; font-size: 11px; font-weight: 500;">
                            📧 sazahlie.sapar@moe.edu.my
                        </a>
                    </div>
                </div>
                <div style="display: flex; align-items: center; justify-content: center; height: 100%;">
                    {logo_html}
                </div>
            </div>
        </div>
        ''',
        unsafe_allow_html=True
    )

st.divider()

# -------------------------------------------------------------
# BAHAGIAN 1: MUAT NAIK FAIL PDF KALENDAR PERSEKOLAHAN KPM (WAJIB)
# -------------------------------------------------------------
st.header("1. 📄 Memuat Naik Fail PDF Kalendar Persekolahan KPM (WAJIB)")

uploaded_file = st.file_uploader(
    "Sila muat naik fail PDF Kalendar Persekolahan Rasmi KPM terlebih dahulu untuk menjana kalender",
    type=["pdf"],
    help="Sistem memerlukan fail PDF ini untuk membaca tarikh cuti perayaan dan penggal rasmi KPM."
)

if uploaded_file is None:
    st.warning("⚠️ **Langkah Wajib**: Sila muat naik fail PDF Kalendar Persekolahan KPM di atas sebelum sistem boleh menjana Kalender Akademik & eRPH.")
    st.stop()

pdf_text = ""
try:
    reader = pypdf.PdfReader(uploaded_file)
    for page in reader.pages:
        pdf_text += page.extract_text() + "\n"
    
    match_tahun = re.search(r"TAHUN\s+(\d{4})", pdf_text, re.IGNORECASE)
    tahun_detected = int(match_tahun.group(1)) if match_tahun else datetime.now().year + 1
    
    st.success(f"✅ Fail PDF '{uploaded_file.name}' berjaya dimuat naik & diekstrak untuk **Tahun {tahun_detected}**!")
except Exception as e:
    st.error(f"❌ Gagal memproses fail PDF: {e}")
    st.stop()

# -------------------------------------------------------------
# SIDEBAR TETAPAN PERSEKOLAHAN & MAKLUMAT HEADER
# -------------------------------------------------------------
st.sidebar.header("🏫 Maklumat Profil Sekolah")

nama_sekolah = st.sidebar.text_input(
    "Nama Sekolah",
    value="SK BANDAR SANDAKAN",
    help="Contoh: SK BANDAR SANDAKAN"
)

kod_sekolah = st.sidebar.text_input(
    "Kod Sekolah",
    value="XBA2004",
    help="Contoh: XBA2004"
)

nama_pentadbir = st.sidebar.text_input(
    "Nama Pengetua / Guru Besar",
    value="EN. SAZAHLIE BIN SAPAR",
    help="Contoh: EN. SAZAHLIE BIN SAPAR"
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Tetapan Asas Sekolah")

peringkat_sekolah = st.sidebar.selectbox(
    "Peringkat / Jenis Sekolah",
    [
        "Sekolah Rendah (SK / SJK)",
        "Sekolah Menengah (SMK / SABK)",
        "Prasekolah / KDNK",
        "Sekolah Agama Negeri (SAN / SRA)",
    ],
)

tahun = st.sidebar.number_input(
    "Tahun Persekolahan", min_value=2024, max_value=2035, value=tahun_detected
)

map_negeri_holidays = {
    "Johor": "JHR", "Kedah": "KDH", "Kelantan": "KTN", "Melaka": "MLK",
    "Negeri Sembilan": "NSN", "Pahang": "PHG", "Perak": "PRK", "Perlis": "PLS",
    "Pulau Pinang": "PNG", "Sabah": "SBH", "Sarawak": "SRW", "Selangor": "SGR",
    "Terengganu": "TRG", "W.P. Kuala Lumpur": "KUL", "W.P. Labuan": "LBN", "W.P. Putrajaya": "PJY",
}

senarai_negeri = list(map_negeri_holidays.keys())
negeri_pilihan = st.sidebar.selectbox("Pilih Negeri Sekolah", senarai_negeri, index=senarai_negeri.index("Sabah"))

if negeri_pilihan in ["Kedah", "Kelantan", "Terengganu"]:
    kumpulan_pilihan = "Kumpulan A (Ahad - Khamis)"
    is_kumpulan_a = True
else:
    kumpulan_pilihan = "Kumpulan B (Isnin - Jumaat)"
    is_kumpulan_a = False

st.sidebar.info(f"📌 **Kumpulan Ditetapkan Otomatik**: `{kumpulan_pilihan}`")

tarikh_mula = st.sidebar.date_input("Tarikh Mula Persekolahan (Minggu 1)", date(tahun, 1, 4))

sasaran_akademik = st.sidebar.number_input(
    "Maksimum Minggu Akademik",
    min_value=1,
    max_value=52,
    value=40,
    help="Bilangan minggu persekolahan rasmi bagi tujuan pengiraan eRPH.",
)

# -------------------------------------------------------------
# BAHAGIAN 2: PENETAPAN CUTI PENGGAL & PERAYAAN KPM AUTOMATIK
# -------------------------------------------------------------
if is_kumpulan_a:
    cuti_penggal_kpm = [
        (date(tahun, 3, 5), date(tahun, 3, 13), "CUTI PENGGAL 1"),
        (date(tahun, 5, 21), date(tahun, 6, 5), "CUTI PERTENGAHAN TAHUN"),
        (date(tahun, 8, 27), date(tahun, 9, 4), "CUTI PENGGAL 2"),
        (date(tahun, 12, 3), date(tahun, 12, 31), "CUTI AKHIR PERSEKOLAHAN"),
    ]
    cuti_perayaan_kpm = [
        (date(tahun, 2, 8), date(tahun, 2, 10), "CUTI PERAYAAN - TAHUN BAHARU CINA"),
        (date(tahun, 5, 16), date(tahun, 5, 19), "CUTI PERAYAAN - HARI RAYA AIDILADHA"),
        (date(tahun, 10, 27), date(tahun, 10, 28), "CUTI PERAYAAN - HARI DEEPAVALI"),
    ]
else:
    cuti_penggal_kpm = [
        (date(tahun, 3, 6), date(tahun, 3, 14), "CUTI PENGGAL 1"),
        (date(tahun, 5, 22), date(tahun, 6, 6), "CUTI PERTENGAHAN TAHUN"),
        (date(tahun, 8, 28), date(tahun, 9, 5), "CUTI PENGGAL 2"),
        (date(tahun, 12, 4), date(tahun, 12, 31), "CUTI AKHIR PERSEKOLAHAN"),
    ]
    cuti_perayaan_kpm = [
        (date(tahun, 2, 5), date(tahun, 2, 12), "CUTI PERAYAAN - TAHUN BAHARU CINA"),
        (date(tahun, 5, 18), date(tahun, 5, 19), "CUTI PERAYAAN - HARI RAYA AIDILADHA"),
        (date(tahun, 10, 27), date(tahun, 10, 29), "CUTI PERAYAAN - HARI DEEPAVALI"),
    ]

kod_subdiv = map_negeri_holidays.get(negeri_pilihan, "SBH")
cuti_google_cal = holidays.Malaysia(years=tahun, subdiv=kod_subdiv)
cuti_umum_list = [(dt, dt, f"{nama} ({dt.strftime('%d/%m/%Y')})") for dt, nama in sorted(cuti_google_cal.items())]

data_rows = []

for w in range(1, 53):
    mula_dt = tarikh_mula + timedelta(weeks=w - 1)
    tamat_dt = mula_dt + timedelta(days=4)

    sabtu_dt = (mula_dt - timedelta(days=1)) if is_kumpulan_a else (mula_dt + timedelta(days=5))
    sat_occ = (sabtu_dt.day - 1) // 7 + 1
    kokum = f"M{sat_occ}" if sat_occ in [2, 4] else ""

    default_jenis = "PdP / PdPr"
    catatan_list = []

    for c_mula, c_tamat, c_nama in cuti_penggal_kpm:
        if not (c_tamat < mula_dt or c_mula > tamat_dt):
            default_jenis = "Cuti Sekolah"
            catatan_list.append(f"{c_nama} ({c_mula.strftime('%d/%m/%Y')} - {c_tamat.strftime('%d/%m/%Y')})")
            break

    for p_mula, p_tamat, p_nama in cuti_perayaan_kpm:
        if not (p_tamat < mula_dt or p_mula > tamat_dt):
            catatan_list.append(f"{p_nama} ({p_mula.strftime('%d/%m/%Y')} - {p_tamat.strftime('%d/%m/%Y')})")

    for u_mula, u_tamat, u_nama in cuti_umum_list:
        if not (u_tamat < mula_dt or u_mula > tamat_dt):
            if u_nama not in catatan_list:
                catatan_list.append(u_nama)

    if w == 1 and default_jenis == "PdP / PdPr":
        default_jenis = "Bukan Akademik"
        catatan_suai = f"Minggu Suai Kenal / Transisi ({mula_dt.strftime('%d/%m/%Y')} - {tamat_dt.strftime('%d/%m/%Y')})"
        catatan_list.insert(0, catatan_suai)

    default_catatan = "\n".join(catatan_list)

    data_rows.append({
        "Minggu Kalendar": f"Minggu {w}",
        "Tarikh Mula": mula_dt.strftime("%d/%m/%Y"),
        "Tarikh Tamat": tamat_dt.strftime("%d/%m/%Y"),
        "Sabtu Kokum": kokum,
        "Jenis Minggu": default_jenis,
        "Catatan / Peristiwa": default_catatan,
        "_mula_dt": mula_dt,
        "_tamat_dt": tamat_dt,
    })

df_input = pd.DataFrame(data_rows)

# -------------------------------------------------------------
# PAPARAN DATA EDITOR
# -------------------------------------------------------------
st.header(f"2. 🗓️ Tetapan Kalender Mingguan ({negeri_pilihan} - {peringkat_sekolah})")

edited_df = st.data_editor(
    df_input[["Minggu Kalendar", "Tarikh Mula", "Tarikh Tamat", "Sabtu Kokum", "Jenis Minggu", "Catatan / Peristiwa"]],
    column_config={
        "Jenis Minggu": st.column_config.SelectboxColumn(
            "Jenis Minggu",
            options=["PdP / PdPr", "Cuti Sekolah", "Bukan Akademik"],
            required=True,
        ),
        "Catatan / Peristiwa": st.column_config.TextColumn("Catatan / Peristiwa"),
        "Minggu Kalendar": st.column_config.Column(disabled=True),
        "Tarikh Mula": st.column_config.Column(disabled=True),
        "Tarikh Tamat": st.column_config.Column(disabled=True),
        "Sabtu Kokum": st.column_config.Column(disabled=True),
    },
    use_container_width=True,
    num_rows="fixed",
    height=400,
)

# -------------------------------------------------------------
# BAHAGIAN 3: PENGIRAAN MINGGU AKADEMIK, BUKAN AKADEMIK & HARI PERSEKOLAHAN
# -------------------------------------------------------------
final_rows = []
running_academic_counter = 0
minggu_persekolahan_kpm = 0
minggu_bukan_akademik = 0
jumlah_hari_persekolahan = 0

for idx, row in edited_df.iterrows():
    mula_dt = df_input.loc[idx, "_mula_dt"]
    tamat_dt = df_input.loc[idx, "_tamat_dt"]
    jenis = row["Jenis Minggu"]

    if jenis != "Cuti Sekolah":
        minggu_persekolahan_kpm += 1

        hari_dalam_minggu = 5
        for u_mula, u_tamat, _ in cuti_umum_list:
            if mula_dt <= u_mula <= tamat_dt:
                hari_dalam_minggu -= 1
        
        for p_mula, p_tamat, _ in cuti_perayaan_kpm:
            if mula_dt <= p_mula <= tamat_dt:
                hari_dalam_minggu -= 1

        jumlah_hari_persekolahan += max(0, hari_dalam_minggu)

    if jenis == "PdP / PdPr":
        if running_academic_counter < sasaran_akademik:
            running_academic_counter += 1
            bil_akademik = str(running_academic_counter)

            dt_buka = mula_dt - timedelta(days=2 if is_kumpulan_a else 3)
            dt_tutup = mula_dt + timedelta(days=7)
            erph_txt = f"{dt_buka.strftime('%d/%m/%Y')} - {dt_tutup.strftime('%d/%m/%Y')}"
        else:
            bil_akademik = "-"
            erph_txt = ""
            dt_buka, dt_tutup = None, None
    elif jenis == "Bukan Akademik":
        minggu_bukan_akademik += 1
        bil_akademik = "-"
        erph_txt = ""
        dt_buka, dt_tutup = None, None
    else:
        bil_akademik = "-"
        erph_txt = ""
        dt_buka, dt_tutup = None, None

    final_rows.append({
        "Minggu Kalendar": row["Minggu Kalendar"],
        "Tarikh Mula": row["Tarikh Mula"],
        "Tarikh Tamat": row["Tarikh Tamat"],
        "Sabtu Kokum": row["Sabtu Kokum"],
        "Jenis Minggu": jenis,
        "Minggu Akademik": bil_akademik,
        "Catatan / Peristiwa": row["Catatan / Peristiwa"],
        "Tarikh Penghantaran eRPH": erph_txt,
        "_mula_dt": mula_dt,
        "_tamat_dt": tamat_dt,
        "_dt_tutup_erph": dt_tutup,
    })

baki_hari = jumlah_hari_persekolahan - 190
if baki_hari > 0:
    cuti_peristiwa_layak = min(4, baki_hari)
    status_cuti_txt = f"Layak {cuti_peristiwa_layak} Hari (Lebihan {baki_hari} Hari)"
    status_cuti = f"✅ Layak **{cuti_peristiwa_layak} Hari** (Lebihan {baki_hari} Hari)"
else:
    cuti_peristiwa_layak = 0
    status_cuti_txt = "Tidak Layak (<= 190 Hari)"
    status_cuti = "❌ Tidak Layak (<= 190 Hari)"

# -------------------------------------------------------------
# PAPARKAN ANALISIS PERSEKOLAHAN DI SIDEBAR
# -------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("📊 Analisis Persekolahan KPM")

st.sidebar.metric(
    label="Minggu Persekolahan KPM",
    value=f"{minggu_persekolahan_kpm} Minggu",
    delta=f"PdP: {running_academic_counter} M | Bukan Akademik: {minggu_bukan_akademik} M",
    help="Jumlah minggu sekolah dibuka (PdP/PdPr + Bukan Akademik)."
)

col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    st.metric("Minggu PdP", f"{running_academic_counter}")
with col_sb2:
    st.metric("Bukan Akademik", f"{minggu_bukan_akademik}")

st.sidebar.metric(
    label="Jumlah Hari Persekolahan Setahun",
    value=f"{jumlah_hari_persekolahan} Hari",
    delta=f"{baki_hari:+d} Hari dari sasaran 190",
    help="Jumlah hari persekolahan rasmi tidak termasuk cuti."
)

st.sidebar.markdown("**Kelayakan Cuti Peristiwa**:")
st.sidebar.info(f"{status_cuti}")

st.sidebar.markdown("---")
st.sidebar.markdown(
    '''
    <div style="text-align: center; font-size: 12px; color: #6c757d;">
        <p style="margin-bottom: 4px;"><b>Sistem Kalender Akademik & eRPH</b></p>
        <p style="margin: 0;">Hakcipta Terpelihara © <b>Sazahlie S.</b></p>
        <p style="margin-top: 4px;"><a href="mailto:sazahlie.sapar@moe.edu.my" style="color: #0d6efd; text-decoration: none;">sazahlie.sapar@moe.edu.my</a></p>
    </div>
    ''',
    unsafe_allow_html=True
)

df_final = pd.DataFrame(final_rows)

st.subheader("📊 Hasil Kalender Akhir")
st.success(
    f"**Sekolah**: `{nama_sekolah.upper()}` ({kod_sekolah.upper()}) | **Pengetua / GB**: `{nama_pentadbir.upper()}` | "
    f"**Negeri**: `{negeri_pilihan}` ({kumpulan_pilihan}) | **Jumlah Minggu Akademik**: `{running_academic_counter} / {sasaran_akademik}` Minggu | "
    f"**Bukan Akademik**: `{minggu_bukan_akademik}` Minggu"
)
st.dataframe(df_final.drop(columns=["_mula_dt", "_tamat_dt", "_dt_tutup_erph"]), use_container_width=True)

st.divider()

# -------------------------------------------------------------
# BAHAGIAN 4: EKSPORT KE EXCEL, PDF & ICS
# -------------------------------------------------------------
st.header("3. 📥 Muat Turun Kalendar & Cetakan Kemas (PDF)")

def generate_excel_bytes():
    """Penjanaan Excel mengikut susunan presisi tajuk & kad analisis contoh imej"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Kalender Sekolah {tahun}"

    font_title = Font(name="Segoe UI", size=13, bold=True, color="1F4E78")
    font_subtitle = Font(name="Segoe UI", size=11, bold=True, color="1F4E78")
    font_info = Font(name="Segoe UI", size=9, bold=False, color="595959")
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Row 1: Nama Sekolah & Kod Sekolah
    ws.merge_cells("A1:H1")
    ws["A1"] = f"{nama_sekolah.upper()} ({kod_sekolah.upper()})"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    # Row 2: Tajuk Utama Kalender
    ws.merge_cells("A2:H2")
    ws["A2"] = f"SISTEM KALENDER AKADEMIK & eRPH TAHUN {tahun}"
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = align_center

    # Row 3: Nama Pengetua / Guru Besar & Maklumat Kumpulan
    ws.merge_cells("A3:H3")
    ws["A3"] = f"PENGETUA / GURU BESAR: {nama_pentadbir.upper()} | NEGERI: {negeri_pilihan.upper()} ({kumpulan_pilihan.upper()})"
    ws["A3"].font = font_info
    ws["A3"].alignment = align_center

    ws.row_dimensions[4].height = 10

    # -------------------------------------------------------------
    # KAD ANALISIS EXCEL (PRESISI MENGIKUT GAMBAR CONTOH)
    # Row 5: Tajuk Kad (A5:H5)
    # Row 6: Minggu Persekolahan KPM: (A6:B6), [44 Minggu] (C6) | Minggu PdP (Akademik): (D6:E6), [40 / 40 Minggu] (F6) | Bukan Akademik: (G6), [1 Minggu] (H6)
    # Row 7: Jumlah Hari Persekolahan: (A7:B7), [208 Hari] (C7) | Kelayakan Cuti Peristiwa: (D7:E7), [Layak 4 Hari (Lebihan 18 Hari)] (F7:H7)
    # -------------------------------------------------------------
    card_hdr_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    card_body_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    
    card_hdr_font = Font(name="Segoe UI", size=9.5, bold=True, color="1F4E78")
    card_lbl_font = Font(name="Segoe UI", size=9, bold=True, color="262626")
    card_val_font = Font(name="Segoe UI", size=9.5, bold=True, color="0D6EFD")

    card_border = Border(
        left=Side(style="thin", color="B0C4DE"), right=Side(style="thin", color="B0C4DE"),
        top=Side(style="thin", color="B0C4DE"), bottom=Side(style="thin", color="B0C4DE")
    )

    # Row 5: Tajuk Kad
    ws.merge_cells("A5:H5")
    ws["A5"] = "📊 RINGKASAN ANALISIS PERSEKOLAHAN KPM"
    ws["A5"].font = card_hdr_font
    ws["A5"].fill = card_hdr_fill
    ws["A5"].alignment = align_center

    # Row 6
    ws.merge_cells("A6:B6")
    ws["A6"] = "Minggu Persekolahan KPM:"
    ws["A6"].font = card_lbl_font
    ws["A6"].alignment = align_right

    ws["C6"] = f"{minggu_persekolahan_kpm} Minggu"
    ws["C6"].font = card_val_font
    ws["C6"].alignment = align_left

    ws.merge_cells("D6:E6")
    ws["D6"] = "Minggu PdP (Akademik):"
    ws["D6"].font = card_lbl_font
    ws["D6"].alignment = align_right

    ws["F6"] = f"{running_academic_counter} / {sasaran_akademik} Minggu"
    ws["F6"].font = card_val_font
    ws["F6"].alignment = align_left

    ws["G6"] = "Bukan Akademik:"
    ws["G6"].font = card_lbl_font
    ws["G6"].alignment = align_right

    ws["H6"] = f"{minggu_bukan_akademik} Minggu"
    ws["H6"].font = card_val_font
    ws["H6"].alignment = align_left

    # Row 7
    ws.merge_cells("A7:B7")
    ws["A7"] = "Jumlah Hari Persekolahan:"
    ws["A7"].font = card_lbl_font
    ws["A7"].alignment = align_right

    ws["C7"] = f"{jumlah_hari_persekolahan} Hari"
    ws["C7"].font = card_val_font
    ws["C7"].alignment = align_left

    ws.merge_cells("D7:E7")
    ws["D7"] = "Kelayakan Cuti Peristiwa:"
    ws["D7"].font = card_lbl_font
    ws["D7"].alignment = align_right

    ws.merge_cells("F7:H7")
    ws["F7"] = f"{status_cuti_txt}"
    ws["F7"].font = card_val_font
    ws["F7"].alignment = align_left

    for r in range(5, 8):
        ws.row_dimensions[r].height = 20
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = card_border
            if r in [6, 7]:
                cell.fill = card_body_fill
            cell.protection = Protection(locked=False)

    # -------------------------------------------------------------
    # HEADER JADUAL TAKWIM (BARIS 9)
    # -------------------------------------------------------------
    headers = [
        "Minggu Kalendar", "Tarikh Mula", "Tarikh Tamat", "Sabtu Kokum",
        "Jenis Minggu", "Minggu Akademik", "Catatan / Peristiwa", "Tarikh Penghantaran eRPH"
    ]

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

    ws.append([]) # Row 8
    ws.append(headers) # Row 9 Header Jadual

    for cell in ws[9]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    cuti_sekolah_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    cuti_sekolah_font = Font(name="Segoe UI", size=9.5, bold=True, color="B25900")

    cuti_perayaan_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    cuti_perayaan_font = Font(name="Segoe UI", size=9.5, bold=True, color="1F4E78")

    bukan_akademik_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    bukan_akademik_font = Font(name="Segoe UI", size=9.5, bold=True, color="C00080")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
    )

    for r_idx, r in enumerate(final_rows, start=10):
        catatan_text = str(r["Catatan / Peristiwa"])
        lines = catatan_text.split("\n")

        ws.row_dimensions[r_idx].height = max(22, len(lines) * 18)

        row_vals = [
            r["Minggu Kalendar"], r["Tarikh Mula"], r["Tarikh Tamat"], r["Sabtu Kokum"],
            r["Jenis Minggu"], r["Minggu Akademik"], catatan_text, r["Tarikh Penghantaran eRPH"]
        ]
        ws.append(row_vals)

        jenis_str = str(r["Jenis Minggu"])
        catatan_str = catatan_text.upper()

        for c_idx in range(1, 9):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.font = Font(name="Segoe UI", size=9.5, italic=False)
            cell.alignment = align_center if c_idx != 7 else align_left
            cell.protection = Protection(locked=False)

            if jenis_str == "Cuti Sekolah":
                cell.fill = cuti_sekolah_fill
                if c_idx == 5:
                    cell.font = cuti_sekolah_font
            elif jenis_str == "Bukan Akademik":
                cell.fill = bukan_akademik_fill
                if c_idx == 5:
                    cell.font = bukan_akademik_font
            elif "CUTI PERAYAAN" in catatan_str or any(k in catatan_str for k in ["HARI RAYA", "DEEPAVALI", "TAHUN BAHARU"]):
                cell.fill = cuti_perayaan_fill
                if c_idx in [5, 7]:
                    cell.font = cuti_perayaan_font

    for r in range(1, 10):
        for col in range(1, 9):
            ws.cell(row=r, column=col).protection = Protection(locked=False)

    last_row = len(final_rows) + 11
    ws.merge_cells(start_row=last_row, start_column=1, end_row=last_row, end_column=8)
    footer_cell = ws.cell(
        row=last_row,
        column=1,
        value=f"🔒 Hakcipta Terpelihara © {tahun} Sazahlie S. (sazahlie.sapar@moe.edu.my) - Sistem Kalender Akademik & eRPH"
    )
    footer_cell.font = Font(name="Segoe UI", size=9, italic=False, color="0D6EFD", bold=True)
    footer_cell.alignment = align_center
    footer_cell.hyperlink = "mailto:sazahlie.sapar@moe.edu.my"

    footer_cell.protection = Protection(locked=True)

    ws.protection.sheet = True
    ws.protection.password = "Sazahlie@2027"
    ws.protection.enable()

    for col in ws.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        max_len = 0
        
        for cell in col:
            if cell.row >= 9 and cell.value and cell.row < last_row:
                cell_lines = str(cell.value).split("\n")
                for line in cell_lines:
                    if len(line) > max_len:
                        max_len = len(line)

        if col_letter == "G":
            ws.column_dimensions[col_letter].width = min(max(max_len + 5, 45), 85)
        else:
            ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def generate_pdf_bytes():
    """Penjanaan PDF Cetakan Kemas (A4 Landscape) dengan Kad Analisis Presisi"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25,
    )

    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=13,
        textColor=colors.HexColor('#1F4E78'), alignment=1, spaceAfter=2,
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10,
        textColor=colors.HexColor('#1F4E78'), alignment=1, spaceAfter=2,
    )
    info_style = ParagraphStyle(
        'DocInfo', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5,
        textColor=colors.HexColor('#595959'), alignment=1, spaceAfter=8,
    )

    card_lbl = ParagraphStyle('CardLbl', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.HexColor('#262626'), alignment=2)
    card_val = ParagraphStyle('CardVal', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#0D6EFD'), alignment=0)
    card_hdr = ParagraphStyle('CardHdr', fontName='Helvetica-Bold', fontSize=8.5, textColor=colors.HexColor('#1F4E78'), alignment=1)

    cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9, alignment=0)
    cell_center_style = ParagraphStyle('TableCellCenter', parent=styles['Normal'], fontName='Helvetica', fontSize=7.5, leading=9, alignment=1)
    cell_header_style = ParagraphStyle('TableHeaderCell', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)

    story = []

    story.append(Paragraph(f"{nama_sekolah.upper()} ({kod_sekolah.upper()})", title_style))
    story.append(Paragraph(f"SISTEM KALENDER AKADEMIK & eRPH TAHUN {tahun}", subtitle_style))
    story.append(Paragraph(f"PENGETUA / GURU BESAR: {nama_pentadbir.upper()} | NEGERI: {negeri_pilihan.upper()} ({kumpulan_pilihan.upper()})", info_style))

    card_pdf_data = [
        [Paragraph("📊 RINGKASAN ANALISIS PERSEKOLAHAN KPM", card_hdr), "", "", "", "", ""],
        [
            Paragraph("Minggu Persekolahan KPM:", card_lbl), Paragraph(f"{minggu_persekolahan_kpm} Minggu", card_val),
            Paragraph("Minggu PdP (Akademik):", card_lbl), Paragraph(f"{running_academic_counter} / {sasaran_akademik} Minggu", card_val),
            Paragraph("Bukan Akademik:", card_lbl), Paragraph(f"{minggu_bukan_akademik} Minggu", card_val)
        ],
        [
            Paragraph("Jumlah Hari Persekolahan:", card_lbl), Paragraph(f"{jumlah_hari_persekolahan} Hari", card_val),
            Paragraph("Kelayakan Cuti Peristiwa:", card_lbl), Paragraph(f"{status_cuti_txt}", card_val), "", ""
        ]
    ]

    card_col_widths = [135, 95, 140, 150, 110, 160]
    pdf_card_table = Table(card_pdf_data, colWidths=card_col_widths)
    pdf_card_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (5, 0)),
        ('SPAN', (3, 2), (5, 2)),
        ('BACKGROUND', (0, 0), (5, 0), colors.HexColor('#DDEBF7')),
        ('BACKGROUND', (0, 1), (5, 2), colors.HexColor('#F9FAFB')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#B0C4DE')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(pdf_card_table)
    story.append(Paragraph("<br/>", info_style))

    headers = [
        "Minggu Kalendar", "Tarikh Mula", "Tarikh Tamat", "Sabtu Kokum",
        "Jenis Minggu", "Minggu Akademik", "Catatan / Peristiwa", "Tarikh Penghantaran eRPH"
    ]
    
    table_data = [[Paragraph(h, cell_header_style) for h in headers]]

    for r in final_rows:
        jenis_str = str(r["Jenis Minggu"])
        catatan_formatted = str(r["Catatan / Peristiwa"]).replace('\n', '<br/>')

        if jenis_str == "Cuti Sekolah":
            c_type_p = Paragraph(f"<b><font color='#B25900'>{jenis_str}</font></b>", cell_center_style)
        elif jenis_str == "Bukan Akademik":
            c_type_p = Paragraph(f"<b><font color='#C00080'>{jenis_str}</font></b>", cell_center_style)
        else:
            c_type_p = Paragraph(jenis_str, cell_center_style)

        row_cells = [
            Paragraph(str(r["Minggu Kalendar"]), cell_center_style),
            Paragraph(str(r["Tarikh Mula"]), cell_center_style),
            Paragraph(str(r["Tarikh Tamat"]), cell_center_style),
            Paragraph(str(r["Sabtu Kokum"]), cell_center_style),
            c_type_p,
            Paragraph(str(r["Minggu Akademik"]), cell_center_style),
            Paragraph(catatan_formatted, cell_style),
            Paragraph(str(r["Tarikh Penghantaran eRPH"]), cell_center_style),
        ]
        table_data.append(row_cells)

    col_widths = [70, 55, 55, 60, 85, 65, 260, 140]
    pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D9D9D9')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
    ]

    for idx, r in enumerate(final_rows, start=1):
        jenis_str = str(r["Jenis Minggu"])
        catatan_str = str(r["Catatan / Peristiwa"]).upper()

        if jenis_str == "Cuti Sekolah":
            t_style.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor('#FFF2CC')))
        elif jenis_str == "Bukan Akademik":
            t_style.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor('#FCE4D6')))
        elif "CUTI PERAYAAN" in catatan_str or any(k in catatan_str for k in ["HARI RAYA", "DEEPAVALI", "TAHUN BAHARU"]):
            t_style.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor('#DDEBF7')))

    pdf_table.setStyle(TableStyle(t_style))
    story.append(pdf_table)

    def add_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.HexColor('#595959'))
        canvas.drawCentredString(
            landscape(A4)[0] / 2.0, 15,
            f"🔒 Hakcipta Terpelihara © {tahun} Sazahlie S. (sazahlie.sapar@moe.edu.my) - Sistem Kalender Akademik & eRPH"
        )
        canvas.restoreState()

    doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
    
    buffer.seek(0)
    return buffer.getvalue()

def generate_ics_text():
    """Penjanaan Fail Kalendar Digital (.ics)"""
    ics_lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        f"PRODID:-//{nama_sekolah}//Sistem Kalender Akademik & eRPH Sazahlie S.//MY",
        "CALSCALE:GREGORIAN", "METHOD:PUBLISH"
    ]
    now_str = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    for r in final_rows:
        mula_str = r["_mula_dt"].strftime("%Y%m%d")
        tamat_str = (r["_tamat_dt"] + timedelta(days=1)).strftime("%Y%m%d")

        ics_lines.extend([
            "BEGIN:VEVENT",
            f"UID:kalender-{mula_str}-{r['Minggu Kalendar'].replace(' ', '')}@{kod_sekolah.lower()}",
            f"DTSTAMP:{now_str}",
            f"DTSTART;VALUE=DATE:{mula_str}",
            f"DTEND;VALUE=DATE:{tamat_str}",
            f"SUMMARY:{r['Minggu Kalendar']} - {r['Jenis Minggu']} (M{r['Minggu Akademik']})",
            f"DESCRIPTION:Sekolah: {nama_sekolah}\nPengetua/GB: {nama_pentadbir}\nCatatan: {r['Catatan / Peristiwa']}\nHakcipta © Sazahlie S. (sazahlie.sapar@moe.edu.my)",
            "END:VEVENT"
        ])

    ics_lines.append("END:VCALENDAR")
    return "\r\n".join(ics_lines)

# -------------------------------------------------------------
# BUTANG MUAT TURUN (EXCEL, PDF & ICS)
# -------------------------------------------------------------
col_btn1, col_btn2, col_btn3 = st.columns(3)

with col_btn1:
    st.download_button(
        label="📥 Muat Turun Fail Excel (.xlsx)",
        data=generate_excel_bytes(),
        file_name=f"Kalender_{kod_sekolah.upper()}_{tahun}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_btn2:
    st.download_button(
        label="🖨️ Muat Turun Cetakan Kemas (.pdf)",
        data=generate_pdf_bytes(),
        file_name=f"Kalender_{kod_sekolah.upper()}_{tahun}.pdf",
        mime="application/pdf",
        help="Format PDF saiz A4 Lanskap yang sedia dicetak terus dengan susunan rasmi yang kemas.",
        use_container_width=True,
    )

with col_btn3:
    st.download_button(
        label="📲 Muat Turun Kalendar Digital (.ics)",
        data=generate_ics_text(),
        file_name=f"Kalender_{kod_sekolah.upper()}_{tahun}.ics",
        mime="text/calendar",
        use_container_width=True,
    )

# Footer Laman Web
st.markdown("---")
st.markdown(
    '''
    <div style="text-align: center; font-size: 13px; color: #6c757d; padding: 10px 0;">
        🔒 <b>Hakcipta Terpelihara © Sazahlie S.</b> | 📧 E-mel: <a href="mailto:sazahlie.sapar@moe.edu.my" style="color: #0d6efd; text-decoration: none; font-weight: 600;">sazahlie.sapar@moe.edu.my</a>
    </div>
    ''',
    unsafe_allow_html=True
)