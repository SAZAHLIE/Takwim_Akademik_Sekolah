import io
import os
import re
from datetime import date, datetime, timedelta
import holidays
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
import pandas as pd
import pypdf
import streamlit as st

# Konfigurasi Laman Web Streamlit
st.set_page_config(
    page_title="Sistem Automasi Takwim Sekolah Global KPM",
    layout="wide",
    page_icon="🏫",
)

# -------------------------------------------------------------
# LOGO & HEADER APLIKASI
# -------------------------------------------------------------
possible_logos = [
    "SEKOLAH",
    "SEKOLAH.png",
    "SEKOLAH.jpg",
    "input_file_0.png",
    "Ori.jpg",
    "logo.png",
    "logo.jpg",
]
logo_file_found = None

for f in possible_logos:
    if os.path.exists(f):
        logo_file_found = f
        break

if not logo_file_found:
    for f in os.listdir("."):
        if f.lower().endswith((".png", ".jpg", ".jpeg")) and not f.startswith("logo_excel"):
            logo_file_found = f
            break

col_logo, col_title = st.columns([1, 6])

with col_logo:
    if logo_file_found:
        st.image(logo_file_found, width=110)
    else:
        st.write("🏫")

with col_title:
    st.title("🏫 Sistem Automasi Takwim Sekolah & eRPH KPM")
    st.markdown("### **🔒 Hak Cipta Terpelihara © SAZAHLIE S.**")
    st.markdown("📧 **E-mel Perhubungan**: `sazahlie.sapar@moe.edu.my`")
    st.markdown(
        "Aplikasi pintar dan fleksibel untuk membaca fail PDF Kalendar Akademik KPM, "
        "menyemak cuti umum mengikut negeri via **Google Calendar / Pustaka Holidays**, "
        "mengurus Jenis Minggu, serta mengira Minggu Akademik & eRPH secara automatik."
    )

st.divider()

# -------------------------------------------------------------
# BAHAGIAN 1: Memuat Naik & Auto-Parsing Fail PDF Takwim KPM
# -------------------------------------------------------------
st.header("1. 📄 Memuat Naik Fail PDF Takwim KPM")
uploaded_file = st.file_uploader(
    "Muat naik fail PDF Takwim Rasmi KPM (Contoh: 2027.pdf, 2028.pdf)",
    type=["pdf"],
)

pdf_text = ""
tahun_detected = 2027

if uploaded_file is not None:
    try:
        reader = pypdf.PdfReader(uploaded_file)
        for page in reader.pages:
            pdf_text += page.extract_text() + "\n"

        match_tahun = re.search(r"TAHUN\s+(\d{4})", pdf_text, re.IGNORECASE)
        if match_tahun:
            tahun_detected = int(match_tahun.group(1))

        st.success(
            f"✅ Fail PDF '{uploaded_file.name}' berjaya dimuat naik & diekstrak!"
        )
    except Exception as e:
        st.error(f"Gagal memproses fail PDF: {e}")
else:
    st.warning(
        "⚠️ **Perhatian**: Sila muat naik fail PDF Takwim Rasmi KPM terlebih"
        " dahulu di atas untuk menjana takwim persekolahan."
    )
    st.stop()

# -------------------------------------------------------------
# SIDEBAR SETTINGS
# -------------------------------------------------------------
st.sidebar.header("⚙️ Tetapan Asas Persekolahan")
tahun = st.sidebar.number_input(
    "Tahun Persekolahan", min_value=2024, max_value=2035, value=tahun_detected
)

map_negeri_holidays = {
    "Johor": "JHR",
    "Kedah": "KDH",
    "Kelantan": "KTN",
    "Melaka": "MLK",
    "Negeri Sembilan": "NSN",
    "Pahang": "PHG",
    "Perak": "PRK",
    "Perlis": "PLS",
    "Pulau Pinang": "PNG",
    "Sabah": "SBH",
    "Sarawak": "SRW",
    "Selangor": "SGR",
    "Terengganu": "TRG",
    "W.P. Kuala Lumpur": "KUL",
    "W.P. Labuan": "LBN",
    "W.P. Putrajaya": "PJY",
}

senarai_negeri = list(map_negeri_holidays.keys())

negeri_pilihan = st.sidebar.selectbox(
    "Pilih Negeri Sekolah",
    senarai_negeri,
    index=senarai_negeri.index("Sabah"),
)

if negeri_pilihan in ["Kedah", "Kelantan", "Terengganu"]:
    kumpulan_pilihan = "Kumpulan A (Ahad - Khamis)"
    is_kumpulan_a = True
else:
    kumpulan_pilihan = "Kumpulan B (Isnin - Jumaat)"
    is_kumpulan_a = False

st.sidebar.info(f"📌 **Kumpulan Ditetapkan Otomatik**: `{kumpulan_pilihan}`")

tarikh_mula = st.sidebar.date_input(
    "Tarikh Mula Persekolahan (Minggu 1)", date(tahun, 1, 4)
)

sasaran_akademik = st.sidebar.number_input(
    "Sasaran Maksimum Minggu Akademik",
    min_value=1,
    max_value=52,
    value=40,
    help="Taipkan bilangan minggu akademik sasaran (Terbatas kepada minggu kalendar)",
)

# Sidebar Branding
st.sidebar.markdown("---")
st.sidebar.markdown(
    "🔒 **Hak Cipta Terpelihara © SAZAHLIE S.**\n\n📧"
    " `sazahlie.sapar@moe.edu.my`\n\n*Fail Excel dijana mengandungi"
    " Perlindungan Digital (Protected Sheet) & Header/Footer Cetakan Terkunci.*"
)

st.divider()

# -------------------------------------------------------------
# BAHAGIAN 2: Data Cuti Mengikut Negeri
# -------------------------------------------------------------
st.header(f"2. 🗓️ Tetapan Jenis Minggu & Catatan Takwim ({negeri_pilihan})")

if is_kumpulan_a:
    cuti_ranges = [
        (
            date(tahun, 3, 5),
            date(tahun, 3, 13),
            f"Cuti Penggal 1 ({date(tahun, 3, 5).strftime('%d/%m/%Y')} - {date(tahun, 3, 13).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 5, 21),
            date(tahun, 6, 5),
            f"Cuti Pertengahan Tahun ({date(tahun, 5, 21).strftime('%d/%m/%Y')} - {date(tahun, 6, 5).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 8, 27),
            date(tahun, 9, 4),
            f"Cuti Penggal 2 ({date(tahun, 8, 27).strftime('%d/%m/%Y')} - {date(tahun, 9, 4).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 12, 3),
            date(tahun, 12, 31),
            f"Cuti Akhir Persekolahan ({date(tahun, 12, 3).strftime('%d/%m/%Y')} - {date(tahun, 12, 31).strftime('%d/%m/%Y')})",
        ),
    ]
    cuti_perayaan_kpm = [
        (
            date(tahun, 2, 8),
            date(tahun, 2, 10),
            f"Cuti Perayaan KPM - Tahun Baharu Cina ({date(tahun, 2, 8).strftime('%d/%m/%Y')} - {date(tahun, 2, 10).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 5, 16),
            date(tahun, 5, 19),
            f"Cuti Perayaan KPM - Hari Raya Aidiladha ({date(tahun, 5, 16).strftime('%d/%m/%Y')} - {date(tahun, 5, 19).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 10, 27),
            date(tahun, 10, 28),
            f"Cuti Perayaan KPM - Hari Deepavali ({date(tahun, 10, 27).strftime('%d/%m/%Y')} - {date(tahun, 10, 28).strftime('%d/%m/%Y')})",
        ),
    ]
else:
    cuti_ranges = [
        (
            date(tahun, 3, 6),
            date(tahun, 3, 14),
            f"Cuti Penggal 1 ({date(tahun, 3, 6).strftime('%d/%m/%Y')} - {date(tahun, 3, 14).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 5, 22),
            date(tahun, 6, 6),
            f"Cuti Pertengahan Tahun ({date(tahun, 5, 22).strftime('%d/%m/%Y')} - {date(tahun, 6, 6).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 8, 28),
            date(tahun, 9, 5),
            f"Cuti Penggal 2 ({date(tahun, 8, 28).strftime('%d/%m/%Y')} - {date(tahun, 9, 5).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 12, 4),
            date(tahun, 12, 31),
            f"Cuti Akhir Persekolahan ({date(tahun, 12, 4).strftime('%d/%m/%Y')} - {date(tahun, 12, 31).strftime('%d/%m/%Y')})",
        ),
    ]
    cuti_perayaan_kpm = [
        (
            date(tahun, 2, 5),
            date(tahun, 2, 12),
            f"Cuti Perayaan KPM - Tahun Baharu Cina ({date(tahun, 2, 5).strftime('%d/%m/%Y')} - {date(tahun, 2, 12).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 5, 18),
            date(tahun, 5, 19),
            f"Cuti Perayaan KPM - Hari Raya Aidiladha ({date(tahun, 5, 18).strftime('%d/%m/%Y')} - {date(tahun, 5, 19).strftime('%d/%m/%Y')})",
        ),
        (
            date(tahun, 10, 27),
            date(tahun, 10, 29),
            f"Cuti Perayaan KPM - Hari Deepavali ({date(tahun, 10, 27).strftime('%d/%m/%Y')} - {date(tahun, 10, 29).strftime('%d/%m/%Y')})",
        ),
    ]

kod_subdiv = map_negeri_holidays.get(negeri_pilihan, "SBH")
cuti_google_cal = holidays.Malaysia(years=tahun, subdiv=kod_subdiv)

cuti_umum_list = []
for dt_cuti, nama_cuti in sorted(cuti_google_cal.items()):
    cuti_umum_list.append(
        (dt_cuti, dt_cuti, f"{nama_cuti} ({dt_cuti.strftime('%d/%m/%Y')})")
    )

data_rows = []

for w in range(1, 53):
    mula_dt = tarikh_mula + timedelta(weeks=w - 1)
    tamat_dt = mula_dt + timedelta(days=4)

    sabtu_dt = (
        (mula_dt - timedelta(days=1))
        if is_kumpulan_a
        else (mula_dt + timedelta(days=5))
    )
    sat_occ = (sabtu_dt.day - 1) // 7 + 1
    kokum = f"M{sat_occ}" if sat_occ in [2, 4] else ""

    default_jenis = "PdP / PdPr"
    default_catatan = ""

    for c_mula, c_tamat, c_nama in cuti_ranges:
        if c_mula <= mula_dt <= c_tamat or c_mula <= tamat_dt <= c_tamat:
            default_jenis = "Cuti Sekolah"
            default_catatan = c_nama
            break

    catatan_minggu_ini = []

    for p_mula, p_tamat, p_nama in cuti_perayaan_kpm:
        if not (p_tamat < mula_dt or p_mula > tamat_dt):
            catatan_minggu_ini.append(p_nama)

    for u_mula, u_tamat, u_nama in cuti_umum_list:
        if not (u_tamat < mula_dt or u_mula > tamat_dt):
            if u_nama not in catatan_minggu_ini:
                catatan_minggu_ini.append(u_nama)

    if catatan_minggu_ini:
        formatted_events = "\n".join(catatan_minggu_ini)
        if default_catatan:
            default_catatan = f"{default_catatan}\n{formatted_events}"
        else:
            default_catatan = formatted_events

    if w == 1 and default_jenis == "PdP / PdPr":
        default_jenis = "Bukan Akademik"
        catatan_suai_kenal = (
            f"Minggu Suai Kenal ({mula_dt.strftime('%d/%m/%Y')} -"
            f" {tamat_dt.strftime('%d/%m/%Y')})"
        )
        default_catatan = (
            f"{catatan_suai_kenal}\n{default_catatan}"
            if default_catatan
            else catatan_suai_kenal
        )

    data_rows.append({
        "Minggu Kalendar": f"Minggu {w}",
        "Tarikh Mula": mula_dt.strftime("%d/%m/%Y"),
        "Tarikh Tamat": tamat_dt.strftime("%d/%m/%Y"),
        "Sabtu Kokum": kokum,
        "Jenis Minggu": default_jenis,
        "Catatan / Peristiwa": default_catatan,
        "_mula_dt": mula_dt,
    })

df_input = pd.DataFrame(data_rows)

column_config_setup = {
    "Jenis Minggu": st.column_config.SelectboxColumn(
        "Jenis Minggu (Pilih Drop-Down)",
        options=["PdP / PdPr", "Cuti Sekolah", "Bukan Akademik"],
        required=True,
    ),
    "Catatan / Peristiwa": st.column_config.TextColumn(
        "Catatan / Peristiwa",
        help="Cuti & peristiwa dipisahkan mengikut baris baharu",
    ),
    "Minggu Kalendar": st.column_config.Column(disabled=True),
    "Tarikh Mula": st.column_config.Column(disabled=True),
    "Tarikh Tamat": st.column_config.Column(disabled=True),
    "Sabtu Kokum": st.column_config.Column(disabled=True),
}

edited_df = st.data_editor(
    df_input[[
        "Minggu Kalendar",
        "Tarikh Mula",
        "Tarikh Tamat",
        "Sabtu Kokum",
        "Jenis Minggu",
        "Catatan / Peristiwa",
    ]],
    column_config=column_config_setup,
    use_container_width=True,
    num_rows="fixed",
    height=420,
)

# -------------------------------------------------------------
# BAHAGIAN 3: Pengiraan Otomatik Berurutan Tepat (Strict 1, 2, 3...)
# -------------------------------------------------------------
final_rows = []
running_academic_counter = 0

for idx, row in edited_df.iterrows():
    mula_dt = df_input.loc[idx, "_mula_dt"]
    jenis = row["Jenis Minggu"]

    if jenis == "PdP / PdPr":
        if running_academic_counter < sasaran_akademik:
            running_academic_counter += 1
            bil_akademik = str(running_academic_counter)

            if is_kumpulan_a:
                buka_erph = (mula_dt - timedelta(days=2)).strftime("%d/%m/%Y")
                tutup_erph = (mula_dt + timedelta(days=7)).strftime("%d/%m/%Y")
            else:
                buka_erph = (mula_dt - timedelta(days=3)).strftime("%d/%m/%Y")
                tutup_erph = (mula_dt + timedelta(days=7)).strftime("%d/%m/%Y")
            erph_txt = f"{buka_erph} - {tutup_erph}"
        else:
            bil_akademik = "-"
            erph_txt = ""
    else:
        bil_akademik = "-"
        erph_txt = ""

    final_rows.append({
        "Minggu Kalendar": row["Minggu Kalendar"],
        "Tarikh Mula": row["Tarikh Mula"],
        "Tarikh Tamat": row["Tarikh Tamat"],
        "Sabtu Kokum": row["Sabtu Kokum"],
        "Jenis Minggu": jenis,
        "Minggu Akademik": bil_akademik,
        "Catatan / Peristiwa": row["Catatan / Peristiwa"],
        "Tarikh Penghantaran eRPH": erph_txt,
    })

df_final = pd.DataFrame(final_rows)

st.subheader("📊 Hasil Takwim Akhir")
st.success(
    f"**Negeri Terpilih**: `{negeri_pilihan}` | **Kumpulan**: `{kumpulan_pilihan}`"
    f" | **Jumlah Minggu Akademik Dikira**: `{running_academic_counter} / {sasaran_akademik}` **Minggu**"
)
st.dataframe(df_final, use_container_width=True)

st.divider()

# Footer Web Branding
st.caption(
    "🔒 © 2026 SAZAHLIE S. (📧 sazahlie.sapar@moe.edu.my). Hak Cipta"
    " Terpelihara. Sistem & Output Diterapkan Perlindungan Digital."
)

# -------------------------------------------------------------
# BAHAGIAN 4: Eksport ke Excel (.xlsx) dengan Watermark Logo & Perlindungan Hak Cipta
# -------------------------------------------------------------
st.header("3. 📥 Muat Turun Fail Excel Takwim (.xlsx)")


def generate_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Takwim Sekolah {tahun}"

    if logo_file_found and os.path.exists(logo_file_found):
        try:
            img_temp = Image.open(logo_file_found)
            img_temp.thumbnail((65, 65))
            img_temp.save("logo_excel_temp.png")
            excel_img = OpenpyxlImage("logo_excel_temp.png")
            ws.add_image(excel_img, "A1")
        except Exception:
            pass

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20

    ws.merge_cells("B1:H1")
    ws["B1"] = (
        f"RANCANGAN TAKWIM PERSEKOLAHAN & PENGHANTARAN eRPH TAHUN {tahun}"
    )
    ws["B1"].font = Font(name="Segoe UI", size=14, bold=True, color="1F4E78")
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B2:H2")
    ws["B2"] = (
        f"Negeri: {negeri_pilihan} ({kumpulan_pilihan}) | Jumlah Minggu Akademik:"
        f" {running_academic_counter} / {sasaran_akademik} Minggu"
    )
    ws["B2"].font = Font(name="Segoe UI", size=10, italic=True, color="595959")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B3:H3")
    ws["B3"] = (
        f"🔒 HAK CIPTA TERPELIHARA © {tahun} SAZAHLIE S. (sazahlie.sapar@moe.edu.my) | DIKUNCI / DIGUNAKAN SECARA RASMI"
    )
    ws["B3"].font = Font(name="Segoe UI", size=9, bold=True, color="1F4E78")
    ws["B3"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        "Minggu Kalendar",
        "Tarikh Mula",
        "Tarikh Tamat",
        "Sabtu Kokum",
        "Jenis Minggu",
        "Minggu Akademik",
        "Catatan / Peristiwa",
        "Tarikh Penghantaran eRPH",
    ]

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

    ws.row_dimensions[5].height = 26
    for col_num, h_title in enumerate(headers, 1):
        c = ws.cell(row=5, column=col_num, value=h_title)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    cuti_fill = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )
    bukan_ak_fill = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )
    alt_fill = PatternFill(
        start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
    )

    max_catatan_length = len("Catatan / Peristiwa")

    for r_idx, r_data in enumerate(final_rows, 6):
        catatan_text = str(r_data["Catatan / Peristiwa"])

        lines = catatan_text.split("\n")
        for line in lines:
            if len(line) > max_catatan_length:
                max_catatan_length = len(line)

        num_lines = len(lines)
        ws.row_dimensions[r_idx].height = max(22, num_lines * 18)

        row_vals = [
            r_data["Minggu Kalendar"],
            r_data["Tarikh Mula"],
            r_data["Tarikh Tamat"],
            r_data["Sabtu Kokum"],
            r_data["Jenis Minggu"],
            r_data["Minggu Akademik"],
            catatan_text,
            r_data["Tarikh Penghantaran eRPH"],
        ]

        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Segoe UI", size=9.5)
            cell.border = thin_border

            # KUNCI SEL HAK CIPTA (Protection Lock)
            cell.protection = openpyxl.styles.Protection(locked=True)

            if c_idx in [1, 2, 3, 4, 5, 6]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

            if r_data["Jenis Minggu"] == "Cuti Sekolah":
                cell.fill = cuti_fill
                if c_idx == 5:
                    cell.font = Font(
                        name="Segoe UI", size=9.5, bold=True, color="C00000"
                    )
            elif r_data["Jenis Minggu"] == "Bukan Akademik":
                cell.fill = bukan_ak_fill
                if c_idx == 5:
                    cell.font = Font(
                        name="Segoe UI", size=9.5, bold=True, color="B25900"
                    )
            else:
                if r_idx % 2 == 0:
                    cell.fill = alt_fill

    footer_row = len(final_rows) + 7
    ws.merge_cells(
        start_row=footer_row, start_column=1, end_row=footer_row, end_column=8
    )
    footer_cell = ws.cell(
        row=footer_row,
        column=1,
        value=(
            f"🔒 HAK CIPTA TERPELIHARA © {tahun} SAZAHLIE S. (sazahlie.sapar@moe.edu.my) - TIDAK BOLEH DIUBAH ATAU DIPINDAH TANPA KEBENARAN."
        ),
    )
    footer_cell.font = Font(name="Segoe UI", size=8.5, bold=True, color="1F4E78")
    footer_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.protection.sheet = True
    ws.protection.objects = True
    ws.protection.scenarios = True
    ws.protection.password = "SAZAHLIE2026"
    ws.protection.enable()

    ws.HeaderFooter.oddFooter.center.text = (
        f"© {tahun} SAZAHLIE S. (sazahlie.sapar@moe.edu.my) - Hak Cipta Terpelihara"
    )
    ws.HeaderFooter.oddFooter.center.size = 8

    dynamic_g_width = min(max(max_catatan_length + 4, 50), 80)

    column_widths = {
        "A": 16,
        "B": 13,
        "C": 13,
        "D": 16,
        "E": 22,
        "F": 18,
        "G": dynamic_g_width,
        "H": 28,
    }
    for col, w in column_widths.items():
        ws.column_dimensions[col].width = w

    output_file = f"Takwim_Sekolah_{negeri_pilihan}_{tahun}.xlsx"
    wb.save(output_file)
    return output_file


excel_filename = generate_excel()

with open(excel_filename, "rb") as f:
    st.download_button(
        label="📥 Klik Untuk Muat Turun Fail Excel (.xlsx)",
        data=f,
        file_name=f"Takwim_Sekolah_{negeri_pilihan}_{tahun}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )