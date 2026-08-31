import streamlit as st
import pandas as pd
import simplekml
from simplekml import Snippet
import zipfile
from datetime import datetime
import requests
import base64


today = datetime.now().strftime("%d %b %Y")

st.set_page_config(
    page_title="Validasi ODP Tools",
    page_icon="📍",
    layout="wide"
)

st.markdown("""
<style>

.block-container{
    padding-top:25px;
    max-width:1400px;
}

section[data-testid="stSidebar"]{
    background:#f8fafc;
    border-right:1px solid #e5e7eb;
}

div[data-testid="stVerticalBlock"]>div{
    gap:1rem;
}

.card{
background:white;
border-radius:14px;
padding:17.4px 15px;
border:1px solid #edf2f7;
box-shadow:0 3px 10px rgba(0,0,0,.04);
}

.hero{
background:linear-gradient(135deg,#0f172a,#1d4ed8);
border-radius:18px;
padding:50px 30px;
padding:40px 30px;
color:white;
margin-bottom:18px;
box-shadow:0 8px 20px rgba(0,0,0,.12);
}

.hero h1{
font-size:34px;
margin:0 0 6px;
}

.hero p{
font-size:14px;
margin:0;
color:#dbeafe;
}

.upload-card{
background:white;
padding:25px;
border-radius:18px;
border:1px solid #e5e7eb;
box-shadow:0 5px 18px rgba(0,0,0,.05);
margin-bottom:25px;
}

.result-card{
background:white;
padding:20px;
border-radius:18px;
border:1px solid #e5e7eb;
box-shadow:0 5px 18px rgba(0,0,0,.05);
}

.metric-card{
background:white;
padding:20px;
border-radius:16px;
border:1px solid #ececec;
text-align:center;
box-shadow:0 5px 12px rgba(0,0,0,.04);
}

.metric-card h1{
font-size:42px;
margin:0;
color:#2563eb;
}

.metric-card p{
margin:0;
color:#64748b;
font-size:14px;
}

.stButton > button{
    width:100%;
    height:78px;
    border-radius:14px;
    font-size:20px;
    font-weight:700;
    border:none;
    transition:0.25s;
    color:white;
    background:linear-gradient(90deg,#2563eb,#3b82f6);
    box-shadow:0 8px 20px rgba(37,99,235,.25);
}

.stButton > button{

width:100%;
height:58px;

border:none;
border-radius:14px;

background:linear-gradient(135deg,#2563eb,#1d4ed8);

color:white;

font-size:17px;
font-weight:700;

transition:.25s;

box-shadow:0 8px 20px rgba(37,99,235,.25);

}

.stButton > button:hover{

transform:translateY(-3px);

box-shadow:0 12px 28px rgba(37,99,235,.35);

}

hr{
margin:25px 0;
}

</style>
""",unsafe_allow_html=True)

st.markdown("""
<div class="hero">

<h1>ODP KML / KMZ Generator</h1>

<p>
Internal tools untuk mengubah data Excel menjadi file KML / KMZ
siap digunakan di Google Earth.
</p>

</div>
""", unsafe_allow_html=True)


st.markdown("""
<div style="
background:white;
padding:5px 14px;
border-radius:10px;
border:1px solid #e5e7eb;
margin-bottom:8px;
box-shadow:0 3px 10px rgba(0,0,0,.04);
">
<h3 style="margin:0;">
1. Upload File Excel ODP
</h3>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload file Excel ODP terbaru",
    type=["xlsx","xls"]
)


IDLE_ICON = "https://maps.google.com/mapfiles/kml/paddle/blue-blank.png"
FULL_ICON = "https://maps.google.com/mapfiles/kml/paddle/red-blank.png"

# =========================
# ICON CST
# =========================

CST_ICONS = {
    "👤 Customer": "https://maps.google.com/mapfiles/kml/paddle/grn-blank.png",
    "📍 Customer 2": "https://maps.google.com/mapfiles/kml/paddle/ylw-blank.png",
    "🏠 Customer 3": "https://maps.google.com/mapfiles/kml/paddle/purple-blank.png",
    "🔵 Customer 4": "https://maps.google.com/mapfiles/kml/paddle/blu-blank.png",
    "🟠 Customer 5": "https://maps.google.com/mapfiles/kml/paddle/orange-blank.png",
}

# =========================
# SESSION STATE ICON CST
# =========================

if "cst_icon_name" not in st.session_state:
    st.session_state.cst_icon_name = "👤 Customer"

if "cst_icon_url" not in st.session_state:
    st.session_state.cst_icon_url = CST_ICONS["👤 Customer"]


# =========================
# POPUP PILIH ICON CST
# =========================

@st.dialog("🎨 Pilih Icon Customer")
def pilih_icon_cst():

    st.write("Pilih icon yang akan digunakan untuk titik Customer / CST.")

    pilihan = st.selectbox(
        "Icon Customer",
        list(CST_ICONS.keys()),
        index=list(CST_ICONS.keys()).index(
            st.session_state.cst_icon_name
        )
    )

    st.markdown(
        f"""
        <div style="
            text-align:center;
            padding:20px;
            border:1px solid #e5e7eb;
            border-radius:12px;
            margin-top:10px;
        ">
            <p style="margin-bottom:8px;color:#64748b;">
                Preview Icon
            </p>
            <img src="{CST_ICONS[pilihan]}"
                 width="50">
            <br>
            <b>{pilihan}</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    if st.button(
        "✅ Gunakan Icon Ini",
        use_container_width=True
    ):
        st.session_state.cst_icon_name = pilihan
        st.session_state.cst_icon_url = CST_ICONS[pilihan]

        st.rerun()

def read_excel_auto_header(file):
    raw = pd.read_excel(file, header=None)

    for i in range(10):
        row_values = raw.iloc[i].astype(str).str.strip().tolist()
        if "Code" in row_values and "Kelurahan" in row_values and "Kecamatan" in row_values:
            df = pd.read_excel(file, header=i)
            df.columns = df.columns.astype(str).str.strip()
            return df

    return pd.read_excel(file)

def find_coordinate_column(df):
    for col in df.columns:
        sample = df[col].dropna().astype(str).head(20)
        for val in sample:
            if "," in val:
                parts = val.split(",")
                if len(parts) == 2:
                    try:
                        float(parts[0].strip())
                        float(parts[1].strip())
                        return col
                    except:
                        pass
    return None

kml_path = "ODP_Master.kml"
kmz_path = "ODP_Master.kmz"

st.sidebar.markdown("""
<div style="text-align:center;">

<h2 style="margin-bottom:2px;">
📍 ODP TOOLS
</h2>

<div style="
font-size:14px;
color:#64748B;
margin-bottom:12px;
">
District Management
</div>

</div>

---
""", unsafe_allow_html=True)

if not uploaded_file:

    st.sidebar.info("Upload file Excel terlebih dahulu.")

    st.sidebar.selectbox(
        "Folder Level 1",
        ["Upload Excel dulu"],
        disabled=True
    )

    st.sidebar.selectbox(
        "Folder Level 2",
        ["Upload Excel dulu"],
        disabled=True
    )

    st.sidebar.selectbox(
        "Folder Level 3",
        ["Upload Excel dulu"],
        disabled=True
    )


if uploaded_file:

    df = read_excel_auto_header(uploaded_file)
    coord_col = find_coordinate_column(df)

    folder_columns = [c for c in df.columns if c != coord_col]

if uploaded_file:

    df = read_excel_auto_header(uploaded_file)
    coord_col = find_coordinate_column(df)

    # =========================
    # JENIS TITIK
    # =========================

    st.markdown("### 📍 Jenis Titik")

    jenis_titik = st.radio(
        "Pilih jenis data yang akan dibuat menjadi KMZ",
        ["ODP", "Customer / CST"],
        horizontal=True
    )

    if jenis_titik == "ODP":

        st.info(
            "🔵 IDLE = Biru  |  🔴 FULL = Merah"
        )

    else:

        col_icon1, col_icon2 = st.columns([3, 1])

        with col_icon1:
            st.info(
                f"👤 Icon CST saat ini: "
                f"**{st.session_state.cst_icon_name}**"
            )

        with col_icon2:
            if st.button(
                "🎨 Pilih Icon",
                use_container_width=True
            ):
                pilih_icon_cst()

    folder_columns = [c for c in df.columns if c != coord_col]
    

    folder1 = st.sidebar.selectbox(
        "Folder Level 1",
        folder_columns,
        index=folder_columns.index("Region") if "Region" in folder_columns else 0
    )

    folder2 = st.sidebar.selectbox(
        "Folder Level 2",
        ["Tidak dipisah"] + folder_columns,
        index=folder_columns.index("District Name") + 1 if "District Name" in folder_columns else 0
    )

    folder3 = st.sidebar.selectbox(
        "Folder Level 3",
        ["Tidak dipisah"] + folder_columns,
        index=0
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 👁 Preview Struktur")

    preview = folder1

    if folder2 != "Tidak dipisah":
        preview += f"\n└── {folder2}"

    if folder3 != "Tidak dipisah":
        preview += f"\n    └── {folder3}"

    preview += f"\n        └── {jenis_titik}"


    st.sidebar.markdown(f"""
    <div style="
    background:white;
    padding:17.4px;
    border-radius:12px;
    border:1px solid #E5E7EB;
    font-family:Consolas;
    line-height:1.7;
    ">
    
    {preview.replace(chr(10),"<br>")}
    
    </div>
    """, unsafe_allow_html=True)



    # =========================
    # VALIDASI FILE
    # =========================
    
    if coord_col is None:
    
        st.error(
            "❌ Kolom koordinat tidak ditemukan. "
            "Pastikan ada satu kolom dengan format Lat,Long."
        )
    
    else:
    
        st.success(
            f"📍 Koordinat terdeteksi di kolom: {coord_col}"
        )
    
        st.info(
            f"📊 Total data yang terbaca: {len(df)} baris"
        )
    
        # lanjut ke tombol Generate / Publish

        col1, col2 = st.columns(2)

        with col1:
            generate = st.button(
                "🚀 Generate\n\nBuat file KML & KMZ",
                use_container_width=True
            )

        with col2:
            publish = st.button(
                "☁️ Publish\n\nPublish ke GitHub",
                use_container_width=True
            )
        st.divider()

        if generate or publish:

            # =========================
            # BUAT KML
            # =========================

            kml = simplekml.Kml(name=f"Update {today}")

            stats = {
                "total": 0,
                "skipped": 0
            }


            # =========================
            # FUNGSI BUAT POINT
            # =========================

            def create_point(target_folder, row):

                # =========================
                # BACA KOORDINAT
                # =========================

                try:

                    coord = str(row[coord_col]).strip()

                    lat, lon = coord.split(",")

                    lat = float(lat.strip())
                    lon = float(lon.strip())

                except Exception:

                    stats["skipped"] += 1
                    return


                # =========================
                # STATUS & ICON
                # =========================

                if jenis_titik == "ODP":

                    if "Capacity" in df.columns and "Active" in df.columns:

                        capacity = pd.to_numeric(
                            row["Capacity"],
                            errors="coerce"
                        )

                        active = pd.to_numeric(
                            row["Active"],
                            errors="coerce"
                        )

                        capacity = (
                            0 if pd.isna(capacity)
                            else int(capacity)
                        )

                        active = (
                            0 if pd.isna(active)
                            else int(active)
                        )

                        status = (
                            "FULL"
                            if capacity > 0 and active >= capacity
                            else "IDLE"
                        )

                    else:

                        status = "IDLE"


                    # ICON ODP

                    if status == "FULL":

                        icon_url = FULL_ICON
                        header_color = "#E53935"

                    else:

                        icon_url = IDLE_ICON
                        header_color = "#4285F4"


                else:

                    # =========================
                    # CUSTOMER / CST
                    # =========================

                    status = "CUSTOMER"

                    icon_url = st.session_state.cst_icon_url

                    header_color = "#16A34A"


                # =========================
                # NAMA TITIK
                # =========================

                if "Code" in df.columns:

                    point_name = str(row["Code"])

                elif "Customer ID" in df.columns:

                    point_name = str(row["Customer ID"])

                elif "ID" in df.columns:

                    point_name = str(row["ID"])

                elif "Nama" in df.columns:

                    point_name = str(row["Nama"])

                elif "Name" in df.columns:

                    point_name = str(row["Name"])

                else:

                    point_name = f"Point {stats['total'] + 1}"


                # =========================
                # DATA POPUP
                # =========================

                table_rows = ""

                for col in df.columns:

                    if col == coord_col:
                        continue

                    value = row[col]

                    if pd.isna(value):
                        value = "-"

                    table_rows += f"""
                    <tr>
                        <td><b>{col}</b></td>
                        <td>{value}</td>
                    </tr>
                    """


                desc = f"""
                <div style="font-family:Arial; font-size:12px;">

                    <table
                        border="1"
                        cellpadding="5"
                        cellspacing="0"
                        width="300"
                    >

                        <tr>
                            <th colspan="2" bgcolor="{header_color}">
                                <font color="white">
                                    {point_name}
                                </font>
                            </th>
                        </tr>

                        {table_rows}

                        <tr>
                            <td><b>Status</b></td>
                            <td>{status}</td>
                        </tr>

                        <tr>
                            <td><b>Lat</b></td>
                            <td>{lat}</td>
                        </tr>

                        <tr>
                            <td><b>Long</b></td>
                            <td>{lon}</td>
                        </tr>

                    </table>

                </div>
                """


                # =========================
                # BUAT POINT
                # =========================

                pnt = target_folder.newpoint(
                    name=point_name,
                    coords=[(lon, lat)]
                )

                pnt.description = ""

                pnt.snippet = Snippet(
                    "",
                    maxlines=0
                )

                pnt.style.balloonstyle.text = desc


                # =========================
                # PASANG ICON
                # =========================

                pnt.style.iconstyle.icon.href = icon_url

                pnt.style.iconstyle.scale = 1.2


                # =========================
                # COUNTER
                # =========================

                stats["total"] += 1


            # =========================
            # BUAT STRUKTUR FOLDER
            # =========================

            for value1, df1 in df.groupby(
                folder1,
                dropna=False
            ):

                folder_a = kml.newfolder(
                    name=str(value1)
                )


                if folder2 == "Tidak dipisah":

                    for _, row in df1.iterrows():

                        create_point(
                            folder_a,
                            row
                        )


                else:

                    for value2, df2 in df1.groupby(
                        folder2,
                        dropna=False
                    ):

                        folder_b = folder_a.newfolder(
                            name=str(value2)
                        )


                        if folder3 == "Tidak dipisah":

                            for _, row in df2.iterrows():

                                create_point(
                                    folder_b,
                                    row
                                )


                        else:

                            for value3, df3 in df2.groupby(
                                folder3,
                                dropna=False
                            ):

                                folder_c = folder_b.newfolder(
                                    name=str(value3)
                                )

                                for _, row in df3.iterrows():

                                    create_point(
                                        folder_c,
                                        row
                                    )


            # =========================
            # SIMPAN KML
            # =========================

            kml.save(kml_path)


            # =========================
            # BUAT KMZ
            # =========================

            with zipfile.ZipFile(
                kmz_path,
                "w",
                zipfile.ZIP_DEFLATED
            ) as kmz:

                kmz.write(
                    kml_path,
                    "doc.kml"
                )


            # =========================
            # HASIL GENERATE
            # =========================

            left, right = st.columns([1.05, 1])


            with left:

                st.subheader("2. Hasil Generate")

                st.success(
                    "Generate selesai!"
                )


                c1, c2, c3 = st.columns(3)


                with c1:

                    st.markdown(
                        f"""
                        <div class="metric-card">

                            <p>Total {jenis_titik}</p>

                            <h1>
                                {stats["total"]}
                            </h1>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )


                with c2:

                    st.markdown(
                        f"""
                        <div class="metric-card">

                            <p>Skipped</p>

                            <h1>
                                {stats["skipped"]}
                            </h1>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )


                with c3:

                    st.markdown(
                        """
                        <div class="metric-card">

                            <p>Status</p>

                            <h1 style="font-size:34px;">
                                Success
                            </h1>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )


                st.success(
                    f"File berhasil dibuat! "
                    f"Total titik: {stats['total']}, "
                    f"dilewati: {stats['skipped']}"
                )


                # =========================
                # DOWNLOAD
                # =========================

                col1, col2 = st.columns(2)


                with col1:

                    st.markdown(
                        "### 📄 Download KML"
                    )

                    with open(
                        kml_path,
                        "rb"
                    ) as f:

                        st.download_button(
                            "File KML",
                            f,
                            file_name="ODP_Master.kml",
                            use_container_width=True
                        )


                with col2:

                    st.markdown(
                        "### 📦 Download KMZ"
                    )

                    with open(
                        kmz_path,
                        "rb"
                    ) as f:

                        st.download_button(
                            "File KMZ",
                            f,
                            file_name="ODP_Master.kmz",
                            use_container_width=True
                        )

            with right:

                    st.subheader("3. Informasi Publish")

                    if publish:
                        token = st.secrets["GITHUB_TOKEN"]
                        repo = st.secrets["GITHUB_REPO"]
                        branch = st.secrets["GITHUB_BRANCH"]

                        with open(kmz_path, "rb") as file:
                            content = base64.b64encode(file.read()).decode()

                        url = f"https://api.github.com/repos/{repo}/contents/ODP_Master.kmz"

                        headers = {
                            "Authorization": f"Bearer {token}",
                            "Accept": "application/vnd.github+json"
                        }

                        get = requests.get(url, headers=headers)

                        sha = None

                        if get.status_code == 200:
                            sha = get.json()["sha"]

                        payload = {
                            "message": f"Update KMZ {today}",
                            "content": content,
                            "branch": branch
                        }

                        if sha:
                            payload["sha"] = sha

                        response = requests.put(
                            url,
                            headers=headers,
                            json=payload
                        )

                        if response.status_code in [200, 201]:

                            st.success("✔ Publish berhasil!")

                            st.markdown(f"""
                            <table style="
                            width:100%;
                            border-collapse:collapse;
                            font-size:15px;
                            line-height:1.2;
                            ">
                            
                            <tr style="background:#F8FAFC;">
                            <th align="left" style="padding:17.4px;border:1px solid #E5E7EB;">Informasi</th>
                            <th align="left" style="padding:17.4px;border:1px solid #E5E7EB;">Nilai</th>
                            </tr>
                            
                            <tr>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">📅 Update</td>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">{today}</td>
                            </tr>
                            
                            <tr>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">📍 Total ODP</td>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">{stats["total"]}</td>
                            </tr>
                            
                            <tr>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">☁️ Status</td>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">GitHub berhasil diperbarui</td>
                            </tr>
                            
                            <tr>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">🌿 Branch</td>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">{branch}</td>
                            </tr>
                            
                            <tr>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">📂 Repository</td>
                            <td style="padding:17.4px;border:1px solid #E5E7EB;">{repo}</td>
                            </tr>
                            
                            </table>
                            """, unsafe_allow_html=True)
