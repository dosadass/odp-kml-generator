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

padding-top:10px;

padding-bottom:20px;

max-width:1600px;

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

border-radius:18px;

padding:14px;

border:1px solid #DBEAFE;

box-shadow:0 8px 24px rgba(0,0,0,.05);

margin-bottom:10px;

}

.hero{
background:linear-gradient(135deg,#0f172a,#1d4ed8);
border-radius:22px;
padding:28px 34px;
color:white;
margin-bottom:15px;
box-shadow:0 12px 30px rgba(0,0,0,.18);
}

.hero h1{
font-size:36px;
margin-bottom:10px;
}

.hero p{
font-size:16px;
color:#dbeafe;
}

.badge{
display:inline-block;
padding:6px 12px;
background:#2563eb;
border-radius:999px;
font-size:13px;
font-weight:700;
margin-bottom:18px;
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

background:linear-gradient(180deg,#FFFFFF,#F8FAFC);

border-radius:14px;

padding:12px;

border:1px solid #DBEAFE;

box-shadow:0 8px 20px rgba(59,130,246,.08);

transition:.25s;

text-align:center;

height:90px;

}

.metric-card:hover{

transform:translateY(-5px);

box-shadow:0 15px 35px rgba(37,99,235,.15);

}

.metric-card h1{

margin-top:10px;

margin-bottom:5px;

font-size:28px;

color:#2563EB;

}

.metric-card p{

margin:0;

font-weight:700;

color:#64748B;

font-size:13px;

}

.metric-card h1{
font-size:28px;
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
height:48px;

border:none;
border-radius:10px;

background:linear-gradient(135deg,#2563eb,#1d4ed8);

color:white;

font-size:15px;
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

<div class="badge">
📍 DISTRICT MANAGEMENT
</div>

<h1>ODP KML / KMZ Generator</h1>

<p>
Internal tools untuk mengubah data Excel menjadi file KML / KMZ
siap digunakan di Google Earth.
</p>

</div>
""", unsafe_allow_html=True)


st.markdown("""
<div class="card">

<h3 style="margin-top:0;">
📤 1. Upload File Excel ODP
</h3>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload file Excel ODP terbaru",
    type=["xlsx","xls"]
)

st.caption(
    "Pastikan file memiliki kolom Code, Kelurahan, Kecamatan, Region, District Name, Capacity, Active, dan Coordinate."
)

st.markdown("</div>", unsafe_allow_html=True)


required_cols = [
    "Code",
    "Promo",
    "Kelurahan",
    "Kecamatan",
    "Region",
    "District Name",
    "Ms. Partner Name",
    "Capacity",
    "Active"
]

IDLE_ICON = "https://maps.google.com/mapfiles/kml/paddle/blu-blank.png"
FULL_ICON = "https://maps.google.com/mapfiles/kml/paddle/red-blank.png"

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

<h2 style="margin-bottom:0;color:#1D4ED8;">
📍 ODP TOOLS
</h2>

<div style="
color:#64748B;
font-size:14px;
margin-top:-4px;
">
District Management
</div>

</div>

<br>
<hr>
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

    st.sidebar.markdown("### 📂 STRUKTUR FOLDER")
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

    st.sidebar.markdown("""
    <div style="height:8px"></div>
    <hr style="margin:10px 0;">
    """, unsafe_allow_html=True)
    
    st.sidebar.markdown("### 👁 Preview Struktur")
    
    preview = folder1
    
    if folder2 != "Tidak dipisah":
        preview += f"\n└── {folder2}"
    
    if folder3 != "Tidak dipisah":
        preview += f"\n    └── {folder3}"
    
    preview += "\n        └── ODP"
    
    
    st.sidebar.markdown(f"""
    <div style="
    margin-top:12px;
    background:white;
    padding:12px;
    border-radius:12px;
    border:1px solid #E5E7EB;
    font-family:Consolas;
    line-height:1.7;
    ">
    
    {preview.replace(chr(10),"<br>")}
    
    </div>
    """, unsafe_allow_html=True)

    if uploaded_file:
        st.sidebar.markdown(f"""
        <div style="
        background:white;
        line-height:1.5;
        padding:12px;
        border-radius:16px;
        border:1px solid #E5E7EB;
        box-shadow:0 5px 15px rgba(0,0,0,.05);
        ">
        
        <h4 style="
        margin:0;
        color:#2563EB;
        ">
        📊 RINGKASAN DATA
        </h4>
        
        <hr>
        
        <b>📅 Update</b><br>
        {today}
        
        <br><br>
        
        <b>📍 Total ODP</b><br>
        {len(df)}
        
        <br><br>
        
        <b>🌍 Region</b><br>
        {df['Region'].nunique()}
        
        <br><br>
        
        <b>🏢 District</b><br>
        {df['District Name'].nunique()}
        
        </div>
        """, unsafe_allow_html=True)

        



    missing = [col for col in required_cols if col not in df.columns]

    if coord_col is None:
        missing.append("Kolom koordinat format Lat,Long")

    if missing:
        st.error(f"Kolom ini belum ada / beda nama: {missing}")
    else:

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
        st.markdown(
        """
        <div style="margin-top:15px;margin-bottom:10px;border-top:1px solid #E5E7EB;"></div>
        """,
        unsafe_allow_html=True
        )
        
        if generate or publish:
        
            kml = simplekml.Kml(name=f"Update {today}")
            stats = {
                "total": 0,
                "skipped": 0
            }

            def create_point(target_folder, row):
                try:
                   coord = str(row[coord_col]).strip()
                   lat, lon = coord.split(",")
                   lat = float(lat.strip())
                   lon = float(lon.strip())
                except:
                   stats["skipped"] += 1
                   return

                capacity = int(row["Capacity"]) if pd.notna(row["Capacity"]) else 0
                active = int(row["Active"]) if pd.notna(row["Active"]) else 0

                status = "FULL" if capacity > 0 and active >= capacity else "IDLE"
                header_color = "#E53935" if status == "FULL" else "#4285F4"
                        
                promo = ""
                        
                if pd.notna(row["Promo"]):
                    promo = str(row["Promo"]).strip()
                        
                if promo:
                    point_name = f"{row['Code']} - {promo}"
                else:
                    point_name = str(row["Code"])
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
                <table border="1" cellpadding="5" cellspacing="0" width="300">
                
                <tr>
                    <th colspan="2" bgcolor="{header_color}">
                        <font color="white">{point_name}</font>
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
    
                            
                            
                pnt = target_folder.newpoint(
                    name=point_name,
                    coords=[(lon, lat)]
                )
        
                pnt.description = ""
                pnt.snippet = Snippet("", maxlines=0)
                pnt.style.balloonstyle.text = desc
        
                if status == "FULL":
                    pnt.style.iconstyle.icon.href = FULL_ICON
                else:
                    pnt.style.iconstyle.icon.href = IDLE_ICON
    
                pnt.style.iconstyle.scale = 1.2
                stats["total"] += 1
        
            for value1, df1 in df.groupby(folder1):

                folder_a = kml.newfolder(name=str(value1))
            
                if folder2 == "Tidak dipisah":
            
                    for _, row in df1.iterrows():
            
                        target_folder = folder_a
            
                        create_point(target_folder, row)
            
                else:
            
                    for value2, df2 in df1.groupby(folder2):
            
                        folder_b = folder_a.newfolder(name=str(value2))
            
                        if folder3 == "Tidak dipisah":
            
                            for _, row in df2.iterrows():
            
                                target_folder = folder_b
            
                                create_point(target_folder, row)
            
                        else:
            
                            for value3, df3 in df2.groupby(folder3):

                                folder_c = folder_b.newfolder(name=str(value3))
                            
                                for _, row in df3.iterrows():
                            
                                    create_point(folder_c, row)


            kml.save(kml_path)
            
            with zipfile.ZipFile(kmz_path, "w", zipfile.ZIP_DEFLATED) as kmz:
                kmz.write(kml_path, "doc.kml")

            
            left,right = st.columns([1,1], gap="small")
        
            with left:
        
                st.markdown("""
                <h2 style="
                margin-bottom:20px;
                font-weight:700;
                ">
                📊 Hasil Generate
                </h2>
                """, unsafe_allow_html=True)
        
                st.markdown("""
                <div style="
                background:#ECFDF5;
                border:1px solid #BBF7D0;
                padding:12px;
                border-radius:14px;
                font-size:18px;
                font-weight:700;
                color:#166534;
                margin-bottom:18px;
                ">
                ✅ Generate berhasil
                </div>
                """, unsafe_allow_html=True)
        
                c1,c2,c3 = st.columns(3)
                
                with c1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <p>Total ODP</p>
                        <h1>{stats["total"]}</h1>
                    </div>
                    """, unsafe_allow_html=True)
                
                with c2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <p>Skipped</p>
                        <h1>{stats["skipped"]}</h1>
                    </div>
                    """, unsafe_allow_html=True)
                
                with c3:
                    st.markdown("""
                    <div class="metric-card">
                        <p>Status</p>
                        <h1 style="font-size:34px;">Success</h1>
                    </div>
                    """, unsafe_allow_html=True)
                    
                st.markdown(f"""
                <div style="
                background:#ECFDF5;
                border:1px solid #BBF7D0;
                padding:14px;
                border-radius:12px;
                color:#166534;
                font-weight:600;
                ">
                
                ✅ File berhasil dibuat
                
                <br><br>
                
                Total titik :
                <b>{stats["total"]}</b>
                
                <br>
                
                Skipped :
                <b>{stats["skipped"]}</b>
                
                </div>
                """, unsafe_allow_html=True)
            
                col1,col2 = st.columns(2, gap="small")
    
                with col1:

                    st.markdown("""
                    <div class="card">
                    📥 Download KML
                    </div>
                    """, unsafe_allow_html=True)
                
                    with open(kml_path,"rb") as f:
                        st.download_button(
                            "ODP_Master.kml",
                            f,
                            file_name="ODP_Master.kml",
                            use_container_width=True
                        )
                
                with col2:

                        st.markdown("📥 Download KMZ")
                    
                        with open(kmz_path,"rb") as f:
                            st.download_button(
                                "ODP_Master.kmz",
                                f,
                                file_name="ODP_Master.kmz",
                                use_container_width=True
                            )

            with right:
                    
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
                    
                            st.markdown(f"""
                            <table style="
                            width:100%;
                            border-collapse:collapse;
                            font-size:14px;
                            ">
                            
                            <tr>
                            <th align="left">Informasi</th>
                            <th align="left">Nilai</th>
                            </tr>
                            
                            <tr>
                            <td>📅 Update</td>
                            <td>{today}</td>
                            </tr>
                            
                            <tr>
                            <td>📍 Total ODP</td>
                            <td>{stats["total"]}</td>
                            </tr>
                            
                            <tr>
                            <td>☁️ Status</td>
                            <td>GitHub berhasil diperbarui</td>
                            </tr>
                            
                            <tr>
                            <td>🌿 Branch</td>
                            <td>{branch}</td>
                            </tr>
                            
                            <tr>
                            <td>📂 Repository</td>
                            <td>{repo}</td>
                            </tr>
                            
                            </table>
                            """, unsafe_allow_html=True)

            
            <tr>
            <td>📅 Update</td>
            <td>{today}</td>
            </tr>
            
            <tr>
            <td>📍 Total ODP</td>
            <td>{stats["total"]}</td>
            </tr>
            
            <tr>
            <td>☁️ Status</td>
            <td>GitHub berhasil diperbarui</td>
            </tr>
            
            <tr>
            <td>🌿 Branch</td>
            <td>{branch}</td>
            </tr>
            
            <tr>
            <td>📂 Repository</td>
            <td>{repo}</td>
            </tr>
            
            </table>
            """, unsafe_allow_html=True)


        




