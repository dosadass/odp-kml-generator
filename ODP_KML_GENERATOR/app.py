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
.block-container {
    padding-top: 2rem;
    max-width: 1100px;
}

.main-title {
    background: linear-gradient(135deg, #1e293b, #334155);
    padding: 28px 34px;
    border-radius: 18px;
    color: white;
    margin-bottom: 24px;
    box-shadow: 0 8px 24px rgba(15,23,42,0.18);
}

.main-title h1 {
    margin: 0;
    font-size: 36px;
    font-weight: 800;
}

.main-title p {
    margin-top: 8px;
    font-size: 15px;
    color: #cbd5e1;
}

.badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(255,255,255,0.08);
    backdrop-filter: blur(8px);
    color: white;
    padding: 8px 14px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 18px;
    border: 1px solid rgba(255,255,255,0.12);
}

.info-card {
    background: #ffffff;
    padding: 18px 22px;
    border-radius: 14px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 4px 14px rgba(15,23,42,0.08);
    margin-bottom: 18px;
}
</style>

<div class="main-title">
    <div class="badge">📍 DISTRICT MANAGEMENT</div>
    <h1>ODP KML / KMZ Generator</h1>
    <p>Internal tools untuk mengubah data ODP Excel menjadi file KML/KMZ siap pakai di Google Earth.</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="info-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload file Excel ODP terbaru", type=["xlsx", "xls"])
st.caption("Pastikan file memiliki kolom Code, Kelurahan, Kecamatan, Region, District Name, Capacity, Active, dan Coordinate.")
st.markdown('</div>', unsafe_allow_html=True)

required_cols = [
    "Code", "Kelurahan", "Kecamatan", "Region", "District Name",
    "Ms. Partner Name", "Capacity", "Active"
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

if uploaded_file:
    df = read_excel_auto_header(uploaded_file)

    st.write("Preview Data:")
    st.dataframe(df.head())

    st.write("Nama Kolom Terdeteksi:")
    st.write(list(df.columns))

    coord_col = find_coordinate_column(df)

    missing = [col for col in required_cols if col not in df.columns]

    if coord_col is None:
        missing.append("Kolom koordinat format Lat,Long")

    if missing:
        st.error(f"Kolom ini belum ada / beda nama: {missing}")
    else:
        st.success(f"Koordinat terdeteksi di kolom: {coord_col}")

        if st.button("🚀 Generate + Publish"):
            kml = simplekml.Kml(name=f"Update {today}")
            total_point = 0
            skipped_point = 0

            for region, region_df in df.groupby("Region"):
                region_folder = kml.newfolder(name=str(region))

                for district, district_df in region_df.groupby("District Name"):
                    district_folder = region_folder.newfolder(name=str(district))

                    for _, row in district_df.iterrows():
                        try:
                            coord = str(row[coord_col]).strip()
                            lat, lon = coord.split(",")
                            lat = float(lat.strip())
                            lon = float(lon.strip())
                        except:
                            skipped_point += 1
                            continue

                        capacity = int(row["Capacity"]) if pd.notna(row["Capacity"]) else 0
                        active = int(row["Active"]) if pd.notna(row["Active"]) else 0

                        status = "FULL" if capacity > 0 and active >= capacity else "IDLE"
                        header_color = "#E53935" if status == "FULL" else "#4285F4"

                        desc = f"""
<div style="font-family:Arial; font-size:12px;">
<table border="1" cellpadding="5" cellspacing="0" width="300">
<tr>
    <th colspan="2" bgcolor="{header_color}">
        <font color="white">{row['Code']}</font>
    </th>
</tr>
<tr><td><b>Code</b></td><td>{row['Code']}</td></tr>
<tr><td><b>Kelurahan</b></td><td>{row['Kelurahan']}</td></tr>
<tr><td><b>Kecamatan</b></td><td>{row['Kecamatan']}</td></tr>
<tr><td><b>Region</b></td><td>{row['Region']}</td></tr>
<tr><td><b>District</b></td><td>{row['District Name']}</td></tr>
<tr><td><b>Partner</b></td><td>{row['Ms. Partner Name']}</td></tr>
<tr><td><b>Capacity</b></td><td>{capacity}</td></tr>
<tr><td><b>Active</b></td><td>{active}</td></tr>
<tr><td><b>Status</b></td><td>{status}</td></tr>
<tr><td><b>Lat</b></td><td>{lat}</td></tr>
<tr><td><b>Long</b></td><td>{lon}</td></tr>
</table>
</div>
"""

                        pnt = district_folder.newpoint(
                            name=str(row["Code"]),
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
                        total_point += 1


            kml.save(kml_path)

            with zipfile.ZipFile(kmz_path, "w", zipfile.ZIP_DEFLATED) as kmz:
                kmz.write(kml_path, "doc.kml")

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
            
            if response.status_code in [200,201]:
                st.success("✅ Publish berhasil!")

                st.markdown(f"""
                ### Informasi Publish
                
                - 📅 **Update** : {today}
                - 📍 **Total ODP** : **{total_point}**
                - ☁️ **Status** : GitHub berhasil diperbarui.
                """)
            else:
                st.write(response.status_code)
                st.write(response.json())

            st.success(f"File berhasil dibuat! Total titik: {total_point}, dilewati: {skipped_point}")

            with open(kml_path, "rb") as f:
                st.download_button("Download KML", f, file_name="ODP_Master.kml")

            with open(kmz_path, "rb") as f:
                st.download_button("Download KMZ", f, file_name="ODP_Master.kmz")

        





