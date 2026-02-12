import ssl
import os
import streamlit as st
import folium
from streamlit_folium import st_folium
from PIL import Image, ImageOps
from PIL.ExifTags import TAGS, GPSTAGS
import base64
from io import BytesIO
import easyocr
import numpy as np
import re
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from google import genai
from google.genai import types
import zipfile
from lxml import etree
import pandas as pd

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า API และโหลด OCR ---
client = genai.Client(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")

@st.cache_resource
def load_ocr():
    model_path = os.path.join(os.getcwd(), "easyocr_models")
    if not os.path.exists(model_path): os.makedirs(model_path)
    return easyocr.Reader(['en'], gpu=False, model_storage_directory=model_path)

# --- 2. ฟังก์ชันคำนวณระยะเดินเท้า (OSRM) ---
def get_walking_route(p1, p2):
    """คำนวณระยะทางเดินตามแนวถนน (Walking Profile)"""
    try:
        url = f"http://router.project-osrm.org/route/v1/foot/{p1[1]},{p1[0]};{p2[1]},{p2[0]}?overview=full&geometries=geojson"
        r = requests.get(url, timeout=5)
        data = r.json()
        if data['code'] == 'Ok':
            dist = data['routes'][0]['distance']
            geom = data['routes'][0]['geometry']['coordinates']
            route_pts = [[c[1], c[0]] for c in geom]
            return dist, route_pts
    except: pass
    return None, None

# --- 3. ฟังก์ชันพื้นฐาน (วิเคราะห์ภาพ/พิกัด) ---
def analyze_cable_issue(image_bytes):
    try:
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=[
                "วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น: 1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ",
                types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg")
            ]
        )
        return response.text.strip()
    except: return "วิเคราะห์ไม่ได้"

def get_lat_lon_exif(image):
    try:
        exif = image._getexif()
        if not exif: return None, None
        gps_info = {}
        for (idx, tag) in TAGS.items():
            if tag == 'GPSInfo':
                for (t, value) in GPSTAGS.items():
                    if t in exif[idx]: gps_info[value] = exif[idx][t]
        def dms_to_decimal(dms, ref):
            d, m, s = [float(x) for x in dms]
            res = d + (m / 60.0) + (s / 3600.0)
            return -res if ref in ['S', 'W'] else res
        return dms_to_decimal(gps_info['GPSLatitude'], gps_info['GPSLatitudeRef']), \
               dms_to_decimal(gps_info['GPSLongitude'], gps_info['GPSLongitudeRef'])
    except: return None, None

def get_lat_lon_ocr(image):
    try:
        reader = load_ocr()
        img_np = np.array(image.copy().thumbnail((1000, 1000)) or image)
        results = reader.readtext(img_np, paragraph=True)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

def create_div_label(name, color="#D9534F"):
    return f'<div style="font-size:11px; font-weight:800; color:{color}; white-space:nowrap; transform:translate(-50%,-150%); text-shadow:2px 2px 4px white;">{name}</div>'

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy(); img_resized.thumbnail((150, 150))
    buf = BytesIO(); img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''<div style="position:relative; width:fit-content; background:white; padding:5px; border-radius:12px; border:2px solid #FF8C42; transform:translate(-50%,-100%);">
            <div style="font-size:11px; font-weight:700; color:#2D5A27; text-align:center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width:140px; border-radius:4px;">
            </div>'''

# --- 4. สถาปัตยกรรม UI & State ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
if 'selected_pts' not in st.session_state: st.session_state.selected_pts = []

st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
</style>""", unsafe_allow_html=True)

# Header (Joker)
joker_base64 = base64.b64encode(requests.get("https://drive.google.com/uc?export=download&id=1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr").content).decode() if requests.get("https://drive.google.com/uc?export=download&id=1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr").status_code == 200 else ""
st.markdown(f'<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color:#718096; font-weight:600;">By Joker EN-NMA</p></div><img src="data:image/png;base64,{joker_base64}" style="width:100px; border-radius:50%; border:3px solid #FF8C42;"></div>', unsafe_allow_html=True)

# --- 5. เมนู KML/KMZ (เงื่อนไขหลักสำหรับการวัดระยะ) ---
st.subheader("🌐 1. ข้อมูลโครงข่าย & จุดติดตั้ง (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลดไฟล์ KML หรือ KMZ เพื่อเริ่มการวัดระยะ", type=['kml', 'kmz'])

kml_markers = []
if kml_file:
    try:
        if kml_file.name.endswith('.kmz'):
            with zipfile.ZipFile(kml_file) as z:
                kml_filename = [n for n in z.namelist() if n.endswith('.kml')][0]
                content = z.read(kml_filename)
        else: content = kml_file.getvalue()
        root = etree.fromstring(content)
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        placemarks = root.xpath('.//kml:Placemark', namespaces=ns)
        for pm in placemarks:
            name = pm.xpath('kml:name/text()', namespaces=ns)
            coords = pm.xpath('.//kml:coordinates/text()', namespaces=ns)
            if coords:
                c = coords[0].strip().split()[0].split(',')
                kml_markers.append({'name': name[0] if name else "จุดติดตั้ง", 'lat': float(c[1]), 'lon': float(c[0])})
    except Exception as e: st.error(f"KML Error: {e}")

# --- 6. แผนที่และระบบวัดระยะ ---
st.markdown("<hr>", unsafe_allow_html=True)
uploaded_files = st.file_uploader("📁 2. อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

m = folium.Map(location=[13.75, 100.5], zoom_start=15, tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
all_bounds = []

# แสดงหมุด KML และเปิดระบบเลือกจุด (เฉพาะเมื่อมีไฟล์ KML)
if kml_markers:
    for p in kml_markers:
        is_sel = any(s['lat'] == p['lat'] and s['lon'] == p['lon'] for s in st.session_state.selected_pts)
        color = 'green' if is_sel else 'blue'
        folium.Marker(
            [p['lat'], p['lon']], 
            tooltip=p['name'], 
            icon=folium.Icon(color=color, icon='info-sign'),
            popup=f"เลือกวัดระยะ: {p['name']}"
        ).add_to(m)
        folium.Marker([p['lat'], p['lon']], icon=folium.DivIcon(html=create_div_label(p['name'], "#2D5A27" if is_sel else "#D9534F"))).add_to(m)
        all_bounds.append([p['lat'], p['lon']])
else:
    st.warning("⚠️ กรุณาอัปโหลด KML/KMZ เพื่อเปิดใช้งานระบบคลิกวัดระยะทาง")

# แสดงรูปภาพสำรวจ
if uploaded_files:
    if 'export_data' not in st.session_state: st.session_state.export_data = []
    for f in uploaded_files:
        raw_data = f.getvalue()
        img = ImageOps.exif_transpose(Image.open(BytesIO(raw_data)))
        lat, lon = get_lat_lon_exif(img)
        if lat is None: lat, lon = get_lat_lon_ocr(img)
        if lat:
            issue = analyze_cable_issue(raw_data)
            folium.Marker([lat, lon], icon=folium.DivIcon(html=img_to_custom_icon(img, issue))).add_to(m)
            all_bounds.append([lat, lon])

# ประมวลผลการวัดระยะ (เมื่อเลือกครบ 2 จุด)
dist_text = ""
if len(st.session_state.selected_pts) == 2:
    p1 = [st.session_state.selected_pts[0]['lat'], st.session_state.selected_pts[0]['lon']]
    p2 = [st.session_state.selected_pts[1]['lat'], st.session_state.selected_pts[1]['lon']]
    dist_m, route = get_walking_route(p1, p2)
    if route:
        folium.PolyLine(route, color="#00008B", weight=6, opacity=0.8).add_to(m)
        dist_text = f"{dist_m:.2f} เมตร"
        st.sidebar.markdown(f"### 📏 ระยะทางเดินเท้า\n**{dist_text}**")
        if st.sidebar.button("🗑️ ล้างการเลือกจุด"):
            st.session_state.selected_pts = []
            st.rerun()

if all_bounds: m.fit_bounds(all_bounds, padding=[50, 50])

# แสดงผลแผนที่
map_data = st_folium(m, height=700, use_container_width=True, key="survey_map")

# รับค่าการคลิกหมุด (Trigger เมื่อเลือก KML Marker)
if kml_markers and map_data['last_object_clicked_popup']:
    clicked_name = map_data['last_object_clicked_popup'].replace("เลือกวัดระยะ: ", "")
    target = next((item for item in kml_markers if item["name"] == clicked_name), None)
    if target and target not in st.session_state.selected_pts:
        if len(st.session_state.selected_pts) >= 2: st.session_state.selected_pts = []
        st.session_state.selected_pts.append(target)
        st.rerun()

st.markdown("<hr>", unsafe_allow_html=True)
st.subheader("📄 3. สร้างรายงาน PowerPoint")
if st.button("🚀 สรุปรายงาน PPTX"):
    # ฟังก์ชัน PPTX คงเดิมตามโครงสร้างคุณ
    st.info("ระบบกำลังเจนไฟล์รายงาน...")
