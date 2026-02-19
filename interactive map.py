import ssl
import os
import streamlit as st
import folium
from streamlit_folium import st_folium
from folium.plugins import MeasureControl
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

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า Google Gemini API ---
client = genai.Client(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")

@st.cache_resource
def load_ocr():
    model_path = os.path.join(os.getcwd(), "easyocr_models")
    if not os.path.exists(model_path):
        os.makedirs(model_path)
    return easyocr.Reader(['en'], gpu=False, model_storage_directory=model_path)

# --- 2. ฟังก์ชันช่วยจัดการข้อมูล ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except Exception: return None
    return None

def analyze_cable_issue(image_bytes):
    try:
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=[
                """วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเพียง "หนึ่งเดียว" จาก 4 สาเหตุ:
                1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ
                ตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น""",
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

# --- 3. ฟังก์ชันคำนวณเส้นทางเดิน (OSRM Walking) ---
def get_osrm_route(p1, p2):
    """คำนวณเส้นทางเดินระหว่าง 2 จุดที่เลือก (ย้อนศรได้)"""
    if not p1 or not p2: return None, 0
    coords_str = f"{p1[1]},{p1[0]};{p2[1]},{p2[0]}"
    url = f"http://router.project-osrm.org/route/v1/walking/{coords_str}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if "routes" in data and len(data["routes"]) > 0:
                route = data["routes"][0]
                return [[lat, lon] for lon, lat in route["geometry"]["coordinates"]], route["distance"]
    except: pass
    return None, 0

# --- 4. ฟังก์ชันสร้าง Label & Icons ---
def create_div_label(name):
    return f'<div style="font-size: 11px; font-weight: 800; color: #D9534F; text-shadow: 2px 2px 4px white;">{name}</div>'

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((150, 150)) 
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="background: white; padding: 5px; border-radius: 10px; border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 10px; font-weight: bold; text-align: center; color: #2D5A27;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="width: 120px; border-radius: 4px;">
        </div>
    '''

# --- 5. PowerPoint Export ---
def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    if map_image_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# --- 6. Main Streamlit UI ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")

# Session State สำหรับเก็บจุดหัว-ท้ายที่คลิกเลือกเอง
if 'manual_points' not in st.session_state: st.session_state.manual_points = []
if 'export_data' not in st.session_state: st.session_state.export_data = []

st.markdown("""<style>
    .stApp { background: #FDFCFB; }
    .header-box { padding: 20px; background: white; border-radius: 20px; border-bottom: 5px solid #FF8C42; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
</style>""", unsafe_allow_html=True)

# Header
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f'''<div class="header-box"><div style="display: flex; align-items: center; justify-content: space-between;">
    <div><h1 style="margin:0; color: #2D5A27;">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA | คลิกเลือกจุดหัว-ท้ายบนแผนที่เพื่อวัดระยะ</p></div>
    {"<img src='data:image/png;base64,"+joker_base64+"' style='width:80px; border-radius:50%; border: 3px solid #FF8C42;'>" if joker_base64 else ""}
</div></div>''', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("⚙️ เมนูควบคุม")
    if st.button("🗑️ ล้างจุดหัว-ท้ายที่เลือกไว้"):
        st.session_state.manual_points = []
        st.rerun()
    st.divider()
    kml_file = st.file_uploader("🌐 อัปโหลด KML/KMZ", type=['kml', 'kmz'])
    uploaded_files = st.file_uploader("📁 อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

# Logic: KML/KMZ
kml_elements = []
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
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords[0].strip().split()]
                kml_elements.append({'name': name[0] if name else "Point", 'points': pts, 'is_point': len(pts) == 1})
    except: pass

# Logic: รูปภาพ
if uploaded_files:
    current_hash = "".join([f.name + str(f.size) for f in uploaded_files])
    if 'last_hash' not in st.session_state or st.session_state.last_hash != current_hash:
        st.session_state.export_data = []
        st.session_state.last_hash = current_hash
        for f in uploaded_files:
            raw_data = f.getvalue()
            raw_img = Image.open(BytesIO(raw_data))
            img_st = ImageOps.exif_transpose(raw_img)
            lat, lon = get_lat_lon_exif(raw_img)
            if lat:
                issue = analyze_cable_issue(raw_data)
                st.session_state.export_data.append({'img_obj': img_st, 'issue': issue, 'lat': lat, 'lon': lon})

# --- คำนวณเส้นทางเดิน (Walking) จากจุดที่เลือกเอง ---
manual_route, manual_dist = None, 0
if len(st.session_state.manual_points) == 2:
    manual_route, manual_dist = get_osrm_route(st.session_state.manual_points[0], st.session_state.manual_points[1])

# --- แสดงผลแผนที่ ---
st.subheader("🗺️ แผนที่สำรวจ (นครราชสีมา - ชัยภูมิ)")

# แสดงสถานะจุดที่เลือก
if len(st.session_state.manual_points) == 1:
    st.warning("📍 เลือกจุด 'หัว' แล้ว... กรุณาคลิกเลือกจุด 'ท้าย' บนแผนที่")
elif len(st.session_state.manual_points) == 2:
    st.success(f"📏 ระยะทางเดิน (หัว-ท้าย): {manual_dist:,.2f} เมตร ({manual_dist/1000:.3f} กม.)")

# ตั้งค่าแผนที่เริ่มต้นไปที่โคราช-ชัยภูมิ
m = folium.Map(location=[15.3, 101.8], zoom_start=9, 
               tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")

# 1. วาดเส้นทางที่คำนวณจากจุดหัว-ท้ายที่เลือกเอง (สีแดง)
if manual_route:
    folium.PolyLine(manual_route, color="#D9534F", weight=6, opacity=0.9, tooltip=f"ระยะทาง: {manual_dist:,.0f} ม.").add_to(m)

# 2. วาด Marker จุดหัว-ท้าย (เลข 1 และ 2)
for i, pt in enumerate(st.session_state.manual_points):
    color = 'green' if i == 0 else 'red'
    folium.Marker(pt, icon=folium.Icon(color=color, icon='info-sign'), popup=f"จุดที่ {i+1}").add_to(m)

# 3. วาด Marker รูปภาพสำรวจ
for data in st.session_state.export_data:
    folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=img_to_custom_icon(data['img_obj'], data['issue']))).add_to(m)

# 4. วาดเส้นโครงข่าย KML (สีเทาจาง)
for elem in kml_elements:
    if not elem['is_point']:
        folium.PolyLine(elem['points'], color="gray", weight=2, opacity=0.4).add_to(m)

m.add_child(MeasureControl(position='topright'))

# Render และดักจับการคลิกเพื่อปักจุดหัว-ท้าย
map_out = st_folium(m, height=750, use_container_width=True, key="main_map")

if map_out and map_out.get("last_clicked"):
    new_pt = [map_out["last_clicked"]["lat"], map_out["last_clicked"]["lng"]]
    # ถ้าเลือกครบ 2 จุดแล้ว คลิกครั้งต่อไปจะเริ่มนับใหม่ (Reset)
    if len(st.session_state.manual_points) >= 2:
        st.session_state.manual_points = [new_pt]
    elif not st.session_state.manual_points or new_pt != st.session_state.manual_points[-1]:
        st.session_state.manual_points.append(new_pt)
    st.rerun()

# --- Export ---
st.divider()
st.subheader("📄 3. สร้างรายงาน PowerPoint")
map_cap = st.file_uploader("📸 Capture แผนที่มาวาง", type=['jpg','png'])
if map_cap and st.button("🚀 ดาวน์โหลดรายงาน PPTX"):
    pptx = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
    st.download_button("📥 Click", data=pptx, file_name="Cable_Report.pptx")
