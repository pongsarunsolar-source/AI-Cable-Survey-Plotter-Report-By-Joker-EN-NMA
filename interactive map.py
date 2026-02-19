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
import math

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

# --- 2. ฟังก์ชันช่วยดึงรูปภาพ Joker ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except Exception: return None
    return None

# --- 3. ฟังก์ชันวิเคราะห์สาเหตุด้วย AI ---
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

# --- 4. ฟังก์ชันจัดการพิกัด ---
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
        img_for_ocr = image.copy()
        img_for_ocr.thumbnail((1000, 1000)) 
        img_np = np.array(img_for_ocr)
        results = reader.readtext(img_np, paragraph=True)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- ฟังก์ชันวัดระยะทาง (OSRM Multi-point Walking) ---
def get_osrm_multi_walking(coords_list):
    if not coords_list or len(coords_list) < 2: return None, 0
    coords_str = ";".join([f"{c[1]},{c[0]}" for c in coords_list])
    url = f"http://router.project-osrm.org/route/v1/walking/{coords_str}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if "routes" in data and len(data["routes"]) > 0:
                route = data["routes"][0]
                geometry = route["geometry"]["coordinates"]
                distance = route["distance"]
                folium_coords = [[lat, lon] for lon, lat in geometry]
                return folium_coords, distance
    except: pass
    return None, 0

def get_farthest_points(coordinates):
    if not coordinates or len(coordinates) < 2: return None, None
    max_dist = -1
    p1_best, p2_best = None, None
    for i in range(len(coordinates)):
        for j in range(i + 1, len(coordinates)):
            lat1, lon1 = coordinates[i]
            lat2, lon2 = coordinates[j]
            dist = (lat1 - lat2)**2 + (lon1 - lon2)**2
            if dist > max_dist:
                max_dist = dist
                p1_best, p2_best = coordinates[i], coordinates[j]
    return p1_best, p2_best

def get_osrm_route_head_tail(start_coord, end_coord):
    if not start_coord or not end_coord: return None, 0
    coords_str = f"{start_coord[1]},{start_coord[0]};{end_coord[1]},{end_coord[0]}"
    url = f"http://router.project-osrm.org/route/v1/walking/{coords_str}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if "routes" in data and len(data["routes"]) > 0:
                route = data["routes"][0]
                geometry = route["geometry"]["coordinates"]
                distance = route["distance"]
                folium_coords = [[lat, lon] for lon, lat in geometry]
                return folium_coords, distance
    except: pass
    return None, 0

# --- 5-7. ฟังก์ชันช่วยอื่นๆ (Label, Icon, PPTX) ---
def create_div_label(name):
    return f'<div style="font-size: 11px; font-weight: 800; color: #D9534F; text-shadow: 2px 2px 4px white;">{name}</div>'

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((150, 150)) 
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="background: white; padding: 5px; border-radius: 12px; border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 11px; font-weight: 700; color: #2D5A27; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width: 140px; border-radius: 4px;">
        </div>
    '''

def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    if map_image_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# --- 8. UI Layout ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")

# State สำหรับเก็บพิกัดที่คลิกวัดระยะ
if 'click_coords' not in st.session_state: st.session_state.click_coords = []
if 'export_data' not in st.session_state: st.session_state.export_data = []

# --- Header ---
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f"""<div style="display: flex; align-items: center; justify-content: space-between; padding: 20px; background: white; border-radius: 20px; border-bottom: 5px solid #FF8C42;">
    <div><h1 style="margin:0; color: #2D5A27;">AI Cable Plotter</h1><p style="margin:0; color: #718096;">Focus: Nakhon Ratchasima - Chaiyaphum</p></div>
    {'<img src="data:image/png;base64,'+joker_base64+'" style="width:80px; border-radius:50%;">' if joker_base64 else ''}
</div>""", unsafe_allow_html=True)

# --- 9-10. การจัดการไฟล์ ---
kml_file = st.sidebar.file_uploader("🌐 อัปโหลด KML/KMZ", type=['kml', 'kmz'])
uploaded_files = st.sidebar.file_uploader("📁 อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if st.sidebar.button("🗑️ ล้างระยะทางที่วัดใหม่"):
    st.session_state.click_coords = []
    st.rerun()

kml_elements, kml_points_pool = [], []
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
            coords = pm.xpath('.//kml:coordinates/text()', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords[0].strip().split()]
                kml_elements.append({'points': pts, 'is_point': len(pts) == 1})
                kml_points_pool.extend(pts)
    except: pass

if uploaded_files:
    # Logic เดิมในการจัดการรูปภาพ
    for f in uploaded_files:
        if not any(d['img_obj'].filename == f.name for d in st.session_state.export_data if hasattr(d['img_obj'], 'filename')):
            raw_data = f.getvalue()
            raw_img = Image.open(BytesIO(raw_data))
            img_st = ImageOps.exif_transpose(raw_img)
            lat, lon = get_lat_lon_exif(raw_img)
            if lat:
                issue = analyze_cable_issue(raw_data)
                st.session_state.export_data.append({'img_obj': img_st, 'issue': issue, 'lat': lat, 'lon': lon})

# --- คำนวณเส้นทาง ---
head_tail_route, dist_kml = get_osrm_route_head_tail(*get_farthest_points(kml_points_pool))
click_route, dist_click = get_osrm_multi_walking(st.session_state.click_coords)

# --- แสดงผลแผนที่ ---
st.subheader("📍 แผนที่สำรวจ (คลิกเพื่อวัดระยะทางเดิน)")

# โฟกัสไปที่ นครราชสีมา-ชัยภูมิ (พิกัดกลางโดยประมาณ)
m = folium.Map(
    location=[15.3, 101.8], zoom_start=9, 
    tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google"
)

# 1. วาดเส้น KMZ (สีน้ำเงิน)
if head_tail_route:
    folium.PolyLine(head_tail_route, color="#007BFF", weight=4, opacity=0.7, dash_array='10').add_to(m)

# 2. วาดเส้นวัดระยะจากการคลิก (สีแดง - เดินย้อนศรได้)
if click_route:
    folium.PolyLine(click_route, color="#D9534F", weight=6, opacity=0.9).add_to(m)
    st.sidebar.warning(f"📏 ระยะทางที่วัดใหม่: {dist_click:,.0f} เมตร")

# 3. วาด Marker ต่างๆ
for pt in st.session_state.click_coords:
    folium.CircleMarker(pt, radius=5, color='red', fill=True).add_to(m)

for data in st.session_state.export_data:
    folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=img_to_custom_icon(data['img_obj'], data['issue']))).add_to(m)

m.add_child(MeasureControl(position='topright'))

# Render และดักจับการคลิก
map_out = st_folium(m, height=800, use_container_width=True, key="survey_map")

if map_out and map_out.get("last_clicked"):
    clicked_pt = [map_out["last_clicked"]["lat"], map_out["last_clicked"]["lng"]]
    if not st.session_state.click_coords or clicked_pt != st.session_state.click_coords[-1]:
        st.session_state.click_coords.append(clicked_pt)
        st.rerun()

# --- Export ---
map_cap = st.file_uploader("📸 Capture แผนที่มาวาง", type=['jpg','png'])
if map_cap and st.button("🚀 สรุปรายงาน PPTX"):
    pptx = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
    st.download_button("📥 ดาวน์โหลด", pptx, "Cable_Report.pptx")
