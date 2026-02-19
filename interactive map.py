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

# --- 1. ตั้งค่า API & OCR ---
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

def get_lat_lon_ocr(image):
    try:
        reader = load_ocr() 
        img_np = np.array(image.convert('RGB'))
        results = reader.readtext(img_np, paragraph=True)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 3. ฟังก์ชันวัดระยะทาง (OSRM Walking) ---
def get_osrm_route_walking(coords_list):
    """
    คำนวณเส้นทางเดินผ่านทุกจุดใน List (ย้อนศรได้)
    coords_list: [[lat, lon], [lat, lon], ...]
    """
    if not coords_list or len(coords_list) < 2:
        return None, 0
    
    # OSRM format: lon,lat;lon,lat
    coords_str = ";".join([f"{c[1]},{c[0]}" for c in coords_list])
    url = f"http://router.project-osrm.org/route/v1/walking/{coords_str}?overview=full&geometries=geojson"
    
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if "routes" in data and len(data["routes"]) > 0:
                route = data["routes"][0]
                geometry = route["geometry"]["coordinates"]
                distance = route["distance"] # เมตร
                folium_coords = [[lat, lon] for lon, lat in geometry]
                return folium_coords, distance
    except: pass
    return None, 0

# --- 4. ฟังก์ชันจัดการ UI/Icons ---
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
            <div style="font-size: 10px; font-weight: bold; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="width: 120px; border-radius: 5px;">
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

# Session State สำหรับการคลิกวัดระยะ
if 'click_coords' not in st.session_state: st.session_state.click_coords = []
if 'export_data' not in st.session_state: st.session_state.export_data = []

st.markdown("""<style>
    .stApp { background: #F8F9FA; }
    .header { padding: 20px; background: white; border-radius: 15px; border-left: 10px solid #FF8C42; margin-bottom: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.05); }
</style>""", unsafe_allow_html=True)

# Header
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f'''<div class="header"><h1>📡 AI Cable Plotter</h1><p>By Joker EN-NMA | คลิกบนแผนที่เพื่อวัดระยะทางเดิน</p></div>''', unsafe_allow_html=True)

# Sidebar Control
with st.sidebar:
    st.header("⚙️ Controls")
    if st.button("🗑️ ล้างจุดที่เลือกวัดระยะ"):
        st.session_state.click_coords = []
        st.rerun()
    
    st.divider()
    kml_file = st.file_uploader("🌐 อัปโหลด KML/KMZ", type=['kml', 'kmz'])
    uploaded_files = st.file_uploader("📁 อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

# --- Logic: จัดการไฟล์ KML ---
kml_elements = []
if kml_file:
    try:
        if kml_file.name.endswith('.kmz'):
            with zipfile.ZipFile(kml_file) as z:
                kml_filename = [n for n in z.namelist() if n.endswith('.kml')][0]
                content = z.read(kml_filename)
        else: content = kml_file.getvalue()
        root = etree.fromstring(content)
        ns = {'kml': 'http://www.opengis.net/kml/2.2', 'earth': 'http://earth.google.com/kml/2.2'}
        placemarks = root.xpath('.//kml:Placemark | .//earth:Placemark', namespaces=ns)
        for pm in placemarks:
            name = pm.xpath('kml:name/text() | earth:name/text()', namespaces=ns)
            coords = pm.xpath('.//kml:coordinates/text() | .//earth:coordinates/text()', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords[0].strip().split()]
                kml_elements.append({'name': name[0] if name else "Point", 'points': pts, 'is_point': len(pts) == 1})
    except Exception as e: st.error(f"KML Error: {e}")

# --- Logic: จัดการรูปภาพ ---
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
            if lat is None: lat, lon = get_lat_lon_ocr(img_st)
            if lat:
                issue = analyze_cable_issue(raw_data)
                st.session_state.export_data.append({'img_obj': img_st, 'issue': issue, 'lat': lat, 'lon': lon})

# --- แสดงผลแผนที่ ---
col_map, col_info = st.columns([4, 1])

with col_map:
    # เริ่มต้น Map
    start_lat = st.session_state.export_data[0]['lat'] if st.session_state.export_data else 13.75
    start_lon = st.session_state.export_data[0]['lon'] if st.session_state.export_data else 100.5
    
    m = folium.Map(location=[start_lat, start_lon], zoom_start=17, 
                   tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
    
    # 1. วาดเส้นทางจากการ "คลิก" (Walking Profile)
    if len(st.session_state.click_coords) >= 2:
        route_line, total_dist = get_osrm_route_walking(st.session_state.click_coords)
        if route_line:
            folium.PolyLine(route_line, color="#D9534F", weight=6, opacity=0.8, dash_array='10').add_to(m)
            st.info(f"📏 ระยะทางเดินรวมที่เลือก: {total_dist:,.2f} เมตร ({total_dist/1000:.3f} กม.)")

    # 2. วาด Marker จุดที่คลิก
    for i, pt in enumerate(st.session_state.click_coords):
        folium.CircleMarker(pt, radius=6, color='red', fill=True, fill_opacity=0.9, popup=f"จุดที่ {i+1}").add_to(m)

    # 3. วาดรูปถ่ายจากการสำรวจ
    for data in st.session_state.export_data:
        folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=img_to_custom_icon(data['img_obj'], data['issue']))).add_to(m)

    # 4. วาดข้อมูล KML
    for elem in kml_elements:
        if elem['is_point']:
            folium.Marker(elem['points'][0], icon=folium.Icon(color='blue')).add_to(m)
        else:
            folium.PolyLine(elem['points'], color="blue", weight=2, opacity=0.5).add_to(m)

    m.add_child(MeasureControl(position='topright'))
    
    # Render แผนที่และดักจับการคลิก
    map_out = st_folium(m, height=700, use_container_width=True, key="main_map")

    # ส่วนดักจับ Click Event
    if map_out and map_out.get("last_clicked"):
        new_pt = [map_out["last_clicked"]["lat"], map_out["last_clicked"]["lng"]]
        if not st.session_state.click_coords or new_pt != st.session_state.click_coords[-1]:
            st.session_state.click_coords.append(new_pt)
            st.rerun()

with col_info:
    st.subheader("📊 รายการ")
    for i, d in enumerate(st.session_state.export_data):
        st.write(f"{i+1}. {d['issue']}")
    
    st.divider()
    map_cap = st.file_uploader("📸 Capture แผนที่มาวาง", type=['jpg','png'])
    if map_cap and st.button("🎁 Download PPTX"):
        pptx = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
        st.download_button("📥 Click", pptx, "Report.pptx")
