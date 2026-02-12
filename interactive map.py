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

# --- 1. ตั้งค่า Google Gemini API ---
client = genai.Client(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")

@st.cache_resource
def load_ocr():
    model_path = os.path.join(os.getcwd(), "easyocr_models")
    if not os.path.exists(model_path):
        os.makedirs(model_path)
    return easyocr.Reader(['en'], gpu=False, model_storage_directory=model_path)

# --- 2. ฟังก์ชันคำนวณระยะทางเดินเท้า (OSRM) ---
def get_walking_distance(start_lat, start_lon, end_lat, end_lon):
    try:
        url = f"http://router.project-osrm.org/route/v1/foot/{start_lon},{start_lat};{end_lon},{end_lat}?overview=full&geometries=geojson"
        response = requests.get(url, timeout=5)
        data = response.json()
        if data['code'] == 'Ok':
            distance = data['routes'][0]['distance']
            geometry = data['routes'][0]['geometry']['coordinates']
            route_points = [[coord[1], coord[0]] for coord in geometry]
            return distance, route_points
    except: pass
    return None, None

# --- ฟังก์ชันช่วยอื่นๆ (Exif, OCR, Icon, PPTX) คงเดิมไว้ตามโครงสร้างหลักของคุณ ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200: return base64.b64encode(response.content).decode()
    except: return None
    return None

def analyze_cable_issue(image_bytes):
    try:
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=["""วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเฉพาะชื่อสาเหตุภาษาไทย: 1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ""",
                      types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg")]
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

def create_div_label(name):
    return f'<div style="font-size: 11px; font-weight: 800; color: #D9534F; white-space: nowrap; transform: translate(-50%, -150%); text-shadow: 2px 2px 4px white;">{name}</div>'

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((150, 150))
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''<div style="position: relative; background: white; padding: 5px; border-radius: 12px; border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
                <div style="font-size: 11px; font-weight: 700; color: #2D5A27; text-align: center;">{issue_text}</div>
                <img src="data:image/jpeg;base64,{img_str}" style="max-width: 140px; border-radius: 4px;">
              </div>'''

def create_summary_pptx(map_image_bytes, image_list, dist_text):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    if map_image_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
        tb = slide.shapes.add_textbox(Inches(0.2), Inches(0.2), Inches(4), Inches(0.5))
        tb.text_frame.text = f"ระยะทางหัว-ท้าย (KML): {dist_text}"
    output = BytesIO(); prs.save(output); return output.getvalue()

# --- 10. UI Layout ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.markdown("""<style> .stApp { background: #FFF5ED; } .main-title { font-weight: 800; font-size: 2.6rem; color: #2D5A27; } </style>""", unsafe_allow_html=True)

# Header
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f"<div><h1 class='main-title'>AI Cable Plotter</h1><p>By Joker EN-NMA</p></div>", unsafe_allow_html=True)

# --- 11. ส่วนจัดการ KML และระยะทาง ---
st.subheader("🌐 1. ข้อมูลโครงข่าย (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลด KML/KMZ เพื่อคำนวณระยะหัว-ท้าย", type=['kml', 'kmz'])

kml_points = []
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
            coords = pm.xpath('.//kml:coordinates/text() | .//earth:coordinates/text()', namespaces=ns)
            if coords:
                # เก็บเฉพาะพิกัดแรกของแต่ละ Placemark
                raw_coord = coords[0].strip().split()[0].split(',')
                lat, lon = float(raw_coord[1]), float(raw_coord[0])
                kml_points.append([lat, lon])
                
                name_node = pm.xpath('kml:name/text()', namespaces=ns)
                name = name_node[0] if name_node else "Point"
                kml_elements.append({'name': name, 'loc': [lat, lon]})
    except Exception as e: st.error(f"Error: {e}")

# --- 12. คำนวณระยะทางจาก KML ---
kml_dist_text = "0 เมตร"
kml_route = []

if len(kml_points) >= 2:
    # จุดแรกสุด และ จุดท้ายสุด ในไฟล์
    start_pt = kml_points[0]
    end_pt = kml_points[-1]
    
    dist, route = get_walking_distance(start_pt[0], start_pt[1], end_pt[0], end_pt[1])
    if dist:
        kml_dist_text = f"{dist:.2f} เมตร"
        kml_route = route
        st.sidebar.success(f"📏 ระยะหัว-ท้าย KML: {kml_dist_text}")
        st.sidebar.info(f"จาก: {kml_elements[0]['name']}\nถึง: {kml_elements[-1]['name']}")

# --- 13. แสดงแผนที่และรูปภาพ ---
uploaded_files = st.file_uploader("📁 2. อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

m = folium.Map(location=[13.75, 100.5], zoom_start=15, tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
all_bounds = []

# วาดจุด KML และเส้นทางเดินเท้า
if kml_elements:
    for i, elem in enumerate(kml_elements):
        color = 'red' if (i==0 or i==len(kml_elements)-1) else 'blue'
        folium.Marker(elem['loc'], icon=folium.Icon(color=color)).add_to(m)
        folium.Marker(elem['loc'], icon=folium.DivIcon(html=create_div_label(elem['name']))).add_to(m)
        all_bounds.append(elem['loc'])
    
    if kml_route:
        folium.PolyLine(kml_route, color="#2D5A27", weight=5, opacity=0.7, tooltip=f"ระยะทาง: {kml_dist_text}").add_to(m)

# จัดการรูปภาพสำรวจ
if uploaded_files:
    if 'export_data' not in st.session_state: st.session_state.export_data = []
    # (ส่วนประมวลผลรูปภาพ Gemini/OCR เหมือนโค้ดเดิมของคุณ)
    for f in uploaded_files:
        raw_data = f.getvalue()
        img = ImageOps.exif_transpose(Image.open(BytesIO(raw_data)))
        lat, lon = get_lat_lon_exif(img)
        if lat:
            issue = analyze_cable_issue(raw_data)
            icon_html = img_to_custom_icon(img, issue)
            folium.Marker([lat, lon], icon=folium.DivIcon(html=icon_html)).add_to(m)
            all_bounds.append([lat, lon])

if all_bounds: m.fit_bounds(all_bounds)
st_folium(m, height=700, use_container_width=True)

# ส่วน Export PPTX
if st.button("🚀 สรุปรายงาน PPTX"):
    # ดึงรูปแผนที่จากไฟล์ที่อัปโหลด (ถ้ามี)
    pptx_data = create_summary_pptx(None, None, kml_dist_text)
    st.download_button("📥 ดาวน์โหลดรายงาน", data=pptx_data, file_name="Report.pptx")
