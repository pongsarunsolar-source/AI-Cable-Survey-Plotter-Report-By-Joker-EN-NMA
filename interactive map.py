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
import google.generativeai as genai
import zipfile
from lxml import etree
import pandas as pd

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า API และ OCR ---
# แนะนำให้ใช้ st.secrets["GEMINI_API_KEY"] บน Cloud
API_KEY = st.secrets.get("GEMINI_API_KEY", "AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
genai.configure(api_key=API_KEY)
model_ai = genai.GenerativeModel('gemini-1.5-flash')

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# --- 2. ฟังก์ชันประมวลผลพิกัดและเส้นทาง ---
@st.cache_data(show_spinner="กำลังคำนวณแนวถนน...")
def get_route_on_roads(points):
    if len(points) < 2: return points
    try:
        coord_str = ";".join([f"{p[1]},{p[0]}" for p in points])
        url = f"http://router.project-osrm.org/route/v1/driving/{coord_str}?overview=full&geometries=geojson"
        response = requests.get(url, timeout=10)
        data = response.json()
        if data['code'] == 'Ok':
            geometry = data['routes'][0]['geometry']['coordinates']
            return [[p[1], p[0]] for p in geometry]
    except: pass
    return points

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

@st.cache_data(show_spinner=False)
def analyze_cable_issue_cached(img_bytes):
    try:
        img = Image.open(BytesIO(img_bytes))
        prompt = "วิเคราะห์รูปสายเคเบิล ตอบเพียง: 1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ"
        response = model_ai.generate_content([prompt, img])
        return response.text.strip()
    except: return "วิเคราะห์ไม่ได้"

# --- 3. UI Styling (เดิม) ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
    .joker-icon { width: 80px; height: 80px; border-radius: 50%; border: 3px solid #FF8C42; }
</style>""", unsafe_allow_html=True)

# Header
joker_url = "https://drive.google.com/uc?export=download&id=1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr"
st.markdown(f'''<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA</p></div><img src="{joker_url}" class="joker-icon"></div>''', unsafe_allow_html=True)

# --- 4. Main Logic ---
st.subheader("🌐 1. ข้อมูลโครงข่าย (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลดไฟล์ KML หรือ KMZ", type=['kml', 'kmz'], key="kml_upload")

kml_elements = []
if kml_file:
    try:
        content = kml_file.getvalue()
        if kml_file.name.endswith('.kmz'):
            with zipfile.ZipFile(BytesIO(content)) as z:
                content = z.read([n for n in z.namelist() if n.endswith('.kml')][0])
        root = etree.fromstring(content)
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for pm in root.xpath('.//kml:Placemark', namespaces=ns):
            name = pm.findtext('.//kml:name', default="Point", namespaces=ns)
            coords = pm.findtext('.//kml:coordinates', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords.strip().split()]
                kml_elements.append({'name': name, 'points': pts, 'is_point': len(pts) == 1})
    except: st.error("ไฟล์ KML ไม่ถูกต้อง")

st.markdown("<hr>", unsafe_allow_html=True)

st.subheader("📁 2. อัปโหลดรูปภาพสำรวจ")
uploaded_files = st.file_uploader("เลือกรูปภาพ...", type=['jpg','jpeg','png'], accept_multiple_files=True, key="img_upload")

# --- 5. การสร้างแผนที่ (จุดที่แก้ไขเรื่อง Map ไม่ขึ้น) ---
# กำหนดจุดเริ่มต้นเริ่มต้นกรณีไม่มีข้อมูล
start_lat, start_lon = 13.75, 100.5
m = folium.Map(location=[start_lat, start_lon], zoom_start=6, tiles="cartodbpositron")
all_bounds = []

# ประมวลผลและวาด KML
for elem in kml_elements:
    if elem['is_point']:
        folium.Marker(elem['points'][0], tooltip=elem['name'], icon=folium.Icon(color='red')).add_to(m)
        all_bounds.append(elem['points'][0])
    else:
        road_route = get_route_on_roads(elem['points'])
        folium.PolyLine(road_route, color="#0078FF", weight=5, opacity=0.7).add_to(m)
        all_bounds.extend(road_route)

# ประมวลผลรูปภาพ
survey_data = []
if uploaded_files:
    for f in uploaded_files:
        img_raw = Image.open(f)
        img_st = ImageOps.exif_transpose(img_raw)
        lat, lon = get_lat_lon_exif(img_raw)
        
        if lat:
            # วิเคราะห์ AI
            buf = BytesIO()
            img_st.save(buf, format="JPEG", quality=50)
            issue = analyze_cable_issue_cached(buf.getvalue())
            
            # สร้าง Icon
            img_thumb = img_st.copy()
            img_thumb.thumbnail((120, 120))
            thumb_buf = BytesIO()
            img_thumb.save(thumb_buf, format="JPEG")
            img_b64 = base64.b64encode(thumb_buf.getvalue()).decode()
            
            icon_html = f'''<div style="width:130px; background:white; padding:5px; border-radius:8px; border:2px solid #FF8C42; text-align:center;">
                            <div style="font-size:10px; font-weight:bold; color:#2D5A27;">{issue}</div>
                            <img src="data:image/jpeg;base64,{img_b64}" style="width:100%; border-radius:4px;"></div>'''
            
            folium.Marker([lat, lon], icon=folium.DivIcon(html=icon_html)).add_to(m)
            all_bounds.append([lat, lon])
            survey_data.append({'img': img_st, 'issue': issue, 'lat': lat, 'lon': lon})

# จัดการการแสดงผลแผนที่
if all_bounds:
    m.fit_bounds(all_bounds)

# แสดงผลแผนที่ (ใช้ Use Container Width เพื่อให้เต็มหน้าจอ)
st_folium(m, width="100%", height=700, key="main_map")

# --- 6. Export PowerPoint ---
st.markdown("<hr>", unsafe_allow_html=True)
if survey_data:
    st.subheader("📄 3. สร้างรายงาน PowerPoint")
    map_cap = st.file_uploader("อัปโหลดภาพ Capture แผนที่", type=['jpg','png'])
    if map_cap and st.button("🚀 ดาวน์โหลดรายงาน"):
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)
        
        # หน้าแรก: แผนที่
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_cap.getvalue()), 0, 0, width=prs.slide_width)
        
        # หน้าต่อๆ ไป: รูปสำรวจ
        for item in survey_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            buf = BytesIO()
            item['img'].save(buf, format="JPEG")
            slide.shapes.add_picture(BytesIO(buf.getvalue()), Inches(0.5), Inches(0.5), height=Inches(4))
            tx = slide.shapes.add_textbox(Inches(0.5), Inches(4.7), Inches(9), Inches(1))
            tx.text = f"สาเหตุ: {item['issue']} | พิกัด: {item['lat']:.5f}, {item['lon']:.5f}"
            
        ppt_buf = BytesIO()
        prs.save(ppt_buf)
        st.download_button("📥 คลิกเพื่อโหลด PPTX", ppt_buf.getvalue(), "Cable_Report.pptx")
