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

# --- 1. ตั้งค่า Google Gemini API ---
API_KEY = st.secrets.get("GEMINI_API_KEY", "AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
genai.configure(api_key=API_KEY)
model_ai = genai.GenerativeModel('gemini-1.5-flash')

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# --- 2. ฟังก์ชันเสริม (Road Routing & Drive Image) ---
@st.cache_data(show_spinner="กำลังลากเส้นตามแนวถนน...")
def get_route_on_roads(points):
    """เรียก OSRM API เพื่อลากเส้นตามถนนจริง"""
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

def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except: return None
    return None

# --- 3. ฟังก์ชันวิเคราะห์ด้วย AI & พิกัด ---
@st.cache_data(show_spinner=False)
def analyze_cable_issue_cached(img_bytes):
    try:
        img = Image.open(BytesIO(img_bytes))
        prompt = """วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเพียง "หนึ่งเดียว" จาก 4 สาเหตุ:
        1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ
        ตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น"""
        response = model_ai.generate_content([prompt, img])
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

@st.cache_data(show_spinner=False)
def get_lat_lon_ocr_cached(img_bytes):
    try:
        img_np = np.array(Image.open(BytesIO(img_bytes)))
        results = reader.readtext(img_np, paragraph=True)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 4. ฟังก์ชันจัดการหน้าตาแผนที่ ---
def create_div_label(name):
    return f'''<div style="font-size: 11px; font-weight: 800; color: #D9534F; white-space: nowrap; transform: translate(-50%, -150%); text-shadow: 2px 2px 4px white;">{name}</div>'''

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((150, 150)) 
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: fit-content; background-color: white; padding: 5px; border-radius: 12px; box-shadow: 0px 8px 24px rgba(0,0,0,0.12); border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 11px; font-weight: 700; color: #2D5A27; margin-bottom: 4px; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width: 140px; display: block; border-radius: 4px;">
            <div style="position: absolute; bottom: -10px; left: 50%; transform: translateX(-50%); width: 0; height: 0; border-left: 10px solid transparent; border-right: 10px solid transparent; border-top: 10px solid #FF8C42;"></div>
        </div>
    '''

def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    if map_image_bytes:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        slide1.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    if image_list:
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        cols, rows = 4, 2
        img_w, img_h = Inches(2.1), Inches(1.5)
        for i, item in enumerate(image_list[:8]):
            curr_row, curr_col = i // cols, i % cols
            x = (curr_col * (img_w + Inches(0.2))) + Inches(0.5)
            y = (curr_row * (img_h + Inches(0.7))) + Inches(0.5)
            buf = BytesIO()
            item['img_obj'].save(buf, format="JPEG")
            slide2.shapes.add_picture(BytesIO(buf.getvalue()), x, y, width=img_w, height=img_h)
            txt = slide2.shapes.add_textbox(x, y + img_h, img_w, Inches(0.5))
            txt.text = f"{item['issue']}\nLat:{item['lat']:.4f}"
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# --- 5. UI Layout (เดิม) ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
    .joker-icon { width: 100px; height: 100px; object-fit: cover; border-radius: 50%; border: 4px solid #FFFFFF; outline: 3px solid #FF8C42; }
</style>""", unsafe_allow_html=True)

# Header
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
header_html = f'''<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA</p></div>
{"<img src='data:image/png;base64,"+joker_base64+"' class='joker-icon'>" if joker_base64 else ""}</div>'''
st.markdown(header_html, unsafe_allow_html=True)

# --- 6. ประมวลผล KML/KMZ ---
st.subheader("🌐 1. ข้อมูลโครงข่าย (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลดไฟล์ KML หรือ KMZ", type=['kml', 'kmz'])
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
            name = pm.findtext('.//kml:name', default="ไม่ระบุชื่อ", namespaces=ns)
            coords = pm.findtext('.//kml:coordinates', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords.strip().split()]
                kml_elements.append({'name': name, 'points': pts, 'is_point': len(pts) == 1})
    except Exception as e: st.error(f"Error KML: {e}")

st.markdown("<hr>", unsafe_allow_html=True)

# --- 7. แผนที่และการประมวลผลรูปภาพ ---
uploaded_files = st.file_uploader("📁 2. อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if 'export_data' not in st.session_state: st.session_state.export_data = []

if uploaded_files or kml_elements:
    m = folium.Map(location=[13.75, 100.5], zoom_start=15, tiles="cartodbpositron")
    all_bounds = []

    # Plot KML (ลากเส้นตามแนวถนน)
    for elem in kml_elements:
        if elem['is_point']:
            loc = elem['points'][0]
            folium.Marker(loc, icon=folium.Icon(color='red', icon='info-sign')).add_to(m)
            folium.Marker(loc, icon=folium.DivIcon(html=create_div_label(elem['name']))).add_to(m)
            all_bounds.append(loc)
        else:
            # ใช้ Road Routing สำหรับเส้นทาง
            road_route = get_route_on_roads(elem['points'])
            folium.PolyLine(road_route, color="#0078FF", weight=5, opacity=0.8, tooltip=elem['name']).add_to(m)
            all_bounds.extend(road_route)

    # Plot รูปภาพ
    if uploaded_files:
        current_hash = "".join([f.name + str(f.size) for f in uploaded_files])
        if 'last_hash' not in st.session_state or st.session_state.last_hash != current_hash:
            st.session_state.export_data = []
            st.session_state.last_hash = current_hash

        for i, f in enumerate(uploaded_files[:20]): # จำกัด 20 รูป
            if i >= len(st.session_state.export_data):
                img_raw = Image.open(f)
                img_st = ImageOps.exif_transpose(img_raw)
                lat, lon = get_lat_lon_exif(img_raw)
                buf = BytesIO()
                img_st.save(buf, format="JPEG", quality=60)
                img_bytes = buf.getvalue()
                if lat is None: lat, lon = get_lat_lon_ocr_cached(img_bytes)
                if lat:
                    issue = analyze_cable_issue_cached(img_bytes)
                    st.session_state.export_data.append({'img_obj': img_st, 'issue': issue, 'lat': lat, 'lon': lon})

            if i < len(st.session_state.export_data):
                data = st.session_state.export_data[i]
                icon_html = img_to_custom_icon(data['img_obj'], data['issue'])
                folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=icon_html)).add_to(m)
                all_bounds.append([data['lat'], data['lon']])

    if all_bounds: m.fit_bounds(all_bounds, padding=[50, 50])
    st_folium(m, height=850, use_container_width=True)

    # --- 8. Export PowerPoint ---
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📄 3. สร้างรายงาน PowerPoint")
    col1, col2 = st.columns([1, 1])
    with col1:
        map_cap = st.file_uploader("อัปโหลดรูป Capture แผนที่", type=['jpg','png'])
    if map_cap and st.session_state.export_data:
        with col2:
            if st.button("🚀 ดาวน์โหลดไฟล์ PPTX"):
                pptx_data = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
                st.download_button("📥 คลิกเพื่อดาวน์โหลด", data=pptx_data, file_name="Cable_Report.pptx")
