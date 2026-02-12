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

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า API และ OCR ---
API_KEY = st.secrets.get("GEMINI_API_KEY", "AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
genai.configure(api_key=API_KEY)
model_ai = genai.GenerativeModel('gemini-1.5-flash')

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# --- 2. ฟังก์ชันคำนวณเส้นทางตามแนวถนน (OSRM) ---
@st.cache_data(show_spinner="กำลังคำนวณเส้นทางตามแนวถนน...")
def get_route_on_roads(points):
    """ส่งพิกัดไปขอเส้นทางที่ลากตามถนนจริงจาก OSRM"""
    if len(points) < 2: return points
    try:
        # รวมพิกัดเป็น string format: lon,lat;lon,lat
        coord_str = ";".join([f"{p[1]},{p[0]}" for p in points])
        url = f"http://router.project-osrm.org/route/v1/driving/{coord_str}?overview=full&geometries=geojson"
        response = requests.get(url, timeout=10)
        data = response.json()
        if data['code'] == 'Ok':
            geometry = data['routes'][0]['geometry']['coordinates']
            return [[p[1], p[0]] for p in geometry]
    except:
        pass
    return points # ถ้าล้มเหลว ให้คืนค่าเส้นตรงเดิม

# --- 3. ฟังก์ชันประมวลผลรูปภาพและ AI ---
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
        results = reader.readtext(img_np)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 4. ฟังก์ชันจัดการ UI แผนที่ ---
def create_div_label(name):
    return f'<div style="font-size:11px; font-weight:800; color:#D9534F; text-shadow:2px 2px 4px white; white-space:nowrap;">{name}</div>'

def img_to_icon(img, issue):
    thumb = img.copy()
    thumb.thumbnail((120, 120))
    buf = BytesIO()
    thumb.save(buf, format="JPEG", quality=60)
    img_b64 = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="width:130px; background:white; padding:5px; border-radius:8px; border:2px solid #FF8C42;">
            <div style="font-size:10px; font-weight:bold; color:#2D5A27; text-align:center;">{issue}</div>
            <img src="data:image/jpeg;base64,{img_b64}" style="width:100%; border-radius:4px;">
        </div>
    '''

# --- 5. Main UI ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.title("🔌 AI Cable Plotter & Road Router")
st.caption("Version: Optimized for Streamlit Cloud (CPU) | ลากเส้นตามถนนอัตโนมัติ")

with st.sidebar:
    st.header("Settings & Upload")
    kml_file = st.file_uploader("1. KML/KMZ Network", type=['kml', 'kmz'])
    img_files = st.file_uploader("2. Survey Photos", type=['jpg','jpeg','png'], accept_multiple_files=True)
    st.divider()
    map_cap = st.file_uploader("3. Capture แผนที่เพื่อทำรายงาน", type=['jpg','png'])
    
    if st.button("🚀 Create PPTX Report"):
        if 'survey_results' in st.session_state and map_cap:
            # (ฟังก์ชันสร้าง PPTX แบบประหยัด RAM)
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
            slide.shapes.add_picture(BytesIO(map_cap.getvalue()), 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            for item in st.session_state.survey_results[:12]: # จำกัดจำนวนรูปในรายงาน
                s = prs.slides.add_slide(prs.slide_layouts[6])
                img_buf = BytesIO()
                item['img'].save(img_buf, format="JPEG")
                s.shapes.add_picture(img_buf, Inches(1), Inches(1), width=Inches(8))
                tb = s.shapes.add_textbox(Inches(1), Inches(4.7), Inches(8), Inches(1))
                tb.text = f"Issue: {item['issue']} | Location: {item['lat']},{item['lon']}"
                
            out = BytesIO()
            prs.save(out)
            st.download_button("📥 Download PPTX", out.getvalue(), "Report.pptx")

# --- 6. แผนที่และการประมวลผล ---
m = folium.Map(location=[13.75, 100.5], zoom_start=14, tiles="cartodbpositron")
all_bounds = []
st.session_state.survey_results = []

# จัดการ KML และลากเส้นตามถนน
if kml_file:
    try:
        content = kml_file.getvalue()
        if kml_file.name.endswith('.kmz'):
            with zipfile.ZipFile(BytesIO(content)) as z:
                content = z.read([n for n in z.namelist() if n.endswith('.kml')][0])
        
        root = etree.fromstring(content)
        ns = {'kml': 'http://www.opengis.net/kml/2.2'}
        for pm in root.xpath('.//kml:Placemark', namespaces=ns):
            name = pm.findtext('.//kml:name', default="", namespaces=ns)
            coords_text = pm.findtext('.//kml:coordinates', namespaces=ns)
            if coords_text:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords_text.strip().split()]
                
                if len(pts) == 1:
                    folium.Marker(pts[0], icon=folium.Icon(color='red')).add_to(m)
                    folium.Marker(pts[0], icon=folium.DivIcon(html=create_div_label(name))).add_to(m)
                    all_bounds.append(pts[0])
                else:
                    # ลากเส้นตามถนนจริง
                    road_pts = get_route_on_roads(pts)
                    folium.PolyLine(road_pts, color="#0078FF", weight=5, opacity=0.8).add_to(m)
                    all_bounds.extend(road_pts)
    except: st.error("KML Processing Error")

# จัดการรูปถ่าย
if img_files:
    for f in img_files[:20]: # จำกัดจำนวนรูปเพื่อความเร็ว
        img_raw = Image.open(f)
        img_fixed = ImageOps.exif_transpose(img_raw)
        lat, lon = get_lat_lon_exif(img_raw)
        
        buf = BytesIO()
        img_fixed.save(buf, format="JPEG", quality=60)
        img_bytes = buf.getvalue()

        if lat is None: lat, lon = get_lat_lon_ocr_cached(img_bytes)

        if lat:
            issue = analyze_cable_issue_cached(img_bytes)
            icon_html = img_to_icon(img_fixed, issue)
            folium.Marker([lat, lon], icon=folium.DivIcon(html=icon_html)).add_to(m)
            all_bounds.append([lat, lon])
            st.session_state.survey_results.append({'img':img_fixed, 'issue':issue, 'lat':lat, 'lon':lon})

if all_bounds: m.fit_bounds(all_bounds)
st_folium(m, width="100%", height=800)
