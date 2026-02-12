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

# --- 2. ฟังก์ชันคำนวณเส้นทางถนน (รองรับการย้อนซ้อน) ---
@st.cache_data
def get_road_route_with_backtrack(points):
    """คำนวณเส้นทางตามพิกัด KML โดยไม่ตัดเส้นทางที่วิ่งย้อนกลับ"""
    if len(points) < 2: return points, 0
    
    # OSRM Service (ใช้ Profile 'route' เพื่อรักษาลำดับพิกัดตาม KML จริง)
    coords_str = ";".join([f"{p[1]},{p[0]}" for p in points])
    url = f"http://router.project-osrm.org/route/v1/driving/{coords_str}?overview=full&geometries=geojson&continue_straight=true"
    
    try:
        r = requests.get(url, timeout=15)
        data = r.json()
        if data['code'] == 'Ok':
            # ดึงพิกัดถนน
            route_coords = [[c[1], c[0]] for c in data['routes'][0]['geometry']['coordinates']]
            # ระยะทางรวม (กิโลเมตร)
            distance_km = data['routes'][0]['distance'] / 1000.0
            return route_coords, distance_km
    except:
        pass
    return points, 0

# --- ฟังก์ชันช่วยดึงรูปภาพ Joker ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except Exception: return None
    return None

# --- ฟังก์ชันวิเคราะห์สาเหตุด้วย AI ---
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

# --- ฟังก์ชันจัดการพิกัด ---
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

# --- UI Helpers ---
def create_div_label(name, color="#D9534F"):
    return f'<div style="font-size: 11px; font-weight: 800; color: {color}; white-space: nowrap; transform: translate(-50%, -150%); text-shadow: 2px 2px 4px white;">{name}</div>'

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy(); img_resized.thumbnail((150, 150))
    buf = BytesIO(); img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: fit-content; background: white; padding: 5px; border-radius: 12px; box-shadow: 0px 8px 24px rgba(0,0,0,0.12); border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 11px; font-weight: 700; color: #2D5A27; margin-bottom: 4px; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width: 140px; display: block; border-radius: 4px;">
            <div style="position: absolute; bottom: -10px; left: 50%; transform: translateX(-50%); width: 0; height: 0; border-left: 10px solid transparent; border-right: 10px solid transparent; border-top: 10px solid #FF8C42;"></div>
        </div>
    '''

def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation(); prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    if map_image_bytes:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        slide1.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    for i, item in enumerate(image_list[:8]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        x, y = Inches(2.5), Inches(1.0)
        buf = BytesIO(); item['img_obj'].save(buf, format="JPEG")
        slide.shapes.add_picture(buf, x, y, width=Inches(5), height=Inches(3.5))
        txt = slide.shapes.add_textbox(Inches(1), Inches(4.7), Inches(8), Inches(0.5)).text_frame
        p = txt.paragraphs[0]; p.text = f"สาเหตุ: {item['issue']} | พิกัด: {item['lat']:.5f}, {item['lon']:.5f}"; p.font.size = Pt(14)
    output = BytesIO(); prs.save(output); return output.getvalue()

# --- 8. UI Layout ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
    .joker-icon { width: 100px; height: 100px; object-fit: cover; border-radius: 50%; border: 4px solid #FFFFFF; outline: 3px solid #FF8C42; }
    .metric-card { background: white; padding: 15px; border-radius: 15px; border-left: 8px solid #2ECC71; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
</style>""", unsafe_allow_html=True)

joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f'<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA</p></div>{"<img src=\'data:image/png;base64,"+joker_base64+"\' class=\'joker-icon\'>" if joker_base64 else ""}</div>', unsafe_allow_html=True)

# --- 9. เมนู KML/KMZ (คำนวณจุดหัว-ท้าย) ---
st.subheader("🌐 1. ข้อมูลโครงข่าย & จุดติดตั้ง (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลดไฟล์ KML หรือ KMZ", type=['kml', 'kmz'])
kml_elements = []
all_kml_pts = []

if kml_file:
    try:
        content = kml_file.getvalue()
        if kml_file.name.endswith('.kmz'):
            with zipfile.ZipFile(BytesIO(content)) as z:
                content = z.read([n for n in z.namelist() if n.endswith('.kml')][0])
        root = etree.fromstring(content)
        ns = {'kml': 'http://www.opengis.net/kml/2.2', 'earth': 'http://earth.google.com/kml/2.2'}
        placemarks = root.xpath('.//kml:Placemark | .//earth:Placemark', namespaces=ns)
        for pm in placemarks:
            name = pm.xpath('kml:name/text() | earth:name/text()', namespaces=ns)
            coords = pm.xpath('.//kml:coordinates/text() | .//earth:coordinates/text()', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords[0].strip().split()]
                all_kml_pts.extend(pts)
                kml_elements.append({'name': name[0].strip() if name else "จุดสำรวจ", 'points': pts, 'is_point': len(pts) == 1})
    except: st.error("ไม่สามารถอ่านไฟล์ KML ได้")

st.markdown("<hr>", unsafe_allow_html=True)

# --- 10. ส่วนการทำงานหลัก ---
uploaded_files = st.file_uploader("📁 2. อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if 'export_data' not in st.session_state: st.session_state.export_data = []

if uploaded_files or kml_elements:
    m = folium.Map(location=[13.75, 100.5], zoom_start=17, tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
    all_bounds = []
    total_dist = 0.0

    # คำนวณหัว-ท้าย และระยะทางถนน (รองรับการย้อนซ้อน)
    if all_kml_pts:
        first_pt = all_kml_pts[0]
        last_pt = all_kml_pts[-1]
        
        # UI Dashboard
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.markdown(f"<div class='metric-card' style='border-left-color: #2D5A27;'><b>📍 มุดแรก (Start)</b><br>{first_pt[0]:.6f}, {first_pt[1]:.6f}</div>", unsafe_allow_html=True)
        with col_m2:
            st.markdown(f"<div class='metric-card' style='border-left-color: #000;'><b>🏁 มุดสุดท้าย (End)</b><br>{last_pt[0]:.6f}, {last_pt[1]:.6f}</div>", unsafe_allow_html=True)
        
        with st.spinner("กำลังคำนวณเส้นทางถนน (รวมระยะย้อนซ้อน)..."):
            road_pts, dist = get_road_route_with_backtrack(all_kml_pts)
            total_dist = dist
            
            # วาดเส้นหลัก
            folium.PolyLine(road_pts, color="#2ECC71", weight=8, opacity=0.8).add_to(m)
            # ปักหมุด Start/End
            folium.Marker(first_pt, icon=folium.Icon(color='green', icon='play'), popup="START").add_to(m)
            folium.Marker(last_pt, icon=folium.Icon(color='black', icon='stop'), popup="END").add_to(m)
            all_bounds.extend(road_pts)
        
        with col_m3:
            st.markdown(f"<div class='metric-card' style='border-left-color: #FF8C42;'><b>📏 ระยะทางถนนสุทธิ</b><br>{total_dist:.3f} กม.</div>", unsafe_allow_html=True)

    # วาดมุด KML อื่นๆ
    for elem in kml_elements:
        if elem['is_point']:
            folium.Marker(elem['points'][0], icon=folium.DivIcon(html=create_div_label(elem['name']))).add_to(m)
            all_bounds.append(elem['points'][0])

    # จัดการรูปภาพสำรวจ
    if uploaded_files:
        current_hash = "".join([f.name + str(f.size) for f in uploaded_files])
        if st.session_state.get('last_hash') != current_hash:
            st.session_state.export_data = []; st.session_state.last_hash = current_hash
            for f in uploaded_files:
                fb = f.getvalue(); img = ImageOps.exif_transpose(Image.open(BytesIO(fb)))
                lat, lon = get_lat_lon_exif(img)
                if lat is None: lat, lon = get_lat_lon_ocr(img)
                if lat:
                    issue = analyze_cable_issue(fb)
                    st.session_state.export_data.append({'img_obj': img, 'issue': issue, 'lat': lat, 'lon': lon})
        
        for d in st.session_state.export_data:
            folium.Marker([d['lat'], d['lon']], icon=folium.DivIcon(html=img_to_custom_icon(d['img_obj'], d['issue']))).add_to(m)
            all_bounds.append([d['lat'], d['lon']])

    if all_bounds: m.fit_bounds(all_bounds, padding=[50, 50])
    st_folium(m, height=800, use_container_width=True, key="survey_map")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📄 3. สร้างรายงาน PowerPoint")
    cap = st.file_uploader("อัปโหลดรูป Capture แผนที่", type=['jpg','png'])
    if cap and st.session_state.get('export_data'):
        if st.button("🚀 ดาวน์โหลดรายงาน PPTX"):
            pptx = create_summary_pptx(cap.getvalue(), st.session_state.export_data)
            st.download_button("📥 คลิกเพื่อดาวน์โหลด", data=pptx, file_name="Cable_Report.pptx")
