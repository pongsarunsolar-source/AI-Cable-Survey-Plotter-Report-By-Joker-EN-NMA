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
# แก้ไข: ใช้ SDK ตัวใหม่ตามคำแนะนำใน Log
from google import genai
from google.genai import types
import zipfile
from lxml import etree

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า Gemini API (SDK ใหม่) ---
# แนะนำให้ใช้ st.secrets["GEMINI_API_KEY"] เพื่อความปลอดภัยในอนาคต
client = genai.Client(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")

@st.cache_resource
def load_ocr():
    # กำหนดที่เก็บ Model ไว้ใน Working Directory เพื่อไม่ให้ App ค้างตอน Download
    model_storage = os.path.join(os.getcwd(), "ocr_models")
    if not os.path.exists(model_storage):
        os.makedirs(model_storage)
    return easyocr.Reader(['en'], gpu=False, model_storage_directory=model_storage)

# โหลดโมเดลเตรียมไว้
reader = load_ocr()

# --- 2. ฟังก์ชันช่วยดึงรูปภาพ (มี Timeout ป้องกัน Connection Reset) ---
@st.cache_data(ttl=3600)
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except: return None
    return None

# --- 3. วิเคราะห์สาเหตุด้วย AI (ใช้ SDK ใหม่) ---
@st.cache_data(show_spinner=False)
def analyze_cable_issue(img_bytes):
    try:
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=[
                """วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเพียง "หนึ่งเดียว" จาก 4 สาเหตุ:
                1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ
                ตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น""",
                types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")
            ]
        )
        return response.text.strip()
    except: return "วิเคราะห์ไม่ได้"

# --- 4. ฟังก์ชันจัดการพิกัด (EXIF & OCR) ---
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
def get_lat_lon_ocr(img_bytes):
    try:
        img_np = np.array(Image.open(BytesIO(img_bytes)))
        results = reader.readtext(img_np)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 5. UI Helpers ---
def img_to_custom_icon(img_obj, issue_text):
    img_resized = img_obj.copy()
    img_resized.thumbnail((150, 150))
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=60)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="width: 160px; background: white; padding: 5px; border-radius: 10px; border: 2px solid #FF8C42; box-shadow: 2px 2px 10px rgba(0,0,0,0.2);">
            <div style="font-size: 11px; font-weight: bold; color: #2D5A27; text-align: center; margin-bottom: 3px;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="width: 100%; border-radius: 5px;">
        </div>
    '''

# --- 6. สร้างรายงาน PPTX ---
def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    
    if map_image_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)

    for i, item in enumerate(image_list):
        if i % 8 == 0: slide = prs.slides.add_slide(prs.slide_layouts[6])
        idx = i % 8
        x = Inches(0.5 + (idx % 4) * 2.3)
        y = Inches(0.5 + (idx // 4) * 2.5)
        
        buf = BytesIO()
        item['img_obj'].save(buf, format="JPEG")
        slide.shapes.add_picture(BytesIO(buf.getvalue()), x, y, width=Inches(2.1), height=Inches(1.5))
        
        tf = slide.shapes.add_textbox(x, y + Inches(1.55), Inches(2.1), Inches(0.5)).text_frame
        p = tf.paragraphs[0]
        p.text = f"{item['issue']}\nLat: {item['lat']:.5f}, Lon: {item['lon']:.5f}"
        p.font.size = Pt(8)

    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# --- 7. Main App UI ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")

# Header & Logo
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; padding: 20px; background: white; border-radius: 15px; border-bottom: 4px solid #FF8C42;">
        <div><h1 style="margin:0; color: #2D5A27;">AI Cable Plotter</h1><small>By Joker EN-NMA</small></div>
        {f'<img src="data:image/png;base64,{joker_base64}" style="width:80px; height:80px; border-radius:50%; border:3px solid #FF8C42;">' if joker_base64 else ""}
    </div>
""", unsafe_allow_html=True)

# 1. KML/KMZ
st.subheader("🌐 1. ข้อมูลโครงข่าย")
kml_file = st.file_uploader("อัปโหลด KML/KMZ", type=['kml', 'kmz'])
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
            name = pm.findtext('.//kml:name', default="Marker", namespaces=ns)
            coords = pm.findtext('.//kml:coordinates', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords.strip().split()]
                kml_elements.append({'name': name, 'points': pts, 'is_point': len(pts) == 1})
    except: st.error("ไม่สามารถอ่านไฟล์ KML ได้")

# 2. Survey Images
st.subheader("📁 2. อัปโหลดรูปสำรวจ")
uploaded_files = st.file_uploader("เลือกไฟล์รูปภาพ (JPG/PNG)", type=['jpg','jpeg','png'], accept_multiple_files=True)

if uploaded_files or kml_elements:
    m = folium.Map(location=[13.75, 100.5], zoom_start=14, tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
    all_bounds = []

    if uploaded_files:
        if 'export_data' not in st.session_state: st.session_state.export_data = []
        
        # ตรวจสอบการอัปโหลดใหม่
        current_hash = hash(tuple([f.name for f in uploaded_files]))
        if st.session_state.get('last_hash') != current_hash:
            st.session_state.export_data = []
            st.session_state.last_hash = current_hash
            
            with st.status("AI กำลังประมวลผล...") as status:
                for f in uploaded_files:
                    f_bytes = f.getvalue()
                    img = ImageOps.exif_transpose(Image.open(BytesIO(f_bytes)))
                    lat, lon = get_lat_lon_exif(img)
                    if lat is None: lat, lon = get_lat_lon_ocr(f_bytes)
                    
                    if lat:
                        issue = analyze_cable_issue(f_bytes)
                        st.session_state.export_data.append({'img_obj': img, 'issue': issue, 'lat': lat, 'lon': lon})
                status.update(label="เสร็จเรียบร้อย!", state="complete")

    # วาดข้อมูลลงแผนที่
    for elem in kml_elements:
        if elem['is_point']:
            folium.Marker(elem['points'][0], tooltip=elem['name']).add_to(m)
            all_bounds.append(elem['points'][0])
        else:
            folium.PolyLine(elem['points'], color="red", weight=3).add_to(m)
            all_bounds.extend(elem['points'])

    for data in st.session_state.get('export_data', []):
        icon_html = img_to_custom_icon(data['img_obj'], data['issue'])
        folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=icon_html)).add_to(m)
        all_bounds.append([data['lat'], data['lon']])

    if all_bounds: m.fit_bounds(all_bounds)
    st_folium(m, height=700, use_container_width=True)

# 3. Export
st.subheader("📄 3. ออกรายงาน")
col1, col2 = st.columns(2)
with col1: map_cap = st.file_uploader("อัปโหลดภาพแผนที่ (Capture)", type=['jpg','png'])
with col2:
    if map_cap and st.session_state.get('export_data'):
        if st.button("ดาวน์โหลดรายงาน PowerPoint"):
            pptx = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
            st.download_button("📥 คลิกเพื่อโหลดไฟล์", data=pptx, file_name="Report.pptx")
