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

# แก้ไขปัญหา SSL สำหรับการดึงโมเดล OCR
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า Google Gemini API ---
# แนะนำให้ใช้ st.secrets เพื่อความปลอดภัย แต่ยังคงโครงสร้างเดิมตามคำขอ
genai.configure(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
model_ai = genai.GenerativeModel('gemini-1.5-flash')

@st.cache_resource
def load_ocr():
    # ปรับใช้ CPU mode เพื่อประหยัด RAM บน Streamlit Cloud
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# --- 2. ฟังก์ชันช่วยดึงรูปภาพ (ปรับปรุง Timeout) ---
@st.cache_data(ttl=3600)
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except: 
        return None
    return None

# --- 3. ฟังก์ชันวิเคราะห์ AI (เพิ่ม Cache เพื่อประหยัด API Quota/Speed) ---
@st.cache_data(show_spinner=False)
def analyze_cable_issue_cached(img_bytes):
    try:
        image = Image.open(BytesIO(img_bytes))
        image.thumbnail((500, 500)) # ลดขนาดก่อนส่งไป AI เพื่อความเร็ว
        prompt = """วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเพียง "หนึ่งเดียว" จาก 4 สาเหตุ:
        1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ
        ตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น"""
        response = model_ai.generate_content([prompt, image])
        return response.text.strip()
    except: 
        return "วิเคราะห์ไม่ได้"

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

@st.cache_data(show_spinner=False)
def get_lat_lon_ocr_cached(img_bytes):
    try:
        image = Image.open(BytesIO(img_bytes))
        img_np = np.array(image)
        results = reader.readtext(img_np)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 5. UI Helpers ---
def create_div_label(name):
    return f'''<div style="font-size: 11px; font-weight: 800; color: #D9534F; white-space: nowrap; transform: translate(-50%, -150%); text-shadow: 2px 2px 4px white;">{name}</div>'''

def img_to_custom_icon(img_obj, issue_text):
    buf = BytesIO()
    img_obj.save(buf, format="JPEG", quality=50) # ลด Quality เพื่อให้แผนที่ลื่นขึ้น
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: 150px; background: white; padding: 5px; border-radius: 8px; border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 10px; font-weight: bold; text-align: center; color: #2D5A27;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="width: 100%; border-radius: 4px;">
        </div>
    '''

# --- 6. Export PowerPoint ---
def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    
    if map_image_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)

    for i in range(0, len(image_list), 8):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        grid = image_list[i:i+8]
        for idx, item in enumerate(grid):
            x = Inches(0.5 + (idx % 4) * 2.3)
            y = Inches(0.5 + (idx // 4) * 2.5)
            
            buf = BytesIO()
            item['img_obj'].save(buf, format="JPEG")
            buf.seek(0)
            slide.shapes.add_picture(buf, x, y, width=Inches(2.1), height=Inches(1.5))
            
            tb = slide.shapes.add_textbox(x, y + Inches(1.5), Inches(2.1), Inches(0.5))
            tb.text_frame.text = f"{item['issue']}\n{item['lat']:.5f}, {item['lon']:.5f}"
            for p in tb.text_frame.paragraphs:
                p.font.size = Pt(8)
    
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# --- 7. Main UI ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")

# CSS 
st.markdown("""<style>
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 20px; background: white; border-radius: 15px; border-bottom: 5px solid #FF8C42; margin-bottom: 20px; }
    .joker-icon { width: 80px; height: 80px; border-radius: 50%; border: 3px solid #FF8C42; }
</style>""", unsafe_allow_html=True)

# Header
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f'''<div class="header-container"><div><h1 style="margin:0;">AI Cable Plotter</h1><p style="margin:0; color: gray;">By Joker EN-NMA</p></div>
{"<img src='data:image/png;base64,"+joker_base64+"' class='joker-icon'>" if joker_base64 else ""}</div>''', unsafe_allow_html=True)

# 1. KML Section
st.subheader("🌐 1. ข้อมูลโครงข่าย (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลดไฟล์ KML/KMZ", type=['kml', 'kmz'])
kml_elements = []

if kml_file:
    with st.spinner("กำลังอ่านไฟล์ KML..."):
        try:
            content = kml_file.getvalue()
            if kml_file.name.endswith('.kmz'):
                with zipfile.ZipFile(BytesIO(content)) as z:
                    content = z.read([n for n in z.namelist() if n.endswith('.kml')][0])
            root = etree.fromstring(content)
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            for pm in root.xpath('.//kml:Placemark', namespaces=ns):
                name = pm.findtext('.//kml:name', default="N/A", namespaces=ns)
                coords = pm.findtext('.//kml:coordinates', namespaces=ns)
                if coords:
                    pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords.strip().split()]
                    kml_elements.append({'name': name, 'points': pts, 'is_point': len(pts) == 1})
        except Exception as e: st.error(f"KML Error: {e}")

# 2. Upload & Map Section
st.subheader("📁 2. รูปภาพสำรวจ")
uploaded_files = st.file_uploader("อัปโหลดรูปภาพ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if uploaded_files or kml_elements:
    m = folium.Map(location=[13.75, 100.5], zoom_start=15, tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
    all_bounds = []

    # Process Images
    if uploaded_files:
        current_hash = hash(tuple([f.name for f in uploaded_files]))
        if 'last_hash' not in st.session_state or st.session_state.last_hash != current_hash:
            st.session_state.export_data = []
            st.session_state.last_hash = current_hash
            
            with st.status("กำลังประมวลผลรูปภาพด้วย AI...") as status:
                for f in uploaded_files:
                    f_bytes = f.getvalue()
                    img = Image.open(BytesIO(f_bytes))
                    img = ImageOps.exif_transpose(img)
                    lat, lon = get_lat_lon_exif(img)
                    if lat is None: lat, lon = get_lat_lon_ocr_cached(f_bytes)
                    
                    if lat:
                        issue = analyze_cable_issue_cached(f_bytes)
                        st.session_state.export_data.append({'img_obj': img, 'issue': issue, 'lat': lat, 'lon': lon})
                status.update(label="ประมวลผลเสร็จสิ้น!", state="complete")

    # Plot KML
    for elem in kml_elements:
        if elem['is_point']:
            folium.Marker(elem['points'][0], tooltip=elem['name']).add_to(m)
            all_bounds.append(elem['points'][0])
        else:
            folium.PolyLine(elem['points'], color="orange", weight=4).add_to(m)
            all_bounds.extend(elem['points'])

    # Plot Image Markers
    for data in st.session_state.get('export_data', []):
        icon_html = img_to_custom_icon(data['img_obj'], data['issue'])
        folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=icon_html)).add_to(m)
        all_bounds.append([data['lat'], data['lon']])

    if all_bounds: m.fit_bounds(all_bounds)
    st_folium(m, height=700, use_container_width=True)

# 3. Export
st.subheader("📄 3. สรุปรายงาน")
col1, col2 = st.columns(2)
with col1:
    map_cap = st.file_uploader("1. อัปโหลดรูป Capture แผนที่", type=['jpg','png'])
with col2:
    if map_cap and st.session_state.get('export_data'):
        if st.button("2. ดาวน์โหลดรายงาน (PPTX)", use_container_width=True):
            pptx_file = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
            st.download_button("📥 Click to Download", data=pptx_file, file_name="Survey_Report.pptx")
