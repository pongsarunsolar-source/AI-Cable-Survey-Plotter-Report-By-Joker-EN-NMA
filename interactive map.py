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

# แก้ไขปัญหา SSL สำหรับการดาวน์โหลดโมเดล OCR บน Cloud
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. การตั้งค่าระบบและการจัดการความปลอดภัย ---
# แนะนำให้ตั้งค่าใน Streamlit Secrets: GEMINI_API_KEY
API_KEY = st.secrets.get("GEMINI_API_KEY", "AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
genai.configure(api_key=API_KEY)
model_ai = genai.GenerativeModel('gemini-1.5-flash')

# โหลด OCR ครั้งเดียวและเก็บไว้ใน Cache (ปิด GPU เพราะ Cloud เป็น CPU)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# --- 2. ฟังก์ชันประมวลผล (Optimized for CPU/RAM) ---

@st.cache_data(show_spinner="วิเคราะห์ภาพด้วย AI...")
def analyze_image_cached(img_bytes):
    """ส่งรูปไปวิเคราะห์ที่ Gemini (ประหยัด CPU ฝั่งเรา)"""
    try:
        img = Image.open(BytesIO(img_bytes))
        prompt = """วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเพียง "หนึ่งเดียว" จาก 4 สาเหตุ:
        1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span เสาไฟฟ้า | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ
        ตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น"""
        response = model_ai.generate_content([prompt, img])
        return response.text.strip()
    except:
        return "วิเคราะห์ไม่ได้"

def get_lat_lon_exif(image):
    """ดึงพิกัดจาก EXIF Data (เร็วและประหยัดทรัพยากรที่สุด)"""
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

@st.cache_data(show_spinner="กำลังอ่านพิกัดจากภาพ (OCR)...")
def get_lat_lon_ocr_cached(img_bytes):
    """กรณีไม่มี EXIF ให้ใช้ OCR อ่านข้อความบนภาพ"""
    try:
        img_np = np.array(Image.open(BytesIO(img_bytes)))
        results = reader.readtext(img_np)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 3. ส่วน UI & Map Visualization ---

def img_to_icon(img, issue):
    """สร้าง Custom Icon สำหรับหมุดบนแผนที่"""
    thumb = img.copy()
    thumb.thumbnail((120, 120))
    buf = BytesIO()
    thumb.save(buf, format="JPEG", quality=50) # ลดคุณภาพลงเพื่อความเร็ว
    img_b64 = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="width: 130px; background: white; padding: 5px; border-radius: 8px; border: 2px solid #FF8C42; box-shadow: 2px 2px 10px rgba(0,0,0,0.2);">
            <div style="font-size: 10px; font-weight: bold; color: #2D5A27; text-align: center; margin-bottom: 3px;">{issue}</div>
            <img src="data:image/jpeg;base64,{img_b64}" style="width: 100%; border-radius: 4px;">
        </div>
    '''

def create_pptx(map_bytes, data_list):
    """สร้างรายงาน PPTX"""
    prs = Presentation()
    if map_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
        slide.shapes.add_picture(BytesIO(map_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    
    # เพิ่มหน้าสรุปรูปถ่าย (หน้าละ 4 รูป เพื่อความสวยงาม)
    for i in range(0, len(data_list), 4):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for j, item in enumerate(data_list[i:i+4]):
            x_pos = Inches(0.5 + (j * 2.3))
            y_pos = Inches(1.0)
            img_buf = BytesIO()
            item['img'].save(img_buf, format="JPEG")
            img_buf.seek(0)
            slide.shapes.add_picture(img_buf, x_pos, y_pos, width=Inches(2), height=Inches(1.5))
            tb = slide.shapes.add_textbox(x_pos, y_pos + Inches(1.6), Inches(2), Inches(0.5))
            tb.text = f"{item['issue']}\nLat: {item['lat']:.4f}"
            
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# --- Main App Interface ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.title("🔌 AI Cable Plotter (Cloud Optimized)")
st.info("ระบบวิเคราะห์และพล็อตจุดสายเคเบิลด้วย AI | รองรับไฟล์ KML/KMZ และการอ่านพิกัดจากรูปภาพ")

col_input, col_map = st.columns([1, 3])

with col_input:
    st.subheader("1. อัปโหลดข้อมูล")
    kml_file = st.file_uploader("ไฟล์โครงข่าย (KML/KMZ)", type=['kml', 'kmz'])
    img_files = st.file_uploader("รูปถ่ายสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)
    
    map_cap = st.file_uploader("📸 อัปโหลดรูปภาพแผนที่ (เพื่อทำรายงาน)", type=['jpg','png'])
    if st.button("生成 PPTX Report"):
        if 'survey_results' in st.session_state and map_cap:
            pptx_data = create_pptx(map_cap.getvalue(), st.session_state.survey_results)
            st.download_button("📩 Download Report", pptx_data, "Cable_Report.pptx")
        else:
            st.warning("กรุณาอัปโหลดรูปและ Capture แผนที่ก่อน")

# ประมวลผลและแสดงผลแผนที่
with col_map:
    m = folium.Map(location=[13.75, 100.5], zoom_start=6)
    all_points = []
    st.session_state.survey_results = []

    # การจัดการ KML
    if kml_file:
        try:
            content = kml_file.getvalue()
            if kml_file.name.endswith('.kmz'):
                with zipfile.ZipFile(BytesIO(content)) as z:
                    content = z.read([n for n in z.namelist() if n.endswith('.kml')][0])
            
            root = etree.fromstring(content)
            ns = {'kml': 'http://www.opengis.net/kml/2.2'}
            for pm in root.xpath('.//kml:Placemark', namespaces=ns):
                coords = pm.findtext('.//kml:coordinates', namespaces=ns)
                if coords:
                    p = [float(c.split(',')[1]) for c in coords.strip().split()]
                    l = [float(c.split(',')[0]) for c in coords.strip().split()]
                    pts = list(zip(p, l))
                    folium.PolyLine(pts, color="red", weight=2).add_to(m)
                    all_points.extend(pts)
        except: st.error("ไม่สามารถอ่านไฟล์ KML ได้")

    # การจัดการรูปภาพ
    if img_files:
        for f in img_files[:20]: # จำกัด 20 รูปเพื่อป้องกัน RAM ล่ม
            img_raw = Image.open(f)
            img_fixed = ImageOps.exif_transpose(img_raw)
            lat, lon = get_lat_lon_exif(img_raw)
            
            # เตรียมไฟล์สำหรับการวิเคราะห์ (ลดขนาดเพื่อประหยัด Data)
            buf = BytesIO()
            img_fixed.save(buf, format="JPEG", quality=70)
            img_bytes = buf.getvalue()

            if lat is None:
                lat, lon = get_lat_lon_ocr_cached(img_bytes)

            if lat:
                issue = analyze_image_cached(img_bytes)
                icon_html = img_to_icon(img_fixed, issue)
                folium.Marker([lat, lon], icon=folium.DivIcon(html=icon_html)).add_to(m)
                all_points.append([lat, lon])
                st.session_state.survey_results.append({'img': img_fixed, 'issue': issue, 'lat': lat, 'lon': lon})

    if all_points:
        m.fit_bounds(all_points)
    
    st_folium(m, width="100%", height=700)
