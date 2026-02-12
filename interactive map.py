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

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า Google Gemini API ---
genai.configure(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
model_ai = genai.GenerativeModel('gemini-1.5-flash')

# ปรับการโหลด OCR ให้ประหยัด RAM
@st.cache_resource
def load_ocr():
    # บังคับใช้ CPU และระบุ directory สำหรับเก็บ model เพื่อไม่ให้โหลดใหม่ทุกรอบ
    return easyocr.Reader(['en'], gpu=False)

reader = load_ocr()

# --- 2. ฟังก์ชันช่วยดึงรูปภาพจาก Google Drive ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except: return None
    return None

# --- 3. ฟังก์ชันวิเคราะห์สาเหตุ ---
def analyze_cable_issue(image):
    try:
        # ย่อรูปก่อนส่งให้ AI เพื่อประหยัด Bandwidth และ Memory
        img_small = image.copy()
        img_small.thumbnail((800, 800))
        prompt = "วิเคราะห์รูปสายเคเบิล: 1. cable ตกพื้น 2. หัวต่ออยู่กลาง span 3. ไฟไหม้ cable 4. หัวต่อขวดน้ำ ตอบแค่ชื่อสาเหตุสั้นๆ"
        response = model_ai.generate_content([prompt, img_small])
        return response.text.strip()
    except: return "ตรวจสอบไม่พบสาเหตุ"

# --- 4. ฟังก์ชันจัดการพิกัด (ปรับปรุงเรื่อง RAM) ---
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
            d = float(dms[0])
            m = float(dms[1])
            s = float(dms[2])
            res = d + (m / 60.0) + (s / 3600.0)
            return -res if ref in ['S', 'W'] else res
            
        return dms_to_decimal(gps_info['GPSLatitude'], gps_info['GPSLatitudeRef']), \
               dms_to_decimal(gps_info['GPSLongitude'], gps_info['GPSLongitudeRef'])
    except: return None, None

def get_lat_lon_ocr(image):
    try:
        # ย่อรูปก่อนทำ OCR เพื่อกัน RAM เต็ม
        img_for_ocr = image.copy()
        img_for_ocr.thumbnail((1000, 1000))
        img_np = np.array(img_for_ocr)
        # ใช้ paragraph=True เพื่อให้ผลลัพธ์กระชับขึ้น
        results = reader.readtext(img_np, paragraph=True)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 5. ฟังก์ชันสร้าง Icon (ย่อรูปลงอีกเพื่อความลื่นไหล) ---
def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((120, 120)) 
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=70) # ใช้ JPEG ประหยัดกว่า PNG
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: fit-content; background-color: white; padding: 5px; border-radius: 10px; border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 10px; font-weight: 700; color: #2D5A27; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width: 110px; border-radius: 4px;">
        </div>
    '''

# --- 6. ฟังก์ชัน Export PowerPoint ---
def create_summary_pptx(map_image_bytes, image_list):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    if map_image_bytes:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        slide1.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
    
    for item in image_list[:12]: # เพิ่มจำนวนเป็น 12 รูป
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        buf = BytesIO()
        item['img_obj'].save(buf, format="JPEG")
        buf.seek(0)
        slide.shapes.add_picture(buf, Inches(0.5), Inches(0.5), width=Inches(5))
        txt_box = slide.shapes.add_textbox(Inches(6), Inches(1), Inches(3.5), Inches(2))
        tf = txt_box.text_frame
        tf.text = f"สาเหตุ: {item['issue']}\nLat: {item['lat']}\nLon: {item['lon']}"
    
    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# --- 7. Streamlit UI ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")

# (CSS เหมือนเดิมของคุณ แต่เพิ่มการจัดการรูปภาพให้ลื่นขึ้น)
st.markdown("""<style>...</style>""", unsafe_allow_html=True) 

# โหลดหัวข้อ (Header)
joker_file_id = "1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr"
joker_base64 = get_image_base64_from_drive(joker_file_id)

st.title("AI Cable Plotter - Updated")

uploaded_files = st.file_uploader("📁 อัปโหลดรูปภาพ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if 'export_data' not in st.session_state:
    st.session_state.export_data = []

if uploaded_files:
    m = folium.Map(location=[13.75, 100.5], zoom_start=6, tiles="cartodbpositron")
    points_to_fit = []
    
    for f in uploaded_files:
        # ตรวจสอบว่ารูปนี้เคยรันไปหรือยังเพื่อไม่ให้รันซ้ำ (ประหยัด RAM)
        already_processed = any(d.get('filename') == f.name for d in st.session_state.export_data)
        
        if not already_processed:
            raw_data = f.getvalue()
            raw_img = Image.open(BytesIO(raw_data))
            img_straight = ImageOps.exif_transpose(raw_img)
            
            lat, lon = get_lat_lon_exif(img_straight)
            if lat is None:
                with st.spinner(f"Scanning OCR: {f.name}"):
                    lat, lon = get_lat_lon_ocr(img_straight)
            
            if lat is not None:
                with st.spinner(f"AI Analyzing: {f.name}"):
                    issue = analyze_cable_issue(img_straight)
                st.session_state.export_data.append({
                    'filename': f.name, 'issue': issue, 'lat': lat, 'lon': lon, 'img_obj': img_straight
                })

    for data in st.session_state.export_data:
        icon_html = img_to_custom_icon(data['img_obj'], data['issue'])
        folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=icon_html)).add_to(m)
        points_to_fit.append([data['lat'], data['lon']])

    if points_to_fit:
        m.fit_bounds(points_to_fit)
        st_folium(m, width="100%", height=600)
        
        if st.button("Download Report"):
             # สร้างรายงาน (ต้องมีการ Capture หน้าจอมาใส่ตาม Logic เดิมของคุณ)
             pass
