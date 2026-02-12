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

# แก้ไขปัญหา SSL สำหรับการดาวน์โหลดโมเดล OCR
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า Google Gemini API ---
genai.configure(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")
model_ai = genai.GenerativeModel('gemini-1.5-flash')

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'])

reader = load_ocr()

# --- 2. ฟังก์ชันช่วยดึงรูปภาพจาก Google Drive มาเป็น Base64 ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except Exception:
        return None
    return None

# --- 3. ฟังก์ชันวิเคราะห์สาเหตุ ---
def analyze_cable_issue(image):
    try:
        prompt = """
        วิเคราะห์รูปภาพสายเคเบิลนี้และเลือกตอบเพียง "หนึ่งเดียว" จาก 4 สาเหตุต่อไปนี้เท่านั้น:
        1. cable ตกพื้น
        2. หัวต่ออยู่กลาง span เสาไฟฟ้า
        3. ไฟไหม้ cable
        4. หัวต่อขวดน้ำ
        
        ตอบเฉพาะชื่อสาเหตุเป็นภาษาไทยเท่านั้น หากไม่ตรงกับข้อใดเลยให้ตอบว่า "ตรวจสอบไม่พบสาเหตุ"
        """
        response = model_ai.generate_content([prompt, image])
        return response.text.strip()
    except Exception:
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

def get_lat_lon_ocr(image):
    try:
        img_np = np.array(image)
        results = reader.readtext(img_np)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 5. ฟังก์ชันสร้าง Icon บนแผนที่ ---
def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((150, 150)) 
    buf = BytesIO()
    img_resized.save(buf, format="PNG")
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: fit-content; background-color: white; padding: 5px; border-radius: 12px; box-shadow: 0px 8px 24px rgba(0,0,0,0.12); border: 2px solid #FF8C42; transform: translate(-50%, -100%); margin-top: -10px;">
            <div style="font-size: 11px; font-weight: 700; color: #2D5A27; margin-bottom: 4px; text-align: center; font-family: 'Inter', sans-serif;">{issue_text}</div>
            <img src="data:image/png;base64,{img_str}" style="max-width: 140px; display: block; border-radius: 4px;">
            <div style="position: absolute; bottom: -10px; left: 50%; transform: translateX(-50%); width: 0; height: 0; border-left: 10px solid transparent; border-right: 10px solid transparent; border-top: 10px solid #FF8C42;"></div>
        </div>
    '''

# --- 6. ฟังก์ชัน Export PowerPoint ---
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
        margin_x = (prs.slide_width - (img_w * cols)) / (cols + 1)
        margin_y = (prs.slide_height - (img_h * rows + Inches(1.0))) / (rows + 1)

        for i, item in enumerate(image_list[:8]):
            curr_row, curr_col = i // cols, i % cols
            x = margin_x + (curr_col * (img_w + margin_x))
            y = margin_y + (curr_row * (img_h + margin_y + Inches(0.5)))
            
            image = item['img_obj'].copy()
            target_ratio = img_w / img_h
            w_px, h_px = image.size
            if (w_px/h_px) > target_ratio:
                new_w = h_px * target_ratio
                left = (w_px - new_w) / 2
                image = image.crop((left, 0, left + new_w, h_px))
            else:
                new_h = w_px / target_ratio
                top = (h_px - new_h) / 2
                image = image.crop((0, top, w_px, top + new_h))
            
            buf = BytesIO()
            image.save(buf, format="JPEG")
            buf.seek(0)
            slide2.shapes.add_picture(buf, x, y, width=img_w, height=img_h)
            
            txt_box = slide2.shapes.add_textbox(x, y + img_h + Inches(0.05), img_w, Inches(0.6))
            tf = txt_box.text_frame
            tf.word_wrap = True
            p1 = tf.paragraphs[0]
            p1.text = f"สาเหตุ: {item['issue']}"
            p1.font.size = Pt(8)
            p1.font.bold = True
            p2 = tf.add_paragraph()
            p2.text = f"Lat: {item['lat']:.5f}\nLong: {item['lon']:.5f}"
            p2.font.size = Pt(7)

    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# --- 7. Streamlit UI (ธีม Minimal ส้ม+เขียว) ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    
    .stApp {
        background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%);
    }
    
    html, body, [class*="css"], .stMarkdown, p, h1, h2, h3 {
        font-family: 'Inter', sans-serif;
        color: #2D3748;
    }
    
    .header-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 20px;
        padding: 25px 35px;
        background: white;
        border-radius: 24px;
        box-shadow: 0 12px 40px rgba(0,0,0,0.06);
        border-bottom: 5px solid #FF8C42;
        margin-bottom: 30px;
    }
    
    .main-title {
        background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        font-size: 2.6rem;
        margin: 0;
    }
    
    .joker-icon {
        width: 100px;
        height: 100px;
        object-fit: cover;
        border-radius: 50%;
        border: 4px solid #FFFFFF;
        outline: 3px solid #FF8C42;
        box-shadow: 0 6px 20px rgba(255, 140, 66, 0.4);
    }
    
    .stButton>button {
        background: #2D5A27;
        color: white;
        border-radius: 14px;
        border: none;
        padding: 12px 35px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(45, 90, 39, 0.2);
    }
    
    .stButton>button:hover {
        background: #FF8C42;
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(255, 140, 66, 0.3);
        color: white;
    }
    
    .stFileUploader section {
        background-color: white !important;
        border: 2px dashed #CBD5E0 !important;
        border-radius: 20px !important;
    }

    hr {
        border: 0;
        height: 1px;
        background: linear-gradient(to right, rgba(0,0,0,0), rgba(255,140,66,0.4), rgba(0,0,0,0));
        margin: 45px 0;
    }
    </style>
    """, unsafe_allow_html=True)

# ส่วนการดึงรูปภาพ Joker
joker_file_id = "1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr"
joker_base64 = get_image_base64_from_drive(joker_file_id)

if joker_base64:
    header_html = f"""
        <div class="header-container">
            <div>
                <h1 class="main-title">AI Cable Plotter</h1>
                <p style="margin:0; color: #718096; font-size: 1.1rem; font-weight: 600;">Report  By Joker EN-NMA</p>
            </div>
            <img src="data:image/png;base64,{joker_base64}" class="joker-icon">
        </div>
    """
else:
    header_html = f"""
        <div class="header-container">
            <div>
                <h1 class="main-title">AI Cable Plotter</h1>
                <p style="margin:0; color: #718096; font-size: 1.1rem; font-weight: 600;">Report System By Joker EN-NMA</p>
            </div>
            <div class="joker-icon" style="display:flex; align-items:center; justify-content:center; background:#eee; font-size:10px;">No Image</div>
        </div>
    """

st.markdown(header_html, unsafe_allow_html=True)

# --- ส่วนการทำงานหลัก ---
uploaded_files = st.file_uploader("📁 1. อัปโหลดรูปภาพสำรวจ (JPG/PNG)", type=['jpg','jpeg','png'], accept_multiple_files=True)

if 'export_data' not in st.session_state:
    st.session_state.export_data = []
if 'last_files_hash' not in st.session_state:
    st.session_state.last_files_hash = ""

if uploaded_files:
    m = folium.Map(location=[13.75, 100.5], zoom_start=6, tiles="cartodbpositron")
    points_to_fit = []
    
    current_hash = "".join([f.name + str(f.size) for f in uploaded_files])
    if st.session_state.last_files_hash != current_hash:
        st.session_state.export_data = []
        st.session_state.last_files_hash = current_hash

    for i, f in enumerate(uploaded_files):
        if i >= len(st.session_state.export_data):
            raw_data = f.getvalue()
            raw_img = Image.open(BytesIO(raw_data))
            img_straight = ImageOps.exif_transpose(raw_img)
            lat, lon = get_lat_lon_exif(raw_img)
            if lat is None:
                with st.spinner(f"OCR Scanning: {f.name}..."):
                    lat, lon = get_lat_lon_ocr(img_straight)
            if lat is not None:
                with st.spinner(f"AI Analyzing: {f.name}..."):
                    issue = analyze_cable_issue(img_straight)
                st.session_state.export_data.append({
                    'content': raw_data, 'issue': issue, 'lat': lat, 'lon': lon, 'img_obj': img_straight
                })
        
        if i < len(st.session_state.export_data):
            data = st.session_state.export_data[i]
            icon_html = img_to_custom_icon(data['img_obj'], data['issue'])
            folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=icon_html, icon_size=(150, 200))).add_to(m)
            points_to_fit.append([data['lat'], data['lon']])

    if points_to_fit:
        m.fit_bounds(points_to_fit)
        st_folium(m, width="100%", height=700, key="map_view")
        
        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("📄 2. สร้างรายงาน PowerPoint")
        
        col1, col2 = st.columns([1, 1])
        with col1:
            map_cap = st.file_uploader("อัปโหลดรูป Capture แผนที่ด้านบน", type=['jpg','png'])
        
        if map_cap:
            with col2:
                st.write("") 
                if st.button("🚀 สรุปรายงานและดาวน์โหลดไฟล์ PPTX"):
                    pptx_data = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
                    st.download_button("📥 คลิกเพื่อดาวน์โหลดรายงาน", data=pptx_data, file_name="Cable_AI_Report.pptx")