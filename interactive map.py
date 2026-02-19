import ssl
import os
import streamlit as st
import folium
from streamlit_folium import st_folium
from folium.plugins import MeasureControl
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
# ใช้ SDK ใหม่ตาม Log
from google import genai
from google.genai import types
import zipfile
from lxml import etree
import math

# แก้ไขปัญหา SSL
ssl._create_default_https_context = ssl._create_unverified_context

# --- 1. ตั้งค่า Google Gemini API ---
# แนะนำ: เพื่อความปลอดภัย ควรย้าย API KEY ไปไว้ใน st.secrets
client = genai.Client(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")

@st.cache_resource
def load_ocr():
    model_path = os.path.join(os.getcwd(), "easyocr_models")
    if not os.path.exists(model_path):
        os.makedirs(model_path)
    return easyocr.Reader(['en'], gpu=False, model_storage_directory=model_path)

# --- 2. ฟังก์ชันช่วยดึงรูปภาพ Joker ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except Exception: return None
    return None

# --- 3. ฟังก์ชันวิเคราะห์สาเหตุด้วย AI ---
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

# --- ฟังก์ชันหาจุดที่ไกลกันที่สุด (Head - Tail) ---
def get_farthest_points(coordinates):
    """
    รับ List ของพิกัด [[lat, lon], ...] เฉพาะจาก KML
    คืนค่า (Start_Point, End_Point)
    """
    if not coordinates or len(coordinates) < 2:
        return None, None
    
    max_dist = -1
    p1_best, p2_best = None, None
    
    # วนลูปหาคู่ที่ไกลที่สุด (Brute Force)
    for i in range(len(coordinates)):
        for j in range(i + 1, len(coordinates)):
            lat1, lon1 = coordinates[i]
            lat2, lon2 = coordinates[j]
            dist = (lat1 - lat2)**2 + (lon1 - lon2)**2
            
            if dist > max_dist:
                max_dist = dist
                p1_best = coordinates[i]
                p2_best = coordinates[j]
                
    return p1_best, p2_best

# --- ฟังก์ชันคำนวณเส้นทางเดิน (Walking) จาก OSRM ---
def get_osrm_route_head_tail(start_coord, end_coord):
    """
    ดึงข้อมูลเส้นทางจาก OSRM (Walking Profile) ระหว่างจุด 2 จุด
    """
    if not start_coord or not end_coord:
        return None, 0

    # OSRM รับค่าเป็น lon,lat
    coords_str = f"{start_coord[1]},{start_coord[0]};{end_coord[1]},{end_coord[0]}"
    
    url = f"http://router.project-osrm.org/route/v1/walking/{coords_str}?overview=full&geometries=geojson"
    
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if "routes" in data and len(data["routes"]) > 0:
                route = data["routes"][0]
                geometry = route["geometry"]["coordinates"] # [[lon, lat], ...]
                distance = route["distance"] # เมตร
                
                # แปลงกลับเป็น [lat, lon] สำหรับ Folium
                folium_coords = [[lat, lon] for lon, lat in geometry]
                return folium_coords, distance
    except Exception as e:
        print(f"OSRM Error: {e}")
    
    return None, 0

# --- 5. ฟังก์ชันสร้าง Label ชื่อสถานที่ ---
def create_div_label(name):
    return f'''
        <div style="
            font-size: 11px; font-weight: 800; color: #D9534F; white-space: nowrap;
            transform: translate(-50%, -150%); background-color: transparent;
            border: none; box-shadow: none;
            text-shadow: 2px 2px 4px white, -2px -2px 4px white, 2px -2px 4px white, -2px 2px 4px white;
            font-family: 'Inter', sans-serif;
        ">
            {name}
        </div>
    '''

# --- 6. ฟังก์ชันสร้าง Icon สำหรับรูปถ่ายบนแผนที่ ---
def img_to_custom_icon(img, issue_text):
    img_resized = img.copy()
    img_resized.thumbnail((150, 150)) 
    buf = BytesIO()
    img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: fit-content; background-color: white; padding: 5px; border-radius: 12px; box-shadow: 0px 8px 24px rgba(0,0,0,0.12); border: 2px solid #FF8C42; transform: translate(-50%, -100%); margin-top: -10px;">
            <div style="font-size: 11px; font-weight: 700; color: #2D5A27; margin-bottom: 4px; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width: 140px; display: block; border-radius: 4px;">
            <div style="position: absolute; bottom: -10px; left: 50%; transform: translateX(-50%); width: 0; height: 0; border-left: 10px solid transparent; border-right: 10px solid transparent; border-top: 10px solid #FF8C42;"></div>
        </div>
    '''

# --- 7. ฟังก์ชัน Export PowerPoint ---
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

# --- 8. UI Layout ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
    .joker-icon { width: 100px; height: 100px; object-fit: cover; border-radius: 50%; border: 4px solid #FFFFFF; outline: 3px solid #FF8C42; }
    .stButton>button { background: #2D5A27; color: white; border-radius: 14px; padding: 12px 35px; font-weight: 600; }
    .stButton>button:hover { background: #FF8C42; color: white; }
</style>""", unsafe_allow_html=True)

# Header
joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
header_html = f'''<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA</p></div>
{"<img src='data:image/png;base64,"+joker_base64+"' class='joker-icon'>" if joker_base64 else ""}</div>'''
st.markdown(header_html, unsafe_allow_html=True)

# --- 9. เมนู KML/KMZ ---
st.subheader("🌐 1. ข้อมูลโครงข่าย & จุดติดตั้ง (KML/KMZ)")
kml_file = st.file_uploader("อัปโหลดไฟล์ KML หรือ KMZ", type=['kml', 'kmz'])

kml_elements = []
kml_points_pool = [] # เก็บพิกัดเฉพาะจาก KML เพื่อคำนวณ Route

if kml_file:
    try:
        if kml_file.name.endswith('.kmz'):
            with zipfile.ZipFile(kml_file) as z:
                kml_filename = [n for n in z.namelist() if n.endswith('.kml')][0]
                content = z.read(kml_filename)
        else:
            content = kml_file.getvalue()
        root = etree.fromstring(content)
        ns = {'kml': 'http://www.opengis.net/kml/2.2', 'mwm': 'https://maps.me', 'earth': 'http://earth.google.com/kml/2.2'}
        placemarks = root.xpath('.//kml:Placemark | .//earth:Placemark', namespaces=ns)
        for pm in placemarks:
            name_node = pm.xpath('kml:name/text() | earth:name/text()', namespaces=ns)
            custom_name = pm.xpath('.//mwm:customName/mwm:lang[@code="default"]/text()', namespaces=ns)
            final_name = custom_name[0].strip() if custom_name else (name_node[0].strip() if name_node else "ไม่ระบุชื่อ")
            coords = pm.xpath('.//kml:coordinates/text() | .//earth:coordinates/text()', namespaces=ns)
            if coords:
                pts = [[float(c.split(',')[1]), float(c.split(',')[0])] for c in coords[0].strip().split()]
                kml_elements.append({'name': final_name, 'points': pts, 'is_point': len(pts) == 1})
                
                # เก็บพิกัดลง Pool ของ KML เท่านั้น
                for p in pts:
                    kml_points_pool.append(p)
                    
    except Exception as e: st.error(f"Error KML: {e}")

st.markdown("<hr>", unsafe_allow_html=True)

# --- 10. ส่วนการทำงานหลัก (Map & Export) ---
uploaded_files = st.file_uploader("📁 2. อัปโหลดรูปภาพสำรวจ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if 'export_data' not in st.session_state: st.session_state.export_data = []

if uploaded_files:
    current_hash = "".join([f.name + str(f.size) for f in uploaded_files])
    if 'last_hash' not in st.session_state or st.session_state.last_hash != current_hash:
        st.session_state.export_data = []
        st.session_state.last_hash = current_hash

    for i, f in enumerate(uploaded_files):
        if i >= len(st.session_state.export_data):
            raw_data = f.getvalue()
            raw_img = Image.open(BytesIO(raw_data))
            img_st = ImageOps.exif_transpose(raw_img)
            lat, lon = get_lat_lon_exif(raw_img)
            if lat is None: lat, lon = get_lat_lon_ocr(img_st)
            
            if lat:
                issue = analyze_cable_issue(raw_data) # ส่ง bytes ให้ SDK ใหม่
                st.session_state.export_data.append({'img_obj': img_st, 'issue': issue, 'lat': lat, 'lon': lon})

# --- คำนวณเส้นทาง (Routing Logic) เฉพาะจาก KML Points ---
route_coords = None
route_distance = 0

# 1. หาจุด Head - Tail (คู่ที่ไกลที่สุด) **จาก kml_points_pool เท่านั้น**
head_point, tail_point = get_farthest_points(kml_points_pool)

# 2. ถ้ามีจุดครบหัวท้ายจาก KML ให้หาเส้นทาง
if head_point and tail_point:
    route_coords, route_distance = get_osrm_route_head_tail(head_point, tail_point)

# --- แสดงผลแผนที่ ---
if uploaded_files or kml_elements:
    m = folium.Map(
        location=[13.75, 100.5], zoom_start=17, 
        tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", 
        attr="Google",
        control_scale=True
    )
    
    # แสดงเส้นทาง (เฉพาะจาก KML)
    if route_coords:
        folium.PolyLine(
            route_coords, 
            color="#007BFF", # สีฟ้า
            weight=5, 
            opacity=0.8, 
            dash_array='10, 10', 
            tooltip=f"🚶 ระยะทางตามแนวโครงข่าย (KMZ): {route_distance:,.0f} เมตร"
        ).add_to(m)
        
        st.info(f"📍 **ระยะทางตามแนวโครงข่าย (คำนวณจากไฟล์ KMZ/KML เท่านั้น):** {route_distance/1000:.3f} กม. ({route_distance:,.0f} เมตร)")
    elif kml_file and not route_coords:
        st.warning("⚠️ ไฟล์ KML ไม่มีข้อมูลพิกัดเพียงพอสำหรับคำนวณเส้นทาง")

    # เครื่องมือวัดระยะ Manual
    m.add_child(MeasureControl(
        position='topright', 
        primary_length_unit='meters', 
        secondary_length_unit='kilometers',
        active_color='#FF8C42',
        completed_color='#2D5A27'
    ))

    all_bounds = []

    # วาด KML (เป็นเส้นจางๆ พื้นหลัง)
    for elem in kml_elements:
        if elem['is_point']:
            loc = elem['points'][0]
            folium.Marker(loc, icon=folium.Icon(color='red', icon='info-sign')).add_to(m)
            folium.Marker(loc, icon=folium.DivIcon(html=create_div_label(elem['name']))).add_to(m)
            all_bounds.append(loc)
        else:
            # เส้น KML เดิม จางลง
            folium.PolyLine(elem['points'], color="gray", weight=2, opacity=0.4, dash_array='5').add_to(m)
            all_bounds.extend(elem['points'])

    # วาดรูปภาพจาก Session State (Marker รูปภาพ) - ไม่เกี่ยวกับเส้นทาง
    for data in st.session_state.export_data:
        icon_html = img_to_custom_icon(data['img_obj'], data['issue'])
        folium.Marker([data['lat'], data['lon']], icon=folium.DivIcon(html=icon_html)).add_to(m)
        all_bounds.append([data['lat'], data['lon']])

    if all_bounds: m.fit_bounds(all_bounds, padding=[50, 50])
    st_folium(m, height=900, use_container_width=True, key="survey_map")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📄 3. สร้างรายงาน PowerPoint")
    col1, col2 = st.columns([1, 1])
    with col1:
        map_cap = st.file_uploader("อัปโหลดรูป Capture แผนที่", type=['jpg','png'])
    if map_cap and st.session_state.export_data:
        with col2:
            st.write("")
            if st.button("🚀 สรุปรายงานและดาวน์โหลดไฟล์ PPTX"):
                pptx_data = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data)
                st.download_button("📥 คลิกเพื่อดาวน์โหลดรายงาน", data=pptx_data, file_name="Cable_AI_Report.pptx")

