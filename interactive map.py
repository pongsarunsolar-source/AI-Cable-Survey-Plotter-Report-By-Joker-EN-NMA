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
from google import genai
from google.genai import types
import zipfile
from lxml import etree
import math

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
                ตอบเฉพาะชื่อสาเหตุภาษาไทยเท่านั้น หากวิเคราะห์ไม่ได้ให้ตอบว่า cable ตกพื้น""",
                types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg")
            ]
        )
        result = response.text.strip()
        return result if result and "วิเคราะห์ไม่ได้" not in result else "cable ตกพื้น"
    except: return "cable ตกพื้น"

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
        img_for_ocr = image.copy().thumbnail((1000, 1000)) 
        img_np = np.array(image.convert('RGB'))
        results = reader.readtext(img_np, paragraph=True, allowlist='0123456789.NE ne')
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN].*?(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

# --- 5. ฟังก์ชันอ่านไฟล์ KML/KMZ ---
def parse_kml_data(file):
    elements, points_pool = [], []
    try:
        if file.name.endswith('.kmz'):
            with zipfile.ZipFile(file) as z:
                kml_filename = [n for n in z.namelist() if n.endswith('.kml')][0]
                content = z.read(kml_filename)
        else: content = file.getvalue()
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
                elements.append({'name': final_name, 'points': pts, 'is_point': len(pts) == 1})
                points_pool.extend(pts)
        return elements, points_pool
    except: return [], []

def get_farthest_points(coordinates):
    if not coordinates or len(coordinates) < 2: return None, None
    try:
        pts = np.array(coordinates)
        if len(pts) > 200:
            test_points = [pts[pts[:,0].argmax()], pts[pts[:,0].argmin()], pts[pts[:,1].argmax()], pts[pts[:,1].argmin()]]
        else: test_points = pts
        max_dist, p1_best, p2_best = -1, None, None
        for i in range(len(test_points)):
            for j in range(i + 1, len(test_points)):
                dist = (test_points[i][0] - test_points[j][0])**2 + (test_points[i][1] - test_points[j][1])**2
                if dist > max_dist: max_dist, p1_best, p2_best = dist, test_points[i], test_points[j]
        return p1_best, p2_best
    except: return None, None

def get_osrm_route_head_tail(start_coord, end_coord):
    if not start_coord or not end_coord: return None, 0
    url = f"http://router.project-osrm.org/route/v1/walking/{start_coord[1]},{start_coord[0]};{end_coord[1]},{end_coord[0]}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if data.get("routes"):
                route = data["routes"][0]
                return [[lat, lon] for lon, lat in route["geometry"]["coordinates"]], route["distance"]
    except: pass
    return None, 0

# --- 6. ฟังก์ชันสร้าง Label ชื่อ ---
def create_div_label(name, color="#D9534F"):
    return f'''<div style="font-size: 11px; font-weight: 800; color: {color}; white-space: nowrap; transform: translate(-50%, -150%); background-color: transparent; text-shadow: 2px 2px 4px white, -2px -2px 4px white, 2px -2px 4px white, -2px 2px 4px white;">{name}</div>'''

def img_to_custom_icon(img, issue_text):
    img_resized = img.copy(); img_resized.thumbnail((150, 150))
    buf = BytesIO(); img_resized.save(buf, format="JPEG", quality=70)
    img_str = base64.b64encode(buf.getvalue()).decode()
    return f'''
        <div style="position: relative; width: fit-content; background-color: white; padding: 5px; border-radius: 12px; box-shadow: 0px 8px 24px rgba(0,0,0,0.12); border: 2px solid #FF8C42; transform: translate(-50%, -100%);">
            <div style="font-size: 11px; font-weight: 700; color: #2D5A27; margin-bottom: 4px; text-align: center;">{issue_text}</div>
            <img src="data:image/jpeg;base64,{img_str}" style="max-width: 140px; display: block; border-radius: 4px;">
            <div style="position: absolute; bottom: -10px; left: 50%; transform: translateX(-50%); width: 0; height: 0; border-left: 10px solid transparent; border-right: 10px solid transparent; border-top: 10px solid #FF8C42;"></div>
        </div>
    '''

# --- 7. ฟังก์ชันสร้างรายงาน PowerPoint (Final Update) ---
def create_summary_pptx(map_image_bytes, image_list, cable_type, route_distance, issue_kml_elements):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    
    # --- หน้า 1: สรุปรายละเอียด ---
    slide0 = prs.slides.add_slide(prs.slide_layouts[6])
    t0 = slide0.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1)).text_frame.paragraphs[0]
    t0.text = f"รายงานสรุปแนวทางแก้ไขปัญหาและเสนอคร่อม Cable ({cable_type} Core)"
    t0.font.bold, t0.font.size = True, Pt(22)
    
    tf = slide0.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(3.5)).text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]; p1.text = f"• Type Cable: {cable_type} Core"; p1.font.size = Pt(16)
    p2 = tf.add_paragraph(); p2.text = f"• ระยะคร่อม Cable รวม: {route_distance:,.0f} เมตร ({route_distance/1000:.3f} กม.)"; p2.font.size = Pt(16)
    p3 = tf.add_paragraph(); p3.text = f"• รายละเอียดจุดปัญหา:"; p3.font.bold, p3.font.size = True, Pt(16)
    
    for el in issue_kml_elements[:10]:
        p_el = tf.add_paragraph()
        p_el.text = f"  - {el['name']} (Lat: {el['points'][0][0]:.5f}, Long: {el['points'][0][1]:.5f})"
        p_el.font.size = Pt(12)

    # --- หน้า 2: Topology Overall (รูปเต็มหน้า + หัวข้อขีดเส้นใต้) ---
    if map_image_bytes:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        slide1.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        title_box1 = slide1.shapes.add_textbox(Inches(0.2), Inches(0.1), Inches(4), Inches(0.5))
        p_title1 = title_box1.text_frame.paragraphs[0]
        p_title1.text = "Topology Overall"
        p_title1.font.bold, p_title1.font.size = True, Pt(24)
        p_title1.font.underline = True # ขีดเส้นใต้ตามสั่ง

    # --- หน้า 3: รูปภาพแสดงจุดที่มีปัญหา ---
    if image_list:
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        t2 = slide2.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text_frame.paragraphs[0]
        t2.text = "รูปภาพแสดงจุดที่มีปัญหา"; t2.font.bold, t2.font.size = True, Pt(20)

        cols, rows = 4, 2
        img_w, img_h = Inches(2.1), Inches(1.5)
        margin_x, start_y = (prs.slide_width - (img_w * cols)) / (cols + 1), Inches(0.8)
        
        for i, item in enumerate(image_list[:8]):
            x, y = margin_x + ((i % cols) * (img_w + margin_x)), start_y + ((i // cols) * (img_h + Inches(0.8)))
            image = item['img_obj'].copy()
            # Crop to aspect ratio 2.1:1.5
            target_ratio = img_w / img_h
            w_px, h_px = image.size
            if (w_px/h_px) > target_ratio:
                new_w = h_px * target_ratio
                image = image.crop(((w_px - new_w) / 2, 0, (w_px + new_w) / 2, h_px))
            else:
                new_h = w_px / target_ratio
                image = image.crop((0, (h_px - new_h) / 2, w_px, (h_px + new_h) / 2))
            
            buf = BytesIO(); image.save(buf, format="JPEG"); buf.seek(0)
            slide2.shapes.add_picture(buf, x, y, width=img_w, height=img_h)
            txt_box = slide2.shapes.add_textbox(x, y + img_h + Inches(0.05), img_w, Inches(0.6)).text_frame
            p_iss = txt_box.paragraphs[0]; p_iss.text = f"สาเหตุ: {item['issue']}"; p_iss.font.size = Pt(8); p_iss.font.bold = True
            p_lat = txt_box.add_paragraph(); p_lat.text = f"Lat: {item['lat']:.5f}\nLong: {item['lon']:.5f}"; p_lat.font.size = Pt(7)
            
    output = BytesIO(); prs.save(output)
    return output.getvalue()

# --- 8. UI Layout ---
st.set_page_config(page_title="AI Cable Survey", layout="wide")
st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
    .joker-icon { width: 100px; height: 100px; object-fit: cover; border-radius: 50%; border: 4px solid #FFFFFF; outline: 3px solid #FF8C42; }
</style>""", unsafe_allow_html=True)

joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f'''<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA</p></div>{"<img src='data:image/png;base64,"+joker_base64+"' class='joker-icon'>" if joker_base64 else ""}</div>''', unsafe_allow_html=True)

# --- 9. Logic ส่วนควบคุม ---
kml_file_yellow = st.file_uploader("Import KMZ - Overall (ภาพรวมแผนที่)", type=['kml', 'kmz'])
kml_file = st.file_uploader("Import KMZ - พิกัดที่มีปัญหาและเสนอคร่อม cable", type=['kml', 'kmz'])

zoom_bounds = []
y_elements, k_elements, k_pool = [], [], []

if kml_file_yellow:
    y_elements, _ = parse_kml_data(kml_file_yellow)
    for el in y_elements: zoom_bounds.extend(el['points'])
if kml_file:
    k_elements, k_pool = parse_kml_data(kml_file)
    for el in k_elements: zoom_bounds.extend(el['points'])

st.subheader("📁 2. อัปโหลดรูปภาพสำรวจ")
uploaded_files = st.file_uploader("ลากและวางไฟล์ที่นี่", type=['jpg','jpeg','png'], accept_multiple_files=True)
if 'export_data' not in st.session_state: st.session_state.export_data = []

if uploaded_files:
    for f in uploaded_files:
        raw_data = f.getvalue(); raw_img = Image.open(BytesIO(raw_data))
        lat, lon = get_lat_lon_exif(raw_img)
        if lat:
            issue = analyze_cable_issue(raw_data)
            st.session_state.export_data.append({'img_obj': ImageOps.exif_transpose(raw_img), 'issue': issue, 'lat': lat, 'lon': lon})
            zoom_bounds.append([lat, lon])

r_coords, r_dist = None, 0
if k_pool:
    p1, p2 = get_farthest_points(k_pool)
    r_coords, r_dist = get_osrm_route_head_tail(p1, p2)

if uploaded_files or k_elements or y_elements:
    m = folium.Map(location=[13.75, 100.5], zoom_start=17, tiles=None)
    folium.TileLayer(tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google", name="Google Maps", opacity=0.4).add_to(m)
    if r_coords: folium.PolyLine(r_coords, color="#FF0000", weight=5, dash_array='10, 10').add_to(m)
    for el in y_elements:
        if el['is_point']: folium.Marker(el['points'][0], icon=folium.DivIcon(html=create_div_label(el['name'], "#CC9900"))).add_to(m)
        else: folium.PolyLine(el['points'], color="#FFD700", weight=4).add_to(m)
    for el in k_elements:
        if el['is_point']: folium.Marker(el['points'][0], icon=folium.DivIcon(html=create_div_label(el['name'], "#D9534F"))).add_to(m)
        else: folium.PolyLine(el['points'], color="gray", weight=2, dash_array='5').add_to(m)
    for d in st.session_state.export_data: folium.Marker([d['lat'], d['lon']], icon=folium.DivIcon(html=img_to_custom_icon(d['img_obj'], d['issue']))).add_to(m)
    if zoom_bounds: m.fit_bounds(zoom_bounds, padding=[50, 50])
    st_folium(m, height=1200, use_container_width=True)

st.subheader("📄 3. สร้างรายงาน PowerPoint")
c1, c2 = st.columns(2)
with c1:
    cable_type = st.selectbox("เลือก Type Cable", ["4", "6", "12", "24", "48", "96"])
    map_cap = st.file_uploader("อัปโหลดรูป Capture แผนที่", type=['jpg','png'])
if map_cap and st.button("🚀 ดาวน์โหลดรายงาน PPTX"):
    pptx_data = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data, cable_type, r_dist, k_elements)
    st.download_button("📥 คลิกเพื่อดาวน์โหลด", data=pptx_data, file_name=f"Cable_Survey_{cable_type}C.pptx")
