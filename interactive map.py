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
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from google import genai
from google.genai import types
import zipfile
from lxml import etree
import math
from datetime import datetime

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

@st.cache_data
def load_template_bytes(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.content
    except: pass
    return None

# --- 2. ฟังก์ชันช่วยดึงรูปภาพ Joker ---
def get_image_base64_from_drive(file_id):
    try:
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return base64.b64encode(response.content).decode()
    except: return None
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
        candidates = [pts[pts[:,0].argmax()], pts[pts[:,0].argmin()], pts[pts[:,1].argmax()], pts[pts[:,1].argmin()]]
        max_dist, p1_best, p2_best = -1, None, None
        for i in range(len(candidates)):
            for j in range(i + 1, len(candidates)):
                dist = (candidates[i][0] - candidates[j][0])**2 + (candidates[i][1] - candidates[j][1])**2
                if dist > max_dist: max_dist, p1_best, p2_best = dist, candidates[i], candidates[j]
        return p1_best, p2_best
    except: return None, None

def get_osrm_route_head_tail(start_coord, end_coord):
    if start_coord is None or end_coord is None: return None, 0
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

# --- 7. ฟังก์ชันสร้างรายงาน PowerPoint ---
def create_summary_pptx(map_image_bytes, image_list, cable_type, route_distance, issue_kml_elements, template_bytes=None):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
    
    def apply_background(slide):
        if template_bytes:
            slide.shapes.add_picture(BytesIO(template_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)

    # ==========================================
    # --- หน้าที่ 1: หน้าปก (Cover Slide) ---
    # ==========================================
    slide_cover = prs.slides.add_slide(prs.slide_layouts[6]); apply_background(slide_cover)
    cover_box = slide_cover.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(7.5), Inches(2))
    tf_cover = cover_box.text_frame
    
    p1 = tf_cover.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run(); r1.text = "เอกสารประกอบ "; r1.font.size = Pt(32); r1.font.color.rgb = RGBColor(0, 86, 179)
    r2 = p1.add_run(); r2.text = "Imp_NMA-XX"; r2.font.size = Pt(36); r2.font.bold = True; r2.font.color.rgb = RGBColor(0, 86, 179)
    
    p2 = tf_cover.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r3 = p2.add_run(); r3.text = "ข้อมูลนำเสนอปรับปรุง EN-NMA OSP\n"; r3.font.size = Pt(28); r3.font.color.rgb = RGBColor(0, 86, 179)
    
    p3 = tf_cover.add_paragraph(); p3.alignment = PP_ALIGN.CENTER
    r4 = p3.add_run(); r4.text = "Improve Site XXXX"; r4.font.size = Pt(36); r4.font.bold = True; r4.font.color.rgb = RGBColor(0, 86, 179)

    ver_box = slide_cover.shapes.add_textbox(Inches(0.2), Inches(5.1), Inches(4), Inches(0.5))
    p_ver = ver_box.text_frame.paragraphs[0]
    p_ver.text = f"Ver.Update Data ปัจจุบัน {datetime.now().strftime('%d/%m/%Y')}"
    p_ver.font.size, p_ver.font.color.rgb = Pt(12), RGBColor(0, 0, 0)

    # ==========================================
    # --- หน้าที่ 2: สรุปแนวทางแก้ไขปัญหา ---
    # ==========================================
    slide0 = prs.slides.add_slide(prs.slide_layouts[6]); apply_background(slide0) 
    
    # หัวข้อหลัก
    title_box = slide0.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(7.5), Inches(0.8))
    p_title = title_box.text_frame.paragraphs[0]
    p_title.text = f"รายงานสรุปแนวทางแก้ไขปัญหาและเสนอคร่อม Cable ({cable_type} Core)"
    p_title.font.bold = True
    p_title.font.size = Pt(22)

    # ปัญหา สาเหตุและผลกระทบ
    prob_box = slide0.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(7.5), Inches(0.5))
    p_prob = prob_box.text_frame.paragraphs[0]
    p_prob.text = "ปัญหา สาเหตุและผลกระทบ"
    p_prob.font.bold = True
    p_prob.font.underline = True
    p_prob.font.size = Pt(16)

    # กล่องสำหรับพิมพ์ปัญหา
    shape_box = slide0.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.4), Inches(7.0), Inches(0.8))
    shape_box.fill.background() 
    shape_box.line.color.rgb = RGBColor(0, 0, 0) 
    p_guide = shape_box.text_frame.paragraphs[0]
    p_guide.text = " (คลิกเพื่อพิมพ์ปัญหา สาเหตุ และผลกระทบ...)"
    p_guide.font.color.rgb = RGBColor(128, 128, 128)
    p_guide.font.size = Pt(12)

    # Scope Of Work
    scope_box = slide0.shapes.add_textbox(Inches(0.5), Inches(2.3), Inches(7.5), Inches(3.0))
    tf_scope = scope_box.text_frame; tf_scope.word_wrap = True

    p_scope = tf_scope.paragraphs[0]
    p_scope.text = "Scope Of Work"
    p_scope.font.bold = True
    p_scope.font.underline = True
    p_scope.font.size = Pt(16)

    p_type = tf_scope.add_paragraph()
    p_type.text = f"• ขอ Replace Cable : {cable_type} Core"
    p_type.font.size = Pt(14)

    p_dist = tf_scope.add_paragraph()
    if route_distance:
        p_dist.text = f"• ระยะคร่อม Cable รวม: {route_distance:,.0f} เมตร ({route_distance/1000:.3f} กม.)"
    else:
        p_dist.text = f"• ระยะคร่อม Cable รวม: 0 เมตร (0.000 กม.)"
    p_dist.font.size = Pt(14)

    p_detail_title = tf_scope.add_paragraph()
    p_detail_title.text = "รายละเอียดจุดปัญหา:"
    p_detail_title.font.bold = True
    p_detail_title.font.underline = True
    p_detail_title.font.size = Pt(14)

    for el in issue_kml_elements[:10]:
        p_el = tf_scope.add_paragraph()
        p_el.text = f"  - {el['name']} (Lat: {el['points'][0][0]:.5f}, Long: {el['points'][0][1]:.5f})"
        p_el.font.size = Pt(12)

    # ==========================================
    # --- หน้าที่ 3: Topology Overall ---
    # ==========================================
    if map_image_bytes:
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        slide1.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
        t_box1 = slide1.shapes.add_textbox(Inches(0.2), Inches(0.1), Inches(5), Inches(0.5))
        p_t1 = t_box1.text_frame.paragraphs[0]; p_t1.text = "Topology Overall"; p_t1.font.bold, p_t1.font.size, p_t1.font.underline = True, Pt(24), True
        
    # ==========================================
    # --- หน้าที่ 4: รูปภาพจุดปัญหา ---
    # ==========================================
    if image_list:
        slide2 = prs.slides.add_slide(prs.slide_layouts[6]); apply_background(slide2)
        t_box2 = slide2.shapes.add_textbox(Inches(0.2), Inches(0.1), Inches(6), Inches(0.5))
        p_t2 = t_box2.text_frame.paragraphs[0]; p_t2.text = "รูปภาพแสดงจุดที่มีปัญหา"; p_t2.font.bold, p_t2.font.size, p_t2.font.underline = True, Pt(22), True
        cols, rows, img_w, img_h = 4, 2, Inches(1.8), Inches(1.3)
        margin_x, start_y = (Inches(7.8) - (img_w * cols)) / (cols + 1), Inches(0.9)
        for i, item in enumerate(image_list[:8]):
            curr_row, curr_col = i // cols, i % cols
            x, y = margin_x + (curr_col * (img_w + margin_x)), start_y + (curr_row * (img_h + Inches(0.8))) 
            image = item['img_obj'].copy(); target_ratio = img_w / img_h; w_px, h_px = image.size
            if (w_px/h_px) > target_ratio:
                new_w = h_px * target_ratio; image = image.crop(((w_px-new_w)/2, 0, (w_px+new_w)/2, h_px))
            else:
                new_h = w_px / target_ratio; image = image.crop((0, (h_px-new_h)/2, w_px, (h_px+new_h)/2))
            buf = BytesIO(); image.save(buf, format="JPEG"); buf.seek(0)
            slide2.shapes.add_picture(buf, x, y, width=img_w, height=img_h)
            txt = slide2.shapes.add_textbox(x, y + img_h + Inches(0.02), img_w, Inches(0.6)).text_frame; txt.word_wrap = True
            p_iss = txt.paragraphs[0]; p_iss.text = f"สาเหตุ: {item['issue']}"; p_iss.font.size, p_iss.font.bold = Pt(8), True
            p_lat = txt.add_paragraph(); p_lat.text = f"Lat: {item['lat']:.5f}\nLong: {item['lon']:.5f}"; p_lat.font.size = Pt(7)
            
    output = BytesIO(); prs.save(output)
    return output.getvalue()

# --- 8. UI Layout & Logic ---
st.set_page_config(page_title="AI Cable Plotter", layout="wide")
st.markdown("""<style>
    .stApp { background: linear-gradient(120deg, #FFF5ED 0%, #F0F9F1 100%); }
    .header-container { display: flex; align-items: center; justify-content: space-between; padding: 25px; background: white; border-radius: 24px; border-bottom: 5px solid #FF8C42; margin-bottom: 30px; }
    .main-title { background: linear-gradient(90deg, #2D5A27 0%, #FF8C42 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800; font-size: 2.6rem; margin: 0; }
    .joker-icon { width: 100px; height: 100px; object-fit: cover; border-radius: 50%; border: 4px solid #FFFFFF; outline: 3px solid #FF8C42; }
    
    /* ปรับแต่งปุ่มดาวน์โหลด (พาสเทล เขียว-ส้ม) */
    .stDownloadButton>button { 
        background: linear-gradient(90deg, #A8E6CF 0%, #FFD3B6 100%); 
        color: #2D5A27; 
        border-radius: 14px; 
        padding: 15px 35px; 
        font-weight: 800; 
        width: 100%; 
        border: none;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.1);
        transition: transform 0.2s;
    }
    .stDownloadButton>button:hover { transform: scale(1.02); }
</style>""", unsafe_allow_html=True)

joker_base64 = get_image_base64_from_drive("1_G_r4yKyBA_vv3Nf8SdFpQ8UKv4bPLBr")
st.markdown(f'''<div class="header-container"><div><h1 class="main-title">AI Cable Plotter</h1><p style="margin:0; color: #718096; font-weight: 600;">By Joker EN-NMA</p></div>{"<img src='data:image/png;base64,"+joker_base64+"' class='joker-icon'>" if joker_base64 else ""}</div>''', unsafe_allow_html=True)

st.subheader("🌐 1. ข้อมูลโครงข่าย & จุดติดตั้ง (KML/KMZ)")
kml_file_y = st.file_uploader("Import KMZ - Overall (ภาพรวมแผนที่)", type=['kml', 'kmz'])
kml_file_r = st.file_uploader("Import KMZ - พิกัดที่มีปัญหาและเสนอคร่อม cable", type=['kml', 'kmz'])

z_bounds, k_els, k_pool, y_els = [], [], [], []
if kml_file_y:
    y_els, _ = parse_kml_data(kml_file_y)
    for el in y_els: z_bounds.extend(el['points'])
if kml_file_r:
    k_els, k_pool = parse_kml_data(kml_file_r)
    for el in k_els: z_bounds.extend(el['points'])

st.markdown("<hr>", unsafe_allow_html=True); st.subheader("📁 2. อัปโหลดรูปภาพสำรวจ")
uploaded_files = st.file_uploader("ลากและวางไฟล์ที่นี่", type=['jpg','jpeg','png'], accept_multiple_files=True, key="survey_uploader")
if 'export_data' not in st.session_state: st.session_state.export_data = []

if uploaded_files:
    curr_h = "".join([f.name + str(f.size) for f in uploaded_files])
    if 'last_hash' not in st.session_state or st.session_state.last_hash != curr_h:
        st.session_state.export_data, st.session_state.last_hash = [], curr_h
        for f in uploaded_files:
            raw_d = f.getvalue(); raw_img = Image.open(BytesIO(raw_d)); lat, lon = get_lat_lon_exif(raw_img)
            if lat is None: lat, lon = get_lat_lon_ocr(raw_img)
            if lat:
                issue = analyze_cable_issue(raw_d)
                st.session_state.export_data.append({'img_obj': ImageOps.exif_transpose(raw_img), 'issue': issue, 'lat': lat, 'lon': lon})
                z_bounds.append([lat, lon])

r_coords, r_dist = None, 0
if k_pool:
    try:
        p1, p2 = get_farthest_points(k_pool)
        if p1 is not None and p2 is not None: r_coords, r_dist = get_osrm_route_head_tail(p1, p2)
    except: pass

if uploaded_files or k_els or y_els:
    m = folium.Map(location=[13.75, 100.5], zoom_start=17, tiles=None, control_scale=True)
    folium.TileLayer(tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google", name="Google Maps", opacity=0.4, overlay=False).add_to(m)
    if r_coords: folium.PolyLine(r_coords, color="#FF0000", weight=5, opacity=0.8, dash_array='10, 10').add_to(m)
    for el in y_els:
        if el['is_point']: folium.Marker(el['points'][0], icon=folium.DivIcon(html=create_div_label(el['name'], "#CC9900"))).add_to(m)
        else: folium.PolyLine(el['points'], color="#FFD700", weight=4, opacity=0.8).add_to(m)
    for el in k_els:
        if el['is_point']: folium.Marker(el['points'][0], icon=folium.DivIcon(html=create_div_label(el['name'], "#D9534F"))).add_to(m)
        else: folium.PolyLine(el['points'], color="gray", weight=2, opacity=0.4, dash_array='5').add_to(m)
    for d in st.session_state.export_data: folium.Marker([d['lat'], d['lon']], icon=folium.DivIcon(html=img_to_custom_icon(d['img_obj'], d['issue']))).add_to(m)
    m.add_child(MeasureControl(position='topright', primary_length_unit='meters'))
    if z_bounds: m.fit_bounds(z_bounds, padding=[50, 50])
    st_folium(m, height=1200, use_container_width=True, key="survey_map")

st.markdown("<hr>", unsafe_allow_html=True); st.subheader("📄 3. สร้างรายงาน PowerPoint")
col_c1, col_c2 = st.columns(2)
with col_c1:
    cable_type = st.selectbox("เลือก Type Cable", ["4", "6", "12", "24", "48", "96"])
    map_cap = st.file_uploader("อัปโหลดรูป Capture แผนที่", type=['jpg','png'])

if map_cap:
    with col_c2:
        try:
            bg_t_id = "1EqtiR6CVnsbsVIg5Gk5j1v901YXYzjkz"
            t_bytes = load_template_bytes(bg_t_id)
            
            # ทำการ Generate PPTX ทันทีที่อัปโหลดแผนที่เสร็จ
            pptx_data = create_summary_pptx(map_cap.getvalue(), st.session_state.export_data, cable_type, r_dist, k_els, t_bytes)
            
            # HTML แต่งปุ่มให้มีรูป Joker
            btn_label = " ดาวน์โหลดรายงาน PPTX"
            if joker_base64:
                st.markdown(f"""
                <div style="text-align:center; margin-bottom:-45px; position:relative; z-index:10; pointer-events:none;">
                    <img src='data:image/png;base64,{joker_base64}' style='width:30px; height:30px; border-radius:50%; border:2px solid white; vertical-align:middle; margin-right:5px;'>
                    <span style='font-weight:800; color:#2D5A27; vertical-align:middle;'>{btn_label}</span>
                </div>
                """, unsafe_allow_html=True)
                btn_label = " " # ให้ตัวหนังสือบนปุ่มจริงว่างเปล่า เพราะเราใช้ HTML ทับแล้ว

            st.download_button(
                label=btn_label, 
                data=pptx_data, 
                file_name=f"Cable_Survey_{cable_type}C.pptx", 
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e: st.error(f"เกิดข้อผิดพลาดในการสร้างรายงาน: {e}")
