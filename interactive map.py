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

# --- 1. การตั้งค่าระบบและการจัดการความปลอดภัย ---
ssl._create_default_https_context = ssl._create_unverified_context
client = genai.Client(api_key="AIzaSyBHAKfkjkb2wdzAZQZ74dFRD4Ib5Dj6cHY")

@st.cache_resource
def load_ocr():
    model_path = os.path.join(os.getcwd(), "easyocr_models")
    if not os.path.exists(model_path): os.makedirs(model_path)
    return easyocr.Reader(['en'], gpu=False, model_storage_directory=model_path)

# --- 2. ฟังก์ชันวิเคราะห์พิกัดและรูปภาพ ---
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
        img_np = np.array(image)
        results = reader.readtext(img_np, paragraph=True)
        full_text = " ".join([res[1] for res in results])
        match = re.search(r'(\d+\.\d+)\s*[nN]\s+(\d+\.\d+)\s*[eE]', full_text)
        if match: return float(match.group(1)), float(match.group(2))
    except: pass
    return None, None

def analyze_cable_issue(image_bytes):
    try:
        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=[
                "วิเคราะห์รูปภาพและเลือกตอบ 1 ข้อ: 1. cable ตกพื้น | 2. หัวต่ออยู่กลาง span | 3. ไฟไหม้ cable | 4. หัวต่อขวดน้ำ (ตอบแค่ชื่อสาเหตุ)",
                types.Part.from_bytes(data=image_bytes, mime_type="image/jpeg")
            ]
        )
        return response.text.strip()
    except: return "ไม่ระบุสาเหตุ"

# --- 3. ฟังก์ชันวัดระยะทาง (OSRM) ---
def get_walking_route(p1, p2):
    try:
        url = f"http://router.project-osrm.org/route/v1/foot/{p1[1]},{p1[0]};{p2[1]},{p2[0]}?overview=full&geometries=geojson"
        r = requests.get(url, timeout=5)
        data = r.json()
        if data['code'] == 'Ok':
            return data['routes'][0]['distance'], [[c[1], c[0]] for c in data['routes'][0]['geometry']['coordinates']]
    except: pass
    return None, None

# --- 4. ฟังก์ชัน Export PowerPoint ---
def create_survey_pptx(map_image_bytes, survey_data, distance_info=""):
    prs = Presentation()
    # Slide 1: แผนที่ภาพรวม
    if map_image_bytes:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(BytesIO(map_image_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
        if distance_info:
            tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(4), Inches(0.5))
            tx.text_frame.text = f"ระยะทางวัดจริง: {distance_info}"
    
    # Slide 2: รายละเอียดรูปภาพ (Grid 2x2)
    for i in range(0, len(survey_data), 4):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for j, item in enumerate(survey_data[i:i+4]):
            x = Inches(0.5 + (j % 2) * 4.8)
            y = Inches(0.5 + (j // 2) * 2.5)
            # เพิ่มรูป
            img_buf = BytesIO()
            item['img_obj'].save(img_buf, format="JPEG")
            img_buf.seek(0)
            slide.shapes.add_picture(img_buf, x, y, height=Inches(1.8))
            # เพิ่มคำบรรยาย
            txt = slide.shapes.add_textbox(x, y + Inches(1.85), Inches(4), Inches(0.5))
            txt.text_frame.text = f"สาเหตุ: {item['issue']}\nพิกัด: {item['lat']:.5f}, {item['lon']:.5f}"
            for p in txt.text_frame.paragraphs: p.font.size = Pt(10)

    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# --- 5. UI Main Logic ---
st.set_page_config(page_title="AI Cable Survey Pro", layout="wide")

if 'survey_data' not in st.session_state: st.session_state.survey_data = []
if 'selected_kml_pts' not in st.session_state: st.session_state.selected_kml_pts = []
if 'distance_result' not in st.session_state: st.session_state.distance_result = ""

st.title("🚧 AI Cable Plotter & Report")

# ส่วนที่ 1: KML (สำหรับวัดระยะ)
st.sidebar.header("📍 ข้อมูลโครงข่าย")
kml_file = st.sidebar.file_uploader("อัปโหลด KML/KMZ", type=['kml', 'kmz'])
kml_pts = []
if kml_file:
    # ... (ส่วนการ parse KML เหมือนเดิม) ...
    content = kml_file.getvalue()
    root = etree.fromstring(content)
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    for pm in root.xpath('.//kml:Placemark', namespaces=ns):
        name = pm.xpath('kml:name/text()', namespaces=ns)
        coord = pm.xpath('.//kml:coordinates/text()', namespaces=ns)
        if coord:
            c = coord[0].strip().split(',')[0:2]
            kml_pts.append({'name': name[0] if name else "P", 'lat': float(c[1]), 'lon': float(c[0])})

# ส่วนที่ 2: อัปโหลดรูปภาพ
st.subheader("📸 อัปโหลดรูปภาพสำรวจ")
up_files = st.file_uploader("เลือกรูปภาพ", type=['jpg','jpeg','png'], accept_multiple_files=True)

if up_files and len(st.session_state.survey_data) == 0:
    for f in up_files:
        raw_bytes = f.getvalue()
        img = ImageOps.exif_transpose(Image.open(BytesIO(raw_bytes)))
        lat, lon = get_lat_lon_exif(img)
        if not lat: lat, lon = get_lat_lon_ocr(img)
        if lat:
            issue = analyze_cable_issue(raw_bytes)
            st.session_state.survey_data.append({'img_obj': img, 'issue': issue, 'lat': lat, 'lon': lon})

# ส่วนที่ 3: แผนที่
m = folium.Map(location=[13.75, 100.5], zoom_start=14, tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}", attr="Google")
bounds = []

for d in st.session_state.survey_data:
    folium.Marker([d['lat'], d['lon']], tooltip=d['issue']).add_to(m)
    bounds.append([d['lat'], d['lon']])

if kml_pts:
    for p in kml_pts:
        is_sel = any(s['name'] == p['name'] for s in st.session_state.selected_kml_pts)
        folium.Marker([p['lat'], p['lon']], icon=folium.Icon(color='green' if is_sel else 'blue'), popup=f"SEL:{p['name']}").add_to(m)
        bounds.append([p['lat'], p['lon']])

if len(st.session_state.selected_kml_pts) == 2:
    d_m, path = get_walking_route([st.session_state.selected_kml_pts[0]['lat'], st.session_state.selected_kml_pts[0]['lon']],
                                  [st.session_state.selected_kml_pts[1]['lat'], st.session_state.selected_kml_pts[1]['lon']])
    if path:
        folium.PolyLine(path, color="red", weight=5).add_to(m)
        st.session_state.distance_result = f"{d_m:.2f} m."

if bounds: m.fit_bounds(bounds)
map_res = st_folium(m, height=500, use_container_width=True)

# รับการคลิกวัดระยะ
if kml_file and map_res['last_object_clicked_popup']:
    clicked = map_res['last_object_clicked_popup'].replace("SEL:", "")
    t = next((x for x in kml_pts if x['name'] == clicked), None)
    if t and t not in st.session_state.selected_kml_pts:
        if len(st.session_state.selected_kml_pts) >= 2: st.session_state.selected_kml_pts = []
        st.session_state.selected_kml_pts.append(t)
        st.rerun()

# ส่วนที่ 4: Export PowerPoint
st.subheader("📊 รายงาน PowerPoint")
map_cap = st.file_uploader("1. อัปโหลดรูป Capture หน้าจอแผนที่ (เพื่อความคมชัด)", type=['jpg','png'])

if st.button("🚀 สร้างรายงานและดาวน์โหลด"):
    if st.session_state.survey_data:
        pptx_file = create_survey_pptx(map_cap.getvalue() if map_cap else None, 
                                      st.session_state.survey_data, 
                                      st.session_state.distance_result)
        st.download_button("📥 คลิกเพื่อดาวน์โหลด PPTX", data=pptx_file, file_name="Survey_Report.pptx")
    else:
        st.error("กรุณาอัปโหลดรูปภาพที่มีพิกัดก่อน")
