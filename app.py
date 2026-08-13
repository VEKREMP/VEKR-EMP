import streamlit as st
import folium
from streamlit_folium import st_folium
import os
import fitz  
import pandas as pd
from docx import Document
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
import io
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ------------------------------------------------
# إعدادات واجهة الصفحة
# ------------------------------------------------
st.set_page_config(page_title="VEKR-EMP | منصة الخطط البيئية", page_icon="🌱", layout="wide")

# ------------------------------------------------
# دالة تصوير الخريطة
# ------------------------------------------------
def capture_map_screenshot(lat, lon, output_path):
    m = folium.Map(location=[lat, lon], zoom_start=15, tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}", attr="Google Satellite Hybrid")
    folium.Marker([lat, lon], tooltip="موقع المنشأة", icon=folium.Icon(color="blue", icon="info-sign")).add_to(m)
    folium.Circle(location=[lat, lon], radius=500, color="red", fill=True, fill_color="red", fill_opacity=0.2).add_to(m)
    html_path = output_path.replace(".png", ".html")
    m.save(html_path)
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--window-size=800,600')
    options.add_argument('--no-sandbox')
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("file://" + os.path.abspath(html_path))
    time.sleep(3)
    driver.save_screenshot(output_path)
    driver.quit()
    if os.path.exists(html_path): os.remove(html_path)

# ------------------------------------------------
# دالة تحويل المرفقات لصور
# ------------------------------------------------
def convert_to_image(uploaded_file, output_folder, base_name):
    file_ext = uploaded_file.name.split('.')[-1].lower()
    img_output_path = os.path.join(output_folder, f"{base_name}.png")
    
    if file_ext == "pdf":
        doc_pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        page = doc_pdf.load_page(0)
        pix = page.get_pixmap(dpi=150)
        pix.save(img_output_path)
    else:
        with open(img_output_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
    return img_output_path

# ------------------------------------------------
# الواجهة الرئيسية
# ------------------------------------------------
st.title("🌱 VEKR-EMP | منصة إعداد تقارير الخطط البيئية")
company_name = st.text_input("اسم الشركة")
city = st.text_input("المدينة")
address = st.text_area("العنوان الوطني")
lat = st.number_input("خط العرض", format="%f", value=0.0)
lon = st.number_input("خط الطول", format="%f", value=0.0)

# --- استعادة خانات رفع الصور ---
st.subheader("📎 مرفقات المستندات الرسمية")
cr_file = st.file_uploader("السجل التجاري", type=["pdf", "png", "jpg"], key="cr")
vat_file = st.file_uploader("شهادة ضريبة القيمة المضافة", type=["pdf", "png", "jpg"], key="vat")
address_file = st.file_uploader("العنوان الوطني", type=["pdf", "png", "jpg"], key="addr")

if st.button("إنشاء التقرير (Word) 📝"):
    if not company_name:
        st.error("الرجاء إدخال اسم الشركة")
    else:
        project_folder = os.path.join("Projects", company_name)
        os.makedirs(project_folder, exist_ok=True)
        
        # معالجة الصور
        cr_img = convert_to_image(cr_file, project_folder, "cr") if cr_file else None
        vat_img = convert_to_image(vat_file, project_folder, "vat") if vat_file else None
        addr_img = convert_to_image(address_file, project_folder, "addr") if address_file else None
        map_path = os.path.join(project_folder, "map.png")
        capture_map_screenshot(lat, lon, map_path)
        
        # إنشاء التقرير
        doc = DocxTemplate("template.docx")
        context = {
            'company_name': company_name,
            'city': city,
            'cr_image': InlineImage(doc, cr_img, width=Inches(5)) if cr_img else "",
            'vat_image': InlineImage(doc, vat_img, width=Inches(5)) if vat_img else "",
            'address_image': InlineImage(doc, addr_img, width=Inches(5)) if addr_img else "",
            'map_image': InlineImage(doc, map_path, width=Inches(5))
        }
        
        doc.render(context)
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.download_button("📥 تحميل التقرير", data=buffer, file_name="Report.docx")
