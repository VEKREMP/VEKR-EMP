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
# دالة تحويل المرفقات
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
# ... (باقي الخانات كما هي في الكود السابق)

# --- إضافة خانة صور الزيارة ---
st.subheader("📸 صور الزيارة الميدانية")
visit_photos = st.file_uploader("ارفع صور الزيارة (توضع في الملاحق):", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

# ------------------------------------------------
# قسم إصدار التقرير
# ------------------------------------------------
if st.button("إنشاء التقرير (Word) 📝"):
    project_folder = os.path.join("Projects", company_name)
    os.makedirs(project_folder, exist_ok=True)
    
    # معالجة صور الزيارة
    visit_image_objects = []
    if visit_photos:
        for i, photo in enumerate(visit_photos):
            photo_path = os.path.join(project_folder, f"visit_{i}.png")
            with open(photo_path, "wb") as f:
                f.write(photo.getbuffer())
            visit_image_objects.append(photo_path)

    # إنشاء التقرير
    doc = DocxTemplate("template.docx")
    
    # تجهيز صور الزيارة لتدخل في الملحق كقائمة من الصور
    visit_images_inline = [InlineImage(doc, img, width=Inches(5)) for img in visit_image_objects]
    
    context = {
        # ... (باقي المتغيرات)
        'visit_images': visit_images_inline 
    }
    
    doc.render(context)
    # ... (باقي كود الحفظ والتحميل)
