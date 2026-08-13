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
# إعدادات واجهة الصفحة (العنوان والأيقونة والهوية)
# ------------------------------------------------
st.set_page_config(
    page_title="VEKR-EMP | منصة الخطط البيئية",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تصميم عصري وخلفية متناسقة للمنصة
page_bg_img = """
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(to bottom right, #f4f9f4, #e2f0e2);
    background-size: cover;
}
[data-testid="stHeader"] {
    background: rgba(0,0,0,0);
}
.main-title {
    color: #1b4332;
    font-weight: 800;
    text-align: center;
    margin-bottom: 20px;
}
.stButton>button {
    background-color: #2d6a4f;
    color: white;
    border-radius: 8px;
    font-weight: bold;
    width: 100%;
}
.stButton>button:hover {
    background-color: #1b4332;
    color: #ffffff;
}
</style>
"""
st.markdown(page_bg_img, unsafe_allow_html=True)

# ------------------------------------------------
# دالة تصوير الخريطة بدقة عالية (مع نصف قطر 500 متر وضمان الحفظ)
# ------------------------------------------------
def capture_map_screenshot(lat, lon, output_path):
    m = folium.Map(
        location=[lat, lon], 
        zoom_start=15,
        tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
        attr="Google Satellite Hybrid"
    )
    folium.Marker([lat, lon], tooltip="موقع المنشأة", icon=folium.Icon(color="blue", icon="info-sign")).add_to(m)
    # رسم دائرة نصف قطرها 500 متر بدقة
    folium.Circle(
        location=[lat, lon], 
        radius=500, 
        color="#ff4d4d", 
        weight=2,
        fill=True, 
        fill_color="#ff4d4d", 
        fill_opacity=0.25
    ).add_to(m)
    
    html_path = output_path.replace(".png", ".html")
    m.save(html_path)
    
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--window-size=1000,800')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("file://" + os.path.abspath(html_path))
    
    # الانتظار حتى تحميل صور الأقمار الصناعية بالكامل
    time.sleep(6)
    driver.save_screenshot(output_path)
    driver.quit()
    
    try:
        os.remove(html_path)
    except:
        pass
        
    # التأكد التام من أن ملف الصورة تم إنشاؤه واستقر على القرص
    for _ in range(10):
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            break
        time.sleep(0.5)

# ------------------------------------------------
# دالة تحويل المرفقات إلى صور بدقة عالية
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
# تحميل بيانات الأنشطة من ملف الأكسل
# ------------------------------------------------
@st.cache_data
def load_activities():
    df = pd.read_excel('Classifications.xlsx')
    activities_dict = df.set_index('رمز النشاط')['الوصف الفني التفصيلي للنشاط'].to_dict()
    activity_names = df.set_index('رمز النشاط')['النشاط (حسب آيزك)'].to_dict()
    return activities_dict, activity_names

activities_desc, activities_names = load_activities()

# ------------------------------------------------
# محتوى الواجهة الرئيسية
# ------------------------------------------------
try:
    st.image("Logo.jpg", width=180) 
except:
    pass

st.markdown("<h1 class='main-title'>🌱 VEKR-EMP | منصة إعداد تقارير الخطط البيئية</h1>", unsafe_allow_html=True)
st.write("الرجاء إدخال بيانات المنشأة والمستندات المطلوبة أدناه لإصدار التقرير البيئي:")

# ------------------------------------------------
# 1. قسم إدخال المعطيات الأساسية
# ------------------------------------------------
st.subheader("📋 البيانات الأساسية للمنشأة")
col_a, col_b = st.columns(2)
with col_a:
    company_name = st.text_input("اسم الشركة")
    city = st.text_input("المدينة")
with col_b:
    tax_id = st.text_input("الرقم الضريبي")
    project_area = st.number_input("مساحة المشروع (متر مربع)", min_value=0.0, format="%g")

address = st.text_area("العنوان الوطني الخاص بالنشاط")
project_objective = st.text_area("هدف المشروع")

st.markdown("---")
st.subheader("🔍 اختيار النشاط (تصنيف آيزك)")
st.caption("في حال عدم وجود النشاط يرجى التواصل مع الدعم الفني من خلال الإيميل: adminsupport@vekr.uk")

search_query = st.selectbox(
    "ابحث أو اختر رمز النشاط:", 
    options=list(activities_desc.keys()), 
    format_func=lambda x: f"{x} - {activities_names.get(x, '')}"
)

selected_activity_name = activities_names.get(search_query, "")
selected_description = activities_desc.get(search_query, "")
activity_description = st.text_area("الوصف الفني للنشاط (مُولد تلقائياً - يمكنك تعديله):", value=selected_description, height=130)

st.markdown("---")
st.subheader("📍 الموقع الجغرافي")

col1, col2 = st.columns(2)
with col1:
    lat = st.number_input("خط العرض (Latitude)", format="%f", value=0.0)
with col2:
    lon = st.number_input("خط الطول (Longitude)", format="%f", value=0.0)

# ------------------------------------------------
# 2. قسم إدارة المخلفات والنفايات
# ------------------------------------------------
st.markdown("---")
st.subheader("♻️ إدارة المخلفات والنفايات")

solid_waste_default = "تتكون المخلفات الصلبة الناتجة عن أنشطة المشروع بشكل رئيسي من مخلفات عامة غير خطرة ناتجة عن الأنشطة اليومية والإدارية. يتم إدارة هذه المخلفات من خلال تجميعها دورياً في حاويات مخصصة ومحكمة الغلق، مع تطبيق مبدأ الفرز من المصدر، ومن ثم نقلها والتخلص منها عبر نظام الحاويات البلدية المعتمد."
liquid_waste_default = "تشمل المخلفات السائلة الناتجة عن المشروع مياه الصرف الصحي الناتجة عن الاستخدامات التشغيلية والخدمية للمرافق. يتم تصريف هذه المياه عبر شبكة الصرف الصحي العامة المعتمدة في المدينة، بما يضمن عدم وصولها إلى التربة أو المياه الجوفية، مع الالتزام بالمعايير الفنية للربط بالشبكة العامة."
hazardous_waste_default = "تُصنف المخلفات الخطرة الناتجة عن أنشطة المشروع ضمن الزيوت المستعملة الناتجة عن أعمال الصيانة الدورية للمركبات والمعدات، بالإضافة إلى الفلاتر المستهلكة وعبوات المواد الكيميائية الفارغة. وتلتزم المنشأة بالتعاقد مع جهة مختصة مرخصة من قبل 'المركز الوطني لإدارة النفايات (موان)' لجمع ونقل ومعالجة هذه المخلفات، مع الاحتفاظ بسجلات دقيقة ووثائق النقل المعتمدة لضمان الامتثال للوائح المركز الوطني للرقابة على الالتزام البيئي."

solid_waste = st.text_area("المخلفات الصلبة:", value=solid_waste_default, height=90)
liquid_waste = st.text_area("المخلفات السائلة:", value=liquid_waste_default, height=90)
hazardous_waste = st.text_area("المخلفات الخطرة:", value=hazardous_waste_default, height=120)

# ------------------------------------------------
# 3. قسم الخريطة التفاعلية (للعرض فقط)
# ------------------------------------------------
st.markdown("---")
st.subheader("🗺️ خريطة الموقع ونطاق المشروع (500 متر)")

if lat != 0.0 and lon != 0.0:
    m_display = folium.Map(
        location=[lat, lon], 
        zoom_start=15,
        tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
        attr="Google Satellite Hybrid"
    )
    folium.Marker([lat, lon], tooltip="موقع المنشأة", icon=folium.Icon(color="blue", icon="info-sign")).add_to(m_display)
    folium.Circle(
        location=[lat, lon], 
        radius=500, 
        color="#ff4d4d", 
        weight=2,
        fill=True, 
        fill_color="#ff4d4d", 
        fill_opacity=0.25
    ).add_to(m_display)
    st_folium(m_display, width=725, height=450)
    
    st.success("✨ سيتم تصوير هذه الخريطة مع نطاق 500 متر بدقة وإدراجها في التقرير تلقائياً!")
else:
    st.info("قم بإدخال خط العرض وخط الطول في الأعلى لكي تظهر الخريطة هنا تلقائياً.")

# ------------------------------------------------
# 4. قسم المرفقات والمستندات الرسمية وصور الزيارة
# ------------------------------------------------
st.markdown("---")
st.subheader("📎 مرفقات المشروع والمستندات الرسمية")

cr_file = st.file_uploader("📄 السجل التجاري (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"], key="cr")
vat_file = st.file_uploader("📄 شهادة ضريبة القيمة المضافة (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"], key="vat")
address_file = st.file_uploader("📄 العنوان الوطني (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"], key="addr")

st.markdown("---")
st.subheader("📸 صور الزيارة الميدانية")
visit_photos = st.file_uploader("ارفع صور الزيارة المتعددة (لتوضع في الملاحق بدون أسماء):", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="visit")

# ------------------------------------------------
# 5. قسم إصدار التقرير النهائي (عبر القالب)
# ------------------------------------------------
st.markdown("---")
st.subheader("📄 إصدار التقرير النهائي")

if st.button("إنشاء مسودة التقرير (Word) 📝"):
    if not company_name:
        st.error("الرجاء التأكد من إدخال اسم الشركة.")
    elif not os.path.exists("template.docx"):
        st.error("لم يتم العثور على ملف القالب. الرجاء التأكد من وجود ملف باسم 'template.docx'.")
    else:
        project_folder = os.path.join("Projects", company_name)
        os.makedirs(project_folder, exist_ok=True)
        
        with st.spinner('⏳ جاري أتمتة الخريطة بنطاق 500م وتجهيز المستندات وصور الزيارة المتعددة... الرجاء الانتظار قليلاً'):
            doc_template = DocxTemplate("template.docx")
            
            # تصوير الخريطة في الخلفية بدقة ونصف قطر 500 متر مع ضمان الحفظ
            map_path = os.path.join(project_folder, "Map_Screenshot.png")
            if lat != 0.0 and lon != 0.0:
                try:
                    capture_map_screenshot(lat, lon, map_path)
                except Exception as e:
                    st.warning(f"حدث خطأ أثناء تصوير الخريطة: {e}")
            
            # معالجة وحفظ المستندات الأساسية كصور
            cr_img = convert_to_image(cr_file, project_folder, "cr") if cr_file else None
            vat_img = convert_to_image(vat_file, project_folder, "vat") if vat_file else None
            addr_img = convert_to_image(address_file, project_folder, "addr") if address_file else None
            
            # معالجة صور الزيارة الميدانية المتعددة
            visit_image_objects = []
            if visit_photos:
                for i, photo in enumerate(visit_photos):
                    photo_path = os.path.join(project_folder, f"visit_{i}.png")
                    with open(photo_path, "wb") as f:
                        f.write(photo.getbuffer())
                    visit_image_objects.append(photo_path)
            
            visit_images_inline = [InlineImage(doc_template, img, width=Inches(5.0)) for img in visit_image_objects]
            
            # ضبط كائن خريطة الموقع للربط السليم مع القالب
            map_inline_image = None
            if os.path.exists(map_path) and os.path.getsize(map_path) > 0:
                map_inline_image = InlineImage(doc_template, map_path, width=Inches(5.5))
            
            # ربط المتغيرات بالرموز البرمجية في قالب الـ Word
            context = {
                'company_name': company_name,
                'tax_id': tax_id,
                'city': city,
                'address': address,
                'project_objective': project_objective,
                'project_area': project_area,
                'lat': lat,
                'lon': lon,
                'activity_code': search_query,
                'activity_name': selected_activity_name,
                'activity_description': activity_description,
                'solid_waste': solid_waste,
                'liquid_waste': liquid_waste,
                'hazardous_waste': hazardous_waste,
                
                # إدراج الصور في مواقعها بدقة عبر InlineImage
                'cr_image': InlineImage(doc_template, cr_img, width=Inches(5.0)) if cr_img else "[لم يتم إرفاق السجل التجاري]",
                'vat_image': InlineImage(doc_template, vat_img, width=Inches(5.0)) if vat_img else "[لم يتم إرفاق شهادة الضريبة]",
                'address_image': InlineImage(doc_template, addr_img, width=Inches(5.0)) if addr_img else "[لم يتم إرفاق العنوان الوطني]",
                'map_image': map_inline_image if map_inline_image else "[لم يتم تصوير خريطة الموقع]",
                'visit_images': visit_images_inline
            }
            
            doc_template.render(context)
            final_io = io.BytesIO()
            doc_template.save(final_io)
            final_io.seek(0)
            
        st.success("🎉 تم التقاط الخريطة بنطاق 500م وتعبئة التقرير وإدراج كافة الصور في مواقعها بنجاح تام!")
        st.download_button(
            label="📥 تحميل التقرير المعبأ الآن (Word)",
            data=final_io,
            file_name=f"VEKR-EMP_Report_{company_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
