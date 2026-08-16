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

# ------------------------------------------------
# إعدادات واجهة الصفحة (العنوان والأيقونة والهوية)
# ------------------------------------------------
st.set_page_config(
    page_title="VEKR-EMP | منصة الخطط البيئية",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تصميم احترافي بألوان رسمية واضحة (مع منع تأثير الوضع المظلم)
page_bg_img = """
<style>
/* Force the background AND global text color to stay light-themed */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #f8f9fa, #e9ecef) !important;
    background-size: cover !important;
    color: #1e293b !important; 
}
[data-testid="stHeader"] {
    background: rgba(0,0,0,0) !important;
}

/* Force all text elements to stay dark, even if Dark Mode is toggled */
p, span, div, h1, h2, h3, h4, h5, h6, label {
    color: #1e293b !important;
}

.main-title {
    color: #1e293b !important; 
    font-weight: 800 !important;
    text-align: center !important;
    margin-bottom: 24px !important;
}

.stButton>button {
    background-color: #0f766e !important; 
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: bold !important;
    width: 100% !important;
    padding: 10px 0 !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important; 
}

.stButton>button:hover {
    background-color: #14b8a6 !important; 
    color: #ffffff !important;
    box-shadow: 0 8px 16px rgba(15, 118, 110, 0.2) !important; 
    transform: translateY(-3px) !important; 
}
</style>
"""
st.markdown(page_bg_img, unsafe_allow_html=True)

# ------------------------------------------------
# دالة تحويل المرفقات (تدعم تحويل جميع صفحات الـ PDF)
# ------------------------------------------------
def convert_to_images(uploaded_file, output_folder, base_name):
    file_ext = uploaded_file.name.split('.')[-1].lower()
    img_paths = []
    
    if file_ext == "pdf":
        doc_pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        for page_num in range(len(doc_pdf)):
            page = doc_pdf.load_page(page_num)
            pix = page.get_pixmap(dpi=150)
            img_output_path = os.path.join(output_folder, f"{base_name}_page_{page_num + 1}.png")
            pix.save(img_output_path)
            img_paths.append(img_output_path)
    else:
        img_output_path = os.path.join(output_folder, f"{base_name}.png")
        with open(img_output_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        img_paths.append(img_output_path)
        
    return img_paths

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
else:
    st.info("قم بإدخال خط العرض وخط الطول في الأعلى لتظهر الخريطة هنا.")

# ------------------------------------------------
# 4. قسم المرفقات والمستندات الرسمية وصور الخريطة والزيارة
# ------------------------------------------------
st.markdown("---")
st.subheader("📎 مرفقات المشروع والمستندات الرسمية")

cr_file = st.file_uploader("📄 السجل التجاري (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"], key="cr")
vat_file = st.file_uploader("📄 شهادة ضريبة القيمة المضافة (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"], key="vat")
address_file = st.file_uploader("📄 العنوان الوطني (PDF أو صورة)", type=["pdf", "png", "jpg", "jpeg"], key="addr")

st.markdown("---")
st.subheader("🗺️ صورة خريطة الموقع (نطاق 500 متر)")
map_file = st.file_uploader("ارفع صورة لقطة شاشة الخريطة:", type=["png", "jpg", "jpeg"], key="map_file")

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
        
        with st.spinner('⏳ جاري معالجة كافة صفحات المستندات والخريطة وصور الزيارة...'):
            doc_template = DocxTemplate("template.docx")
            
            # معالجة المستندات (وتحويل كافة الصفحات لصور إن وجدت متعددة الصفحات)
            cr_paths = convert_to_images(cr_file, project_folder, "cr") if cr_file else []
            vat_paths = convert_to_images(vat_file, project_folder, "vat") if vat_file else []
            addr_paths = convert_to_images(address_file, project_folder, "addr") if address_file else []
            map_paths = convert_to_images(map_file, project_folder, "map_shot") if map_file else []
            
            # إذا كان المستند يتكون من عدة صفحات (مثل عقد أو شهادة من صفحتين)، نعرضها كقائمة متسلسلة تلقائياً
            cr_images_inline = [InlineImage(doc_template, p, width=Inches(5.0)) for p in cr_paths]
            vat_images_inline = [InlineImage(doc_template, p, width=Inches(5.0)) for p in vat_paths]
            addr_images_inline = [InlineImage(doc_template, p, width=Inches(5.0)) for p in addr_paths]
            map_image_inline = InlineImage(doc_template, map_paths[0], width=Inches(5.5)) if map_paths else None
            
            # معالجة صور الزيارة الميدانية المتعددة
            visit_image_objects = []
            if visit_photos:
                for i, photo in enumerate(visit_photos):
                    photo_path = os.path.join(project_folder, f"visit_{i}.png")
                    with open(photo_path, "wb") as f:
                        f.write(photo.getbuffer())
                    visit_image_objects.append(photo_path)
            
            visit_images_inline = [InlineImage(doc_template, img, width=Inches(5.0)) for img in visit_image_objects]
            
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
                
                # إدراج الصور (تدعم تعدد الصفحات الآن)
                'cr_image': cr_images_inline[0] if cr_images_inline else "[لم يتم إرفاق السجل التجاري]",
                'vat_image': vat_images_inline[0] if vat_images_inline else "[لم يتم إرفاق شهادة الضريبة]",
                'address_image': addr_images_inline[0] if addr_images_inline else "[لم يتم إرفاق العنوان الوطني]",
                'map_image': map_image_inline if map_image_inline else "[لم يتم إرفاق خريطة الموقع]",
                'visit_images': visit_images_inline
            }
            
            doc_template.render(context)
            final_io = io.BytesIO()
            doc_template.save(final_io)
            final_io.seek(0)
            
        st.success("🎉 تم معالجة كافة صفحات المستندات وإدراج الصور في مواقعها بنجاح تام!")
        st.download_button(
            label="📥 تحميل التقرير المعبأ الآن (Word)",
            data=final_io,
            file_name=f"VEKR-EMP_Report_{company_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
