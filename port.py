import streamlit as st
import pandas as pd
import os
import zipfile
import base64
import io
from PIL import Image

# ----------------- Directories Setup -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_DIR = os.path.join(DATA_DIR, "excel")
PHOTOS_DIR = os.path.join(DATA_DIR, "photos")
OUTPUT_DIR = os.path.join(DATA_DIR, "output_html")

for d in [EXCEL_DIR, PHOTOS_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

SAVED_EXCEL_PATH = os.path.join(EXCEL_DIR, "master_data.xlsx")

st.set_page_config(page_title="Aditya Birla IC Portfolio Generator", layout="wide")

# ----------------- Helper Functions -----------------
def get_image_base64(roll_no):
    r_str = str(roll_no).strip()
    for ext in ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']:
        candidate = os.path.join(PHOTOS_DIR, f"{r_str}{ext}")
        if os.path.exists(candidate):
            with open(candidate, "rb") as img_file:
                b64 = base64.b64encode(img_file.read()).decode('utf-8')
                mime = "image/png" if ext.lower() == '.png' else "image/jpeg"
                return f"data:{mime};base64,{b64}"
    return ""

def clean_val(val, default=""):
    if pd.isna(val) or val is None:
        return default
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

# ----------------- Exact HTML Builder -----------------
def generate_student_html(row):
    # Mapping columns as per your format
    # A to V: Personal | W to AG: Attendance | AH to AM: Test Marks | AN to BQ: Activities
    roll_no = clean_val(row[0] if len(row) > 0 else "")
    cls_name = clean_val(row[1] if len(row) > 1 else "", "12-B")
    sr_no = clean_val(row[2] if len(row) > 2 else "")
    pen_no = clean_val(row[4] if len(row) > 4 else "")
    dob_dd = clean_val(row[6] if len(row) > 6 else "")
    dob_mm = clean_val(row[7] if len(row) > 7 else "")
    dob_yyyy = clean_val(row[8] if len(row) > 8 else "")
    dob = f"{dob_yyyy}-{dob_mm.zfill(2)}-{dob_dd.zfill(2)}" if dob_yyyy and dob_mm else "—"
    
    st_name = clean_val(row[9] if len(row) > 9 else "")
    f_name = clean_val(row[10] if len(row) > 10 else "")
    m_name = clean_val(row[11] if len(row) > 11 else "")
    mobile = clean_val(row[17] if len(row) > 17 else "")
    address = clean_val(row[16] if len(row) > 16 else "")

    short_term_goal = clean_val(row[39] if len(row) > 39 else "", "0.85")
    long_term_goal = clean_val(row[40] if len(row) > 40 else "", "CUET")
    reflection = clean_val(row[69] if len(row) > 69 else "", "ताकत: परिश्रम व अनुशासन | सुधार क्षेत्र: समय प्रबंधन।")

    img_b64 = get_image_base64(roll_no)
    img_tag = f'<img src="{img_b64}" style="width: 95px; height: 115px; object-fit: cover; border-radius: 6px; border: 2px solid #1E3A8A;"/>' if img_b64 else '<div style="width: 95px; height: 115px; border-radius: 6px; border: 2px dashed #94A3B8; display:flex; align-items:center; justify-content:center; color:#64748B; font-size:11px;">फ़ोटो उपलब्ध नहीं</div>'

    # Activities List 1 to 8 (As in reference sample)
    activities_meta = [
        ("27.08.2026", "Tata Building India School Essay Competition", "साहित्यिक (निबंध)", 41, 42),
        ("27.08.2026", "रंगोली प्रतियोगिता", "कला एवं संस्कृति", 43, 44),
        ("27.08.2026", "मेहंदी प्रतियोगिता", "कला एवं संस्कृति", 45, 46),
        ("20.08.2026", "राखी निर्माण प्रतियोगिता", "क्राफ्ट एवं रचनात्मक कौशल", 47, 48),
        ("13.08.2026", "चित्रकला प्रतियोगिता", "दृश्य कला (Drawing)", 49, 50),
        ("06.08.2026", "निबंध प्रतियोगिता", "साहित्यिक (निबंध)", 51, 52),
        ("30.07.2026", "कक्षा सज्जा एवं शैक्षणिक चार्ट प्रतियोगिता", "रचनात्मक एवं शैक्षणिक कौशल", 53, 54),
        ("02.07.2026", "लेख प्रतियोगिता (Article Writing)", "सामाजिक जागरूकता / वैचारिक लेखन", 55, 56),
    ]

    act_rows_html = ""
    for date_str, title, cat, p_idx, r_idx in activities_meta:
        sub_text = clean_val(row[p_idx] if len(row) > p_idx else "", "प्रतिभाग किया")
        learning_text = clean_val(row[r_idx] if len(row) > r_idx else "", sub_text)
        
        act_rows_html += f"""
            <tr style="border-bottom: 1px solid #E2E8F0; font-size: 12px;">
                <td style="padding: 7px; text-align: center;">{date_str}</td>
                <td style="padding: 7px; font-weight: 600; color: #1E3A8A;">{title}<br><span style="font-weight: normal; color: #475569; font-size: 11px;">{sub_text}</span></td>
                <td style="padding: 7px; text-align: center;">{cat}</td>
                <td style="padding: 7px; color: #0284C7; font-style: italic;">{learning_text}</td>
                <td style="padding: 7px; text-align: center; font-weight: bold; color: #059669;">5/5</td>
            </tr>
        """

    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Portfolio - {st_name}</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f8fafc; padding: 15px; color: #1e293b; }}
        .page {{ max-width: 850px; margin: 0 auto 25px auto; background: #ffffff; border: 2px solid #1E3A8A; border-radius: 10px; padding: 25px; box-shadow: 0 4px 10px rgba(0,0,0,0.06); }}
        @media print {{
            body {{ background: none; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; border: 2px solid #000; page-break-after: always; }}
        }}
    </style>
</head>
<body>
    <!-- ================= PAGE 1 ================= -->
    <div class="page">
        <div style="text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 12px; margin-bottom: 18px;">
            <h2 style="margin: 0; color: #1E3A8A; font-size: 22px; text-transform: uppercase; letter-spacing: 1px;">ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT, SONEBHADRA (UP)</h2>
            <h3 style="margin: 4px 0 0 0; color: #059669; font-size: 17px;">छात्र पोर्टफोलियो एवं सतत आंतरिक मूल्यांकन रिकॉर्ड</h3>
            <div style="font-size: 13px; color: #475569; margin-top: 4px;">सत्र: 2026 - 2027 | कक्षा: {cls_name}</div>
            <div style="display: inline-block; background: #1E3A8A; color: white; padding: 3px 14px; border-radius: 12px; font-size: 11px; margin-top: 6px; font-weight: 600;">भाग 1 : व्यक्तिगत विवरण एवं स्व-मूल्यांकन</div>
        </div>

        <div style="display: flex; gap: 15px; margin-bottom: 20px;">
            <table style="width: 72%; border-collapse: collapse; font-size: 13px;">
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold; width: 35%;">छात्र/छात्रा का नाम:</td><td style="padding: 6px; color: #1E3A8A; font-weight: bold; font-size: 14px;">{st_name}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">अनुक्रमांक (Roll No.):</td><td style="padding: 6px; font-weight: bold;">{roll_no}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold;">S.R. No. / PEN:</td><td style="padding: 6px;">{sr_no} / {pen_no}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">पिता का नाम:</td><td style="padding: 6px;">{f_name}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold;">माता का नाम:</td><td style="padding: 6px;">{m_name}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">जन्म तिथि (D.O.B.):</td><td style="padding: 6px;">{dob}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold;">संपर्क सूत्र (Mobile):</td><td style="padding: 6px;">{mobile}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">निवास पता:</td><td style="padding: 6px;">{address}</td></tr>
            </table>
            <div style="width: 28%; border: 2px dashed #94A3B8; border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; background: #F8FAFC; padding: 10px; text-align: center;">
                {img_tag}
                <div style="font-weight: bold; font-size: 13px; color: #1E3A8A; margin-top: 6px;">{st_name}</div>
                <div style="font-size: 11px; color: #64748B;">कक्षा: {cls_name}</div>
                <div style="font-size: 10px; color: #059669; margin-top: 4px; border: 1px solid #059669; padding: 2px 6px; border-radius: 8px;">सत्यापित विद्यार्थी</div>
            </div>
        </div>

        <div style="margin-top: 15px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 14px; margin-bottom: 6px;">🎯 शैक्षणिक लक्ष्य एवं संकल्प (Academic Vision & Career Goals):</div>
            <div style="display: flex; gap: 12px; margin-top: 5px;">
                <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #3B82F6; padding: 8px 12px; border-radius: 4px; font-size: 12.5px; color: #1e293b;">
                    <strong style="color: #1E3A8A;">📌 अल्पकालिक लक्ष्य (Short-Term Goal 2026-27):</strong><br>
                    {short_term_goal}
                </div>
                <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #059669; padding: 8px 12px; border-radius: 4px; font-size: 12.5px; color: #1e293b;">
                    <strong style="color: #059669;">🎯 दीर्घकालिक लक्ष्य (Long-Term Goal - Career):</strong><br>
                    {long_term_goal}
                </div>
            </div>
        </div>

        <div style="margin-top: 15px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 14px; margin-bottom: 6px;">💡 क्षमताएं एवं सुधार क्षेत्र (Self-Reflection):</div>
            <div style="background: #F8FAFC; border-left: 4px solid #10B981; padding: 10px 14px; border-radius: 4px; font-size: 13px; color: #334155; line-height: 1.5;">{reflection}</div>
        </div>
    </div>

    <!-- ================= PAGE 2 ================= -->
    <div class="page">
        <div style="text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 12px; margin-bottom: 15px;">
            <h3 style="margin: 0; color: #1E3A8A; font-size: 19px; text-transform: uppercase;">सह-पाठ्यचर्या एवं गतिविधि मूल्यांकन प्रपत्र</h3>
            <div style="display: inline-block; background: #059669; color: white; padding: 3px 14px; border-radius: 12px; font-size: 11px; margin-top: 6px; font-weight: 600;">भाग 2 : गतिविधि विवरण, छात्र चिंतन एवं रूब्रिक्स</div>
        </div>

        <div style="margin-bottom: 15px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 13px; margin-bottom: 8px;">📋 सत्र 2026-27 में संपादित प्रमुख गतिविधियां एवं प्रतियोगिताएं:</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 12px; border: 1px solid #CBD5E1;">
                <thead>
                    <tr style="background: #1E3A8A; color: white; text-align: left;">
                        <th style="padding: 7px; width: 12%; text-align: center;">तिथि</th>
                        <th style="padding: 7px; width: 38%;">गतिविधि / प्रतियोगिता का नाम</th>
                        <th style="padding: 7px; width: 18%; text-align: center;">श्रेणी</th>
                        <th style="padding: 7px; width: 22%;">विद्यार्थी की सीख / प्रस्तुति</th>
                        <th style="padding: 7px; width: 10%; text-align: center;">अंक</th>
                    </tr>
                </thead>
                <tbody>
                    {act_rows_html}
                </tbody>
            </table>
        </div>

        <div style="border: 1px solid #CBD5E1; border-radius: 6px; padding: 12px; background: #F8FAFC; margin-top: 20px;">
            <div style="margin: 0 0 8px 0; color: #1E3A8A; font-weight: bold; font-size: 13px;">📝 आंतरिक मूल्यांकन रूब्रिक्स (UP Board Marking Criteria - पूर्णांक: 20)</div>
            <div style="display: flex; gap: 8px; font-size: 12px; text-align: center;">
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>1. नियमितता व सहभागिता</strong><br>(5 अंक)</div>
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>2. मौलिकता व शुद्धता</strong><br>(5 अंक)</div>
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>3. रचनात्मकता व कौशल</strong><br>(5 अंक)</div>
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>4. प्रस्तुतिकरण व आचरण</strong><br>(5 अंक)</div>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 30px; padding-top: 10px; border-top: 1px dashed #94A3B8; font-size: 12px;">
                <div>
                    <div><strong>विद्यार्थी के हस्ताक्षर:</strong> _____________________</div>
                    <div style="color: #64748B; font-size: 11px; margin-top: 4px;">दिनांक: 04-09-2026</div>
                </div>
                <div style="text-align: right;">
                    <div><strong>कक्षा अध्यापक / प्रभारी हस्ताक्षर:</strong> _____________________</div>
                    <div style="color: #64748B; font-size: 11px; margin-top: 4px;">कक्षा अध्यापक ({cls_name})</div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""
    return html_content

# ----------------- Streamlit UI -----------------
st.title("🎓 UP Board Portfolio Generator - HTML Format")
st.caption("Aditya Birla Intermediate College | Bulk JPEG Upload & Auto-Matched Portfolios")

# Sidebar
st.sidebar.header("📁 फ़ाइल प्रबंधन (File Management)")

# 1. Master Excel Upload
excel_file = st.sidebar.file_uploader("1. मास्टर एक्सेल शीट अपलोड करें (.xlsx)", type=["xlsx"])
if excel_file:
    with open(SAVED_EXCEL_PATH, "wb") as f:
        f.write(excel_file.getbuffer())
    st.sidebar.success("✅ मास्टर एक्सेल फ़ाइल स्थायी रूप से सेव हो गई!")

# 2. Bulk Photo Upload (Multiple JPEG / PNG Selection)
st.sidebar.subheader("📸 बल्क फ़ोटो अपलोड (Multiple Photos)")
multi_photos = st.sidebar.file_uploader(
    "एक साथ सभी छात्र फ़ोटो चुनें (JPEG/JPG/PNG)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True,
    help="फ़ोटो का नाम रोल नंबर के आधार पर होना चाहिए (जैसे: 39.jpg, 40.jpeg आदि)"
)

if multi_photos:
    saved_count = 0
    for p in multi_photos:
        filename = os.path.basename(p.name)
        save_dest = os.path.join(PHOTOS_DIR, filename)
        with open(save_dest, "wb") as f:
            f.write(p.getbuffer())
        saved_count += 1
    st.sidebar.success(f"✅ {saved_count} फ़ोटो सफलतापूर्वक अपलोड और लिंक हो गईं!")

# Check if data exists
if not os.path.exists(SAVED_EXCEL_PATH):
    st.info("👈 कृपया बाएँ साइडबार से अपनी मास्टर एक्सेल शीट अपलोड करें।")
    st.stop()

@st.cache_data
def load_data(path):
    return pd.read_excel(path, header=0)

df = load_data(SAVED_EXCEL_PATH)

tab1, tab2, tab3 = st.tabs(["👤 व्यक्तिगत छात्र पोर्टफोलियो", "📦 पूरी क्लास का बल्क डाउनलोड (ZIP)", "📷 एक फ़ोटो बदलें/अपलोड करें"])

# ----------------- TAB 1: Single Student -----------------
with tab1:
    st.subheader("छात्र का चयन करें एवं HTML पोर्टफोलियो डाउनलोड करें")
    roll_col = df.columns[0]
    name_col = df.columns[9] if len(df.columns) > 9 else df.columns[1]

    student_options = {f"Roll: {clean_val(row[roll_col])} - {clean_val(row[name_col])}": idx for idx, row in df.iterrows()}
    selected_label = st.selectbox("विद्यार्थी चुनें:", list(student_options.keys()))

    if selected_label:
        s_idx = student_options[selected_label]
        student_row = df.iloc[s_idx].tolist()
        roll_no = clean_val(student_row[0])
        student_name = clean_val(student_row[9])
        
        c1, c2 = st.columns([3, 1])
        with c1:
            st.write(f"**नाम:** {student_name}")
            st.write(f"**रोल नंबर:** {roll_no}")
            st.write(f"**कक्षा:** {clean_val(student_row[1], '12-B')}")
        with c2:
            b64_img = get_image_base64(roll_no)
            if b64_img:
                st.markdown(f'<img src="{b64_img}" width="100" style="border-radius:6px; border:1px solid #1E3A8A;"/>', unsafe_allow_html=True)
            else:
                st.warning("⚠️ फ़ोटो अपलोड नहीं है (नाम: {0}.jpg होना चाहिए)".format(roll_no))

        # Generate HTML
        html_code = generate_student_html(student_row)
        file_name = f"UPBoard_12B_Roll_{roll_no}_{student_name}.html"

        st.download_button(
            label=f"⬇️ डाउनलोड पोर्टफोलियो ({file_name})",
            data=html_code,
            file_name=file_name,
            mime="text/html"
        )

        with st.expander("👁️ पोर्टफोलियो का लाइव प्रीव्यू देखें"):
            st.components.v1.html(html_code, height=900, scrolling=True)

# ----------------- TAB 2: Bulk Generation (ZIP) -----------------
with tab2:
    st.subheader("पूरी कक्षा के सभी पोर्टफोलियो एक क्लिक में डाउनलोड करें")
    st.write(f"शीट में कुल विद्यार्थी: **{len(df)}**")

    if st.button("🚀 सभी छात्रों के HTML पोर्टफोलियो जनरेट करें (ZIP Archive)"):
        progress_bar = st.progress(0)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for idx, row in df.iterrows():
                s_list = row.tolist()
                r_no = clean_val(s_list[0])
                s_nm = clean_val(s_list[9])
                s_html = generate_student_html(s_list)
                arc_name = f"UPBoard_12B_Roll_{r_no}_{s_nm}.html"
                zip_file.writestr(arc_name, s_html)
                progress_bar.progress((idx + 1) / len(df))

        st.success("✅ सभी छात्रों के HTML पोर्टफोलियो तैयार हैं!")
        st.download_button(
            label="⬇️ सभी पोर्टफोलियो डाउनलोड करें (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="All_Students_UPBoard_Portfolios.zip",
            mime="application/zip"
        )

# ----------------- TAB 3: Single Photo Upload -----------------
with tab3:
    st.subheader("किसी एक छात्र की फ़ोटो जोड़ें या बदलें")
    s_roll = st.text_input("छात्र का रोल नंबर दर्ज करें (जैसे: 39):")
    s_file = st.file_uploader("फ़ोटो फ़ाइल चुनें (JPEG/PNG):", type=["jpg", "jpeg", "png"], key="tab3_photo")

    if s_roll and s_file:
        if st.button("💾 फ़ोटो सेव करें"):
            ext = os.path.splitext(s_file.name)[1].lower()
            save_path = os.path.join(PHOTOS_DIR, f"{s_roll.strip()}{ext}")
            with open(save_path, "wb") as f:
                f.write(s_file.getbuffer())
            st.success(f"✅ रोल नंबर {s_roll} के लिए फ़ोटो सफलतापूर्वक अपडेट हो गई!")
