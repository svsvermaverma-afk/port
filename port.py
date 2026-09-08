import streamlit as st
import pandas as pd
import os
import zipfile
import base64
import io
import re
import requests
from PIL import Image

# ----------------- Directories Setup -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_DIR = os.path.join(DATA_DIR, "excel")
PHOTOS_DIR = os.path.join(DATA_DIR, "photos")
OUTPUT_DIR = os.path.join(DATA_DIR, "output_html")

for d in [EXCEL_DIR, PHOTOS_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

SAVED_EXCEL_PATH = os.path.join(EXCEL_DIR, "all_school_master_data.xlsx")
SAVED_FORM_DATA_PATH = os.path.join(EXCEL_DIR, "form_sync_data.csv")

st.set_page_config(page_title="Complete Student Portfolio Generator", layout="wide")

# ----------------- Helper Functions -----------------
def clean_val(val, default="—"):
    if pd.isna(val) or val is None:
        return default
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str if val_str != "" else default

@st.cache_data(show_spinner=False)
def load_image_base64_from_drive(drive_url):
    """Google Form file upload link ko bina gdown ke Base64 me convert karta hai."""
    try:
        file_id = None
        id_match = re.search(r'id=([a-zA-Z0-9_-]+)', str(drive_url))
        if id_match:
            file_id = id_match.group(1)
        else:
            d_match = re.search(r'/d/([a-zA-Z0-9_-]+)', str(drive_url))
            if d_match:
                file_id = d_match.group(1)

        if not file_id:
            return ""

        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        session = requests.Session()
        response = session.get(download_url, timeout=10)

        # Google Drive confirmation token check
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                response = session.get(f"{download_url}&confirm={value}", timeout=10)
                break

        if response.status_code == 200 and len(response.content) > 200:
            pil_img = Image.open(io.BytesIO(response.content))
            output_buffer = io.BytesIO()
            img_format = pil_img.format if pil_img.format else "JPEG"
            pil_img.save(output_buffer, format=img_format)
            b64 = base64.b64encode(output_buffer.getvalue()).decode('utf-8')
            return f"data:image/{img_format.lower()};base64,{b64}"
    except Exception:
        pass
    return ""

def get_image_base64(roll_no, drive_photo_url=None):
    if drive_photo_url and str(drive_photo_url).strip() not in ["", "—", "nan"]:
        b64_from_drive = load_image_base64_from_drive(drive_photo_url)
        if b64_from_drive:
            return b64_from_drive

    r_str = str(roll_no).strip()
    for ext in ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']:
        candidate = os.path.join(PHOTOS_DIR, f"{r_str}{ext}")
        if os.path.exists(candidate):
            try:
                with open(candidate, "rb") as img_file:
                    b64 = base64.b64encode(img_file.read()).decode('utf-8')
                    mime = "image/png" if ext.lower() == '.png' else "image/jpeg"
                    return f"data:{mime};base64,{b64}"
            except Exception:
                pass
    return ""

def remove_existing_photos(roll_no):
    r_str = str(roll_no).strip()
    for ext in ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']:
        p = os.path.join(PHOTOS_DIR, f"{r_str}{ext}")
        if os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass

# ----------------- 3-Page HTML Builder -----------------
def generate_full_portfolio_html(row_dict):
    sr_no       = clean_val(row_dict.get('SerialNo'))
    adm_no      = clean_val(row_dict.get('AdmNo'))
    st_name     = clean_val(row_dict.get('Name'))
    cls_name    = clean_val(row_dict.get('Class'))
    dob         = clean_val(row_dict.get('DateOfBirth'))
    religion    = clean_val(row_dict.get('Religion'))
    house       = clean_val(row_dict.get('House'))
    email       = clean_val(row_dict.get('EmailId'))
    roll_no     = clean_val(row_dict.get('RollNo'))
    caste       = clean_val(row_dict.get('Caste'))
    gender      = clean_val(row_dict.get('Gender'))
    raw_aadhar  = clean_val(row_dict.get('AadharCard'))
    aadhar_no   = f"XXXX-XXXX-{raw_aadhar[-4:]}" if (raw_aadhar != "—" and len(raw_aadhar) >= 4) else "[Aadhaar Redacted]"
    mobile      = clean_val(row_dict.get('CommunicationNo'))
    f_name      = clean_val(row_dict.get('FatherName'))
    m_name      = clean_val(row_dict.get('MotherName'))
    pen_no      = clean_val(row_dict.get('PENNo'))
    address     = clean_val(row_dict.get('PresentAddress'))

    short_goal  = clean_val(row_dict.get('ShortGoal'), "शैक्षणिक विषयों में दक्षता प्राप्त करना एवं उत्कृष्ट प्रदर्शन।")
    long_goal   = clean_val(row_dict.get('LongGoal'), "उच्च शिक्षा एवं प्रतिष्ठित करियर निर्माण।")
    reflection  = clean_val(row_dict.get('Reflection'), "नियमित अभ्यास, अनुशासन एवं समय प्रबंधन पर विशेष ध्यान।")
    drive_photo = row_dict.get('PhotoDriveLink', None)

    img_b64 = get_image_base64(roll_no, drive_photo)
    img_tag = f'<img src="{img_b64}" style="width: 105px; height: 130px; object-fit: cover; border-radius: 6px; border: 2px solid #1E3A8A;"/>' if img_b64 else '<div style="width: 105px; height: 130px; border-radius: 6px; border: 2px dashed #94A3B8; display:flex; align-items:center; justify-content:center; color:#64748B; font-size:11px; text-align:center; padding:5px;">फ़ोटो उपलब्ध नहीं</div>'

    att_months = ["अप्रैल", "मई", "जुलाई", "अगस्त", "सितंबर", "अक्टूबर", "नवंबर", "दिसंबर", "जनवरी", "कुल उपस्थिति", "प्रतिशत (%)"]
    att_headers_html = "".join([f'<th style="padding: 6px; border: 1px solid #CBD5E1; text-align: center;">{m}</th>' for m in att_months])
    att_values_html = "".join([f'<td style="padding: 6px; border: 1px solid #CBD5E1; text-align: center; font-weight: 600;">{clean_val(row_dict.get(m, "—"))}</td>' for m in att_months])

    act_part1 = [
        ("27.08.2026", "1. Tata Building India School Essay Competition", "साहित्यिक (निबंध)", "Act1_Remark"),
        ("27.08.2026", "2. रंगोली प्रतियोगिता (Rangoli Making)", "कला एवं संस्कृति", "Act2_Remark"),
        ("27.08.2026", "3. मेहंदी प्रतियोगिता (Mehndi Design)", "कला एवं संस्कृति", "Act3_Remark"),
        ("20.08.2026", "4. राखी निर्माण प्रतियोगिता (Rakhi Making)", "क्राफ्ट एवं रचनात्मकता", "Act4_Remark"),
        ("13.08.2026", "5. चित्रकला प्रतियोगिता (Drawing)", "दृश्य कला (Fine Arts)", "Act5_Remark"),
        ("06.08.2026", "6. निबंध प्रतियोगिता (Essay Writing)", "साहित्यिक कौशल", "Act6_Remark"),
        ("30.07.2026", "7. बाल संसद गतिविधियां (Bal Sansad)", "नेतृत्व एवं सामाजिक कौशल", "Act7_Remark"),
    ]
    act_part1_rows = ""
    for dt, title, cat_item, key_rem in act_part1:
        rem_text = clean_val(row_dict.get(key_rem), "सक्रिय प्रतिभाग एवं सराहनीय प्रदर्शन")
        act_part1_rows += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{dt}</td>
                <td style="padding: 6px; font-weight: 600; color: #1E3A8A; border: 1px solid #CBD5E1;">{title}<br><span style="font-weight: normal; color: #475569; font-size: 11px;">सक्रिय सहभागिता</span></td>
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{cat_item}</td>
                <td style="padding: 6px; color: #0284C7; font-style: italic; border: 1px solid #CBD5E1;">{rem_text}</td>
                <td style="padding: 6px; text-align: center; font-weight: bold; color: #059669; border: 1px solid #CBD5E1;">5/5</td>
            </tr>
        """

    act_part2 = [
        ("24.07.2026", "8. कक्षा सज्जा एवं चार्ट (Class Decoration)", "रचनात्मक एवं नवाचार", "Act8_Remark"),
        ("16.07.2026", "9. भाषण प्रतियोगिता (Speech/Elocution)", "वाक कौशल व आत्मविश्वास", "Act9_Remark"),
        ("09.07.2026", "10. कहानी लेखन (Story Writing)", "साहित्यिक सृजन", "Act10_Remark"),
        ("02.07.2026", "11. आई.ई.पी. पोर्टफोलियो (IEP Portfolio)", "शैक्षणिक पोर्टफोलियो कार्य", "Act11_Remark"),
        ("25.04.2026", "12. समूह चर्चा (Group Discussion)", "संवाद एवं संप्रेषण कौशल", "Act12_Remark"),
        ("18.04.2026", "13. लेख प्रतियोगिता (Article Writing)", "वैचारिक एवं सामाजिक लेखन", "Act13_Remark"),
        ("10.04.2026", "14. मौलिक रचना (Creative Story/Writing)", "मौलिक रचनात्मकता", "Act14_Remark"),
    ]
    act_part2_rows = ""
    for dt, title, cat_item, key_rem in act_part2:
        rem_text = clean_val(row_dict.get(key_rem), "सक्रिय प्रतिभाग एवं सराहनीय प्रदर्शन")
        act_part2_rows += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{dt}</td>
                <td style="padding: 6px; font-weight: 600; color: #1E3A8A; border: 1px solid #CBD5E1;">{title}<br><span style="font-weight: normal; color: #475569; font-size: 11px;">सक्रिय सहभागिता</span></td>
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{cat_item}</td>
                <td style="padding: 6px; color: #0284C7; font-style: italic; border: 1px solid #CBD5E1;">{rem_text}</td>
                <td style="padding: 6px; text-align: center; font-weight: bold; color: #059669; border: 1px solid #CBD5E1;">5/5</td>
            </tr>
        """

    html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Portfolio - {st_name} (Roll: {roll_no})</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f8fafc; padding: 10px; color: #1e293b; line-height: 1.35; margin: 0; }}
        .page {{ max-width: 850px; margin: 0 auto 25px auto; background: #ffffff; border: 2px solid #1E3A8A; border-radius: 8px; padding: 22px; box-shadow: 0 4px 10px rgba(0,0,0,0.06); }}
        .header {{ text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 10px; margin-bottom: 14px; }}
        .sec-title {{ color: #1E3A8A; font-weight: bold; font-size: 13.5px; margin-top: 14px; margin-bottom: 6px; display: flex; align-items: center; gap: 6px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        @media print {{
            body {{ background: none; padding: 0; }}
            .page {{ 
                box-shadow: none; 
                margin: 0; 
                border: 2px solid #000; 
                page-break-after: always; 
                break-after: page;
                min-height: 100vh;
            }}
            .page:last-child {{
                page-break-after: avoid;
                break-after: avoid;
            }}
        }}
    </style>
</head>
<body>

    <!-- PAGE 1 -->
    <div class="page">
        <div class="header">
            <h2 style="margin: 0; color: #1E3A8A; font-size: 21px; text-transform: uppercase; letter-spacing: 0.5px;">ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT, SONEBHADRA (UP)</h2>
            <h3 style="margin: 3px 0 0 0; color: #059669; font-size: 16px;">छात्र संपूर्ण पोर्टफोलियो एवं सतत आंतरिक मूल्यांकन रिकॉर्ड</h3>
            <div style="font-size: 12.5px; color: #475569; margin-top: 3px;">सत्र: 2026 - 2027 | कक्षा: {cls_name}</div>
            <div style="display: inline-block; background: #1E3A8A; color: white; padding: 2px 12px; border-radius: 10px; font-size: 11px; margin-top: 5px; font-weight: 600;">भाग 1 : व्यक्तिगत विवरण एवं पहचान अभिलेख</div>
        </div>

        <div style="display: flex; gap: 15px; margin-bottom: 10px;">
            <table style="width: 74%; border: 1px solid #CBD5E1;">
                <tr style="background: #F1F5F9;">
                    <td style="padding: 5px; font-weight: bold; width: 28%; border: 1px solid #CBD5E1;">छात्र का नाम:</td>
                    <td style="padding: 5px; color: #1E3A8A; font-weight: bold; font-size: 13px; border: 1px solid #CBD5E1;">{st_name}</td>
                    <td style="padding: 5px; font-weight: bold; width: 22%; border: 1px solid #CBD5E1;">अनुक्रमांक (Roll No.):</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">{roll_no}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">प्रवेश सं० (Adm No.):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{adm_no}</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">क्र० सं० (Serial No.):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{sr_no}</td>
                </tr>
                <tr style="background: #F1F5F9;">
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">कक्षा (Class):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{cls_name}</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">सदन (House):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{house}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">PEN Number:</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{pen_no}</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">जन्म तिथि (D.O.B.):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{dob}</td>
                </tr>
                <tr style="background: #F1F5F9;">
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">पिता का नाम:</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{f_name}</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">माता का नाम:</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{m_name}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">लिंग / धर्म:</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{gender} / {religion}</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">जाति (Caste):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{caste}</td>
                </tr>
                <tr style="background: #F1F5F9;">
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">संपर्क सूत्र (Mobile):</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{mobile}</td>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">आधार संख्या:</td>
                    <td style="padding: 5px; border: 1px solid #CBD5E1;">{aadhar_no}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">ईमेल (Email):</td>
                    <td colspan="3" style="padding: 5px; border: 1px solid #CBD5E1; word-break: break-all;">{email}</td>
                </tr>
                <tr style="background: #F1F5F9;">
                    <td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">वर्तमान पता (Address):</td>
                    <td colspan="3" style="padding: 5px; border: 1px solid #CBD5E1;">{address}</td>
                </tr>
            </table>
            
            <div style="width: 26%; border: 2px dashed #94A3B8; border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; background: #F8FAFC; padding: 8px; text-align: center;">
                {img_tag}
                <div style="font-weight: bold; font-size: 12.5px; color: #1E3A8A; margin-top: 6px;">{st_name}</div>
                <div style="font-size: 11px; color: #64748B;">रोल: {roll_no} | कक्षा: {cls_name}</div>
                <div style="font-size: 10px; color: #059669; margin-top: 4px; border: 1px solid #059669; padding: 2px 6px; border-radius: 8px;">सत्यापित विद्यार्थी</div>
            </div>
        </div>

        <div class="sec-title">🎯 शैक्षणिक लक्ष्य एवं संकल्प (Academic Vision & Goals):</div>
        <div style="display: flex; gap: 10px;">
            <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #3B82F6; padding: 7px 10px; border-radius: 4px; font-size: 12px;">
                <strong style="color: #1E3A8A;">📌 अल्पकालिक लक्ष्य (Short-Term Goal):</strong><br>
                {short_goal}
            </div>
            <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #059669; padding: 7px 10px; border-radius: 4px; font-size: 12px;">
                <strong style="color: #059669;">🎯 दीर्घकालिक लक्ष्य (Long-Term Goal):</strong><br>
                {long_goal}
            </div>
        </div>

        <div class="sec-title">💡 स्व-चिंतन एवं सुधार क्षेत्र (Self-Reflection):</div>
        <div style="background: #F8FAFC; border-left: 4px solid #10B981; padding: 8px 12px; border-radius: 4px; font-size: 12px; color: #334155;">
            {reflection}
        </div>
    </div>

    <!-- PAGE 2 -->
    <div class="page">
        <div class="header">
            <h3 style="margin: 0; color: #1E3A8A; font-size: 18px; text-transform: uppercase;">सह-पाठ्यचर्या, उपस्थिति एवं गतिविधि मूल्यांकन प्रपत्र (भाग - 1)</h3>
            <div style="font-size: 12px; color: #475569; margin-top: 3px;">विद्यार्थी: <strong>{st_name}</strong> | अनुक्रमांक: <strong>{roll_no}</strong> | कक्षा: <strong>{cls_name}</strong></div>
            <div style="display: inline-block; background: #059669; color: white; padding: 2px 12px; border-radius: 10px; font-size: 11px; margin-top: 5px; font-weight: 600;">भाग 2 : माह-वार उपस्थिति एवं गतिविधियां (1 से 7)</div>
        </div>

        <div class="sec-title">📅 माह-वार उपस्थिति विवरण (Monthly Attendance Progression):</div>
        <table style="border: 1px solid #CBD5E1; margin-bottom: 15px;">
            <tr style="background: #1E3A8A; color: white;">
                {att_headers_html}
            </tr>
            <tr style="background: #F8FAFC;">
                {att_values_html}
            </tr>
        </table>

        <div class="sec-title">📋 पाठ्य-सहगामी गतिविधियां एवं प्रदर्शन रिकॉर्ड (Activities 1 to 7):</div>
        <table style="border: 1px solid #CBD5E1;">
            <thead>
                <tr style="background: #1E3A8A; color: white; text-align: left;">
                    <th style="padding: 6px; width: 12%; text-align: center; border: 1px solid #CBD5E1;">तिथि</th>
                    <th style="padding: 6px; width: 38%; border: 1px solid #CBD5E1;">गतिविधि / प्रतियोगिता का नाम</th>
                    <th style="padding: 6px; width: 18%; text-align: center; border: 1px solid #CBD5E1;">श्रेणी</th>
                    <th style="padding: 6px; width: 22%; border: 1px solid #CBD5E1;">विद्यार्थी की सीख / प्रस्तुति</th>
                    <th style="padding: 6px; width: 10%; text-align: center; border: 1px solid #CBD5E1;">अंक</th>
                </tr>
            </thead>
            <tbody>
                {act_part1_rows}
            </tbody>
        </table>
    </div>

    <!-- PAGE 3 -->
    <div class="page">
        <div class="header">
            <h3 style="margin: 0; color: #1E3A8A; font-size: 18px; text-transform: uppercase;">सह-पाठ्यचर्या, गतिविधि मूल्यांकन एवं सत्यापन (भाग - 2)</h3>
            <div style="font-size: 12px; color: #475569; margin-top: 3px;">विद्यार्थी: <strong>{st_name}</strong> | अनुक्रमांक: <strong>{roll_no}</strong> | कक्षा: <strong>{cls_name}</strong></div>
            <div style="display: inline-block; background: #D97706; color: white; padding: 2px 12px; border-radius: 10px; font-size: 11px; margin-top: 5px; font-weight: 600;">भाग 3 : गतिविधियां (8 से 14) व आंतरिक मूल्यांकन सत्यापन</div>
        </div>

        <div class="sec-title">📋 पाठ्य-सहगामी गतिविधियां एवं प्रदर्शन रिकॉर्ड (Activities 8 to 14):</div>
        <table style="border: 1px solid #CBD5E1; margin-bottom: 15px;">
            <thead>
                <tr style="background: #1E3A8A; color: white; text-align: left;">
                    <th style="padding: 6px; width: 12%; text-align: center; border: 1px solid #CBD5E1;">तिथि</th>
                    <th style="padding: 6px; width: 38%; border: 1px solid #CBD5E1;">गतिविधि / प्रतियोगिता का नाम</th>
                    <th style="padding: 6px; width: 18%; text-align: center; border: 1px solid #CBD5E1;">श्रेणी</th>
                    <th style="padding: 6px; width: 22%; border: 1px solid #CBD5E1;">विद्यार्थी की सीख / प्रस्तुति</th>
                    <th style="padding: 6px; width: 10%; text-align: center; border: 1px solid #CBD5E1;">अंक</th>
                </tr>
            </thead>
            <tbody>
                {act_part2_rows}
            </tbody>
        </table>

        <div style="border: 1px solid #CBD5E1; border-radius: 6px; padding: 12px; background: #F8FAFC; margin-top: 15px;">
            <div style="margin: 0 0 8px 0; color: #1E3A8A; font-weight: bold; font-size: 13px;">📝 सतत आंतरिक मूल्यांकन रूब्रिक्स (UP Board Criteria - पूर्णांक: 20):</div>
            <div style="display: flex; gap: 8px; font-size: 11.5px; text-align: center;">
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>1. नियमितता व सहभागिता</strong><br>(5 अंक)</div>
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>2. मौलिकता व शुद्धता</strong><br>(5 अंक)</div>
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>3. रचनात्मकता व नवाचार</strong><br>(5 अंक)</div>
                <div style="flex: 1; background: white; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>4. प्रस्तुतिकरण व अनुशासन</strong><br>(5 अंक)</div>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 40px; padding-top: 10px; border-top: 1px dashed #94A3B8; font-size: 12px;">
                <div>
                    <div>___________________________</div>
                    <div style="font-weight: bold; margin-top: 4px;">विद्यार्थी के हस्ताक्षर</div>
                    <div style="color: #64748B; font-size: 11px;">दिनांक: 08-09-2026</div>
                </div>
                <div style="text-align: center;">
                    <div>___________________________</div>
                    <div style="font-weight: bold; margin-top: 4px;">कक्षा अध्यापक हस्ताक्षर</div>
                    <div style="color: #64748B; font-size: 11px;">कक्षा अध्यापक ({cls_name})</div>
                </div>
                <div style="text-align: right;">
                    <div>___________________________</div>
                    <div style="font-weight: bold; margin-top: 4px;">प्रधानाचार्य / सील</div>
                    <div style="color: #64748B; font-size: 11px;">आदित्य बिड़ला इण्टर कॉलेज</div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>"""
    return html_content

# ----------------- Streamlit UI Application -----------------
st.title("🎓 Complete School Portfolio Generator & Form Sync")
st.caption("Aditya Birla Intermediate College | Form Auto-Sync with Direct Drive Fetcher")

# Sidebar: Permanent Data Storage Management
st.sidebar.header("📁 डेटा एवं फॉर्म सिंक")

is_master_saved = os.path.exists(SAVED_EXCEL_PATH)

if is_master_saved:
    st.sidebar.success("🔒 मास्टर प्रोफ़ाइल स्थायी रूप से सुरक्षित है!")
    with st.sidebar.expander("🗑️ मास्टर शीट डिलीट विकल्प"):
        if st.button("⚠️ सुरक्षित डेटा डिलीट करें"):
            os.remove(SAVED_EXCEL_PATH)
            if os.path.exists(SAVED_FORM_DATA_PATH):
                os.remove(SAVED_FORM_DATA_PATH)
            st.cache_data.clear()
            st.rerun()
else:
    st.sidebar.info("📌 केवल एक बार मास्टर एक्सेल फ़ाइल अपलोड करें:")
    master_file = st.sidebar.file_uploader("मास्टर शीट (.xlsx) चुनें", type=["xlsx"])
    if master_file:
        with open(SAVED_EXCEL_PATH, "wb") as f:
            f.write(master_file.getbuffer())
        st.cache_data.clear()
        st.sidebar.success("✅ डेटा डिस्क पर स्थायी सेव हो गया!")
        st.rerun()

# Google Form Sync
st.sidebar.markdown("---")
st.sidebar.subheader("🔗 गूगल फॉर्म रिस्पॉन्स सिंक")
form_sheet_url = st.sidebar.text_input(
    "Google Form की Sheet का शेयर लिंक डालें:",
    placeholder="https://docs.google.com/spreadsheets/d/.../edit?usp=sharing"
)

form_file_upload = st.sidebar.file_uploader("या फॉर्म रिस्पॉन्स CSV/Excel अपलोड करें:", type=["csv", "xlsx"])

if st.sidebar.button("🔄 गूगल फॉर्म डेटा सिंक करें"):
    synced_df = None
    if form_sheet_url:
        try:
            m = re.search(r'/d/([a-zA-Z0-9_-]+)', form_sheet_url)
            if m:
                sheet_id = m.group(1)
                export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                synced_df = pd.read_csv(export_url)
            else:
                st.sidebar.error("अमान्य Google Sheet URL!")
        except Exception as e:
            st.sidebar.error(f"शीट फ़ेच करने में त्रुटि: {e}")
    elif form_file_upload:
        try:
            if form_file_upload.name.endswith('.csv'):
                synced_df = pd.read_csv(form_file_upload)
            else:
                synced_df = pd.read_excel(form_file_upload)
        except Exception as e:
            st.sidebar.error(f"फ़ाइल पढ़ने में त्रुटि: {e}")

    if synced_df is not None:
        synced_df.to_csv(SAVED_FORM_DATA_PATH, index=False)
        st.cache_data.clear()
        st.sidebar.success(f"✅ फॉर्म डेटा से {len(synced_df)} रिकॉर्ड सिंक हो गए!")
        st.rerun()

if os.path.exists(SAVED_FORM_DATA_PATH):
    st.sidebar.info("⚡ फॉर्म रिस्पॉन्स डेटा सक्रिय है।")
    if st.sidebar.button("सिंक रीसेट करें"):
        os.remove(SAVED_FORM_DATA_PATH)
        st.cache_data.clear()
        st.rerun()

if not is_master_saved:
    st.info("👈 कृपया बाएँ साइडबार से 'All School Data' वाली मास्टर एक्सेल फ़ाइल अपलोड करें।")
    st.stop()

# ----------------- Data Loading and Auto Merge Logic -----------------
@st.cache_data
def load_and_merge_data():
    master = pd.read_excel(SAVED_EXCEL_PATH, header=0)
    master.columns = [c.strip() if isinstance(c, str) else c for c in master.columns]

    master['RollNo_Clean'] = master['RollNo'].apply(lambda x: clean_val(x, "")) if 'RollNo' in master.columns else master.iloc[:, 0].apply(lambda x: clean_val(x, ""))
    master['Class_Clean'] = master['Class'].apply(lambda x: clean_val(x, "General")) if 'Class' in master.columns else "General"

    if os.path.exists(SAVED_FORM_DATA_PATH):
        try:
            f_df = pd.read_csv(SAVED_FORM_DATA_PATH)
            f_df.columns = [c.strip() if isinstance(c, str) else c for c in f_df.columns]

            roll_col = next((c for c in f_df.columns if 'roll' in c.lower() or 'अनुक्रमांक' in c.lower()), None)
            photo_col = next((c for c in f_df.columns if 'photo' in c.lower() or 'image' in c.lower() or 'फोटो' in c.lower() or 'चित्र' in c.lower()), None)
            short_col = next((c for c in f_df.columns if 'short' in c.lower() or 'अल्पकालिक' in c.lower()), None)
            long_col = next((c for c in f_df.columns if 'long' in c.lower() or 'दीर्घकालिक' in c.lower() or 'career' in c.lower()), None)
            refl_col = next((c for c in f_df.columns if 'reflection' in c.lower() or 'चिंतन' in c.lower() or 'सुधार' in c.lower()), None)

            if roll_col:
                f_df['Roll_Key'] = f_df[roll_col].apply(lambda x: clean_val(x, ""))
                f_df = f_df.drop_duplicates(subset=['Roll_Key'], keep='last').set_index('Roll_Key')

                for idx, row in master.iterrows():
                    r_val = row['RollNo_Clean']
                    if r_val in f_df.index:
                        f_row = f_df.loc[r_val]
                        if short_col and pd.notna(f_row.get(short_col)):
                            master.at[idx, 'ShortGoal'] = str(f_row.get(short_col)).strip()
                        if long_col and pd.notna(f_row.get(long_col)):
                            master.at[idx, 'LongGoal'] = str(f_row.get(long_col)).strip()
                        if refl_col and pd.notna(f_row.get(refl_col)):
                            master.at[idx, 'Reflection'] = str(f_row.get(refl_col)).strip()
                        if photo_col and pd.notna(f_row.get(photo_col)):
                            master.at[idx, 'PhotoDriveLink'] = str(f_row.get(photo_col)).strip()
        except Exception:
            pass

    return master

df_master = load_and_merge_data()

# ----------------- TABS -----------------
tab1, tab2, tab3 = st.tabs(["👤 क्लास व रोल नंबर से खोजें", "📦 क्लास-वार बल्क डाउनलोड (ZIP)", "📷 लोकल फ़ोटो अपलोड"])

# ----------------- TAB 1: Single Portfolio with Dual Filter -----------------
with tab1:
    st.subheader("व्यक्तिगत छात्र पोर्टफोलियो खोज (Class & Roll No Filter)")

    available_classes = sorted(list(df_master['Class_Clean'].unique()))
    col_filter1, col_filter2 = st.columns([1, 2])

    with col_filter1:
        chosen_class = st.selectbox("1️⃣ कक्षा चुनें (Class):", available_classes)

    df_class = df_master[df_master['Class_Clean'] == chosen_class].copy()

    with col_filter2:
        student_display_map = {}
        for idx, row in df_class.iterrows():
            r_val = row['RollNo_Clean']
            n_val = clean_val(row.get('Name'))
            disp = f"Roll No: {r_val} | Name: {n_val}"
            student_display_map[disp] = idx

        selected_student_key = st.selectbox(
            f"2️⃣ रोल नंबर व छात्र चुनें (कक्षा {chosen_class} में कुल {len(df_class)} छात्र):", 
            list(student_display_map.keys())
        )

    if selected_student_key:
        s_idx = student_display_map[selected_student_key]
        student_dict = df_class.loc[s_idx].to_dict()

        roll_no = clean_val(student_dict.get('RollNo'))
        student_name = clean_val(student_dict.get('Name'))
        cls_val = clean_val(student_dict.get('Class'))

        st.markdown("---")
        info_col, photo_col = st.columns([3, 1])
        with info_col:
            st.write(f"### {student_name} (Roll No: {roll_no})")
            st.write(f"**कक्षा:** {cls_val} | **Adm No:** {clean_val(student_dict.get('AdmNo'))} | **House:** {clean_val(student_dict.get('House'))}")
            st.write(f"**पिता का नाम:** {clean_val(student_dict.get('FatherName'))} | **माता का नाम:** {clean_val(student_dict.get('MotherName'))}")
            st.write(f"**अल्पकालिक लक्ष्य (Goal):** {clean_val(student_dict.get('ShortGoal'))}")
            st.write(f"**दीर्घकालिक लक्ष्य (Career):** {clean_val(student_dict.get('LongGoal'))}")
        with photo_col:
            b64_img = get_image_base64(roll_no, student_dict.get('PhotoDriveLink'))
            if b64_img:
                st.markdown(f'<img src="{b64_img}" width="110" style="border-radius:6px; border:2px solid #1E3A8A;"/>', unsafe_allow_html=True)
            else:
                st.info("फ़ोटो उपलब्ध नहीं")

        html_code = generate_full_portfolio_html(student_dict)
        file_name = f"Portfolio_{cls_val}_Roll_{roll_no}_{student_name}.html"

        st.download_button(
            label=f"⬇️ डाउनलोड 3-पेज पोर्टफोलियो ({student_name})",
            data=html_code,
            file_name=file_name,
            mime="text/html"
        )

        with st.expander("👁️ 3-पेज पोर्टफोलियो का लाइव प्रीव्यू देखें"):
            st.components.v1.html(html_code, height=900, scrolling=True)

# ----------------- TAB 2: Class-Wise Bulk ZIP -----------------
with tab2:
    st.subheader("कक्षा-वार / पूरे विद्यालय का बल्क ZIP डाउनलोड")

    zip_class_options = ["संपूर्ण विद्यालय (All Classes)"] + available_classes
    zip_selection = st.selectbox("किस कक्षा के सभी पोर्टफोलियो जनरेट करने हैं?", zip_class_options)

    if zip_selection == "संपूर्ण विद्यालय (All Classes)":
        target_df = df_master.copy()
        zip_file_name = "All_School_Portfolios.zip"
    else:
        target_df = df_master[df_master['Class_Clean'] == zip_selection].copy()
        zip_file_name = f"Portfolios_Class_{zip_selection}.zip"

    st.write(f"तैयार होने वाले कुल पोर्टफोलियो: **{len(target_df)}**")

    if st.button(f"🚀 ज़िप फ़ाइल जनरेट करें ({zip_selection})"):
        progress_bar = st.progress(0)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for i, (_, row) in enumerate(target_df.iterrows()):
                s_dict = row.to_dict()
                r_no = clean_val(s_dict.get('RollNo'))
                s_nm = clean_val(s_dict.get('Name'))
                c_nm = clean_val(s_dict.get('Class'))
                s_html = generate_full_portfolio_html(s_dict)

                entry_name = f"{c_nm}/Portfolio_Roll_{r_no}_{s_nm}.html"
                zip_file.writestr(entry_name, s_html)
                progress_bar.progress((i + 1) / len(target_df))

        st.success("✅ सभी चयनित पोर्टफोलियो तैयार हो चुके हैं!")
        st.download_button(
            label=f"⬇️ डाउनलोड ZIP फ़ाइल ({zip_file_name})",
            data=zip_buffer.getvalue(),
            file_name=zip_file_name,
            mime="application/zip"
        )

# ----------------- TAB 3: Single Photo Update -----------------
with tab3:
    st.subheader("लोकल फ़ोटो अपलोड / अपडेट करें")
    s_roll = st.text_input("छात्र का रोल नंबर (RollNo) दर्ज करें:")
    s_file = st.file_uploader("फ़ोटो फ़ाइल चुनें (JPEG/PNG):", type=["jpg", "jpeg", "png"], key="tab3_p")

    if s_roll and s_file:
        if st.button("💾 फ़ोटो सेव करें"):
            remove_existing_photos(s_roll)
            ext = os.path.splitext(s_file.name)[1].lower()
            save_path = os.path.join(PHOTOS_DIR, f"{s_roll.strip()}{ext}")
            with open(save_path, "wb") as f:
                f.write(s_file.getbuffer())
            st.success(f"✅ रोल नंबर {s_roll} के लिए फ़ोटो सुरक्षित हो गई!")
            st.rerun()
