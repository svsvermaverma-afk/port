import streamlit as st
import pandas as pd
import os
import zipfile
import base64
import io

# ----------------- Directories Setup -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_DIR = os.path.join(DATA_DIR, "excel")
PHOTOS_DIR = os.path.join(DATA_DIR, "photos")
OUTPUT_DIR = os.path.join(DATA_DIR, "output_html")

for d in [EXCEL_DIR, PHOTOS_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

SAVED_EXCEL_PATH = os.path.join(EXCEL_DIR, "master_data.xlsx")

st.set_page_config(page_title="Complete Student Portfolio Generator", layout="wide")

# ----------------- Helper Functions -----------------
def clean_val(val, default="—"):
    if pd.isna(val) or val is None:
        return default
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str if val_str != "" else default

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

# ----------------- 3-Page HTML Builder -----------------
def generate_full_portfolio_html(row):
    # A to V: Personal Details (Indices 0 to 21)
    roll_no     = clean_val(row[0] if len(row) > 0 else "")
    cls_name    = clean_val(row[1] if len(row) > 1 else "", "12-B")
    sr_no       = clean_val(row[2] if len(row) > 2 else "")
    roll_10th   = clean_val(row[3] if len(row) > 3 else "")
    pen_no      = clean_val(row[4] if len(row) > 4 else "")
    aadhar_no   = clean_val(row[5] if len(row) > 5 else "")
    dob_dd      = clean_val(row[6] if len(row) > 6 else "")
    dob_mm      = clean_val(row[7] if len(row) > 7 else "")
    dob_yyyy    = clean_val(row[8] if len(row) > 8 else "")
    dob         = f"{dob_dd}/{dob_mm}/{dob_yyyy}" if (dob_dd != "—" and dob_yyyy != "—") else "—"
    st_name     = clean_val(row[9] if len(row) > 9 else "")
    f_name      = clean_val(row[10] if len(row) > 10 else "")
    m_name      = clean_val(row[11] if len(row) > 11 else "")
    gender      = clean_val(row[12] if len(row) > 12 else "")
    caste       = clean_val(row[13] if len(row) > 13 else "")
    cat         = clean_val(row[14] if len(row) > 14 else "")
    religion    = clean_val(row[15] if len(row) > 15 else "")
    address     = clean_val(row[16] if len(row) > 16 else "")
    mobile      = clean_val(row[17] if len(row) > 17 else "")
    email       = clean_val(row[18] if len(row) > 18 else "")
    occupation  = clean_val(row[19] if len(row) > 19 else "")
    ecode       = clean_val(row[20] if len(row) > 20 else "")
    dept        = clean_val(row[21] if len(row) > 21 else "")

    # AH to AM: Monthly Test Marks (Indices 33 to 38)
    m_hindi     = clean_val(row[33] if len(row) > 33 else "")
    m_eng       = clean_val(row[34] if len(row) > 34 else "")
    m_maths     = clean_val(row[35] if len(row) > 35 else "")
    m_phy       = clean_val(row[36] if len(row) > 36 else "")
    m_che       = clean_val(row[37] if len(row) > 37 else "")
    m_total     = clean_val(row[38] if len(row) > 38 else "")

    # Goals & Reflections
    short_goal  = clean_val(row[39] if len(row) > 39 else "")
    long_goal   = clean_val(row[40] if len(row) > 40 else "")
    reflection  = clean_val(row[69] if len(row) > 69 else "नियमित अभ्यास, अनुशासन एवं समय प्रबंधन पर विशेष ध्यान।")

    # Photo Tag
    img_b64 = get_image_base64(roll_no)
    img_tag = f'<img src="{img_b64}" style="width: 105px; height: 130px; object-fit: cover; border-radius: 6px; border: 2px solid #1E3A8A;"/>' if img_b64 else '<div style="width: 105px; height: 130px; border-radius: 6px; border: 2px dashed #94A3B8; display:flex; align-items:center; justify-content:center; color:#64748B; font-size:11px; text-align:center; padding:5px;">फ़ोटो उपलब्ध नहीं</div>'

    # W to AG: Attendance (Indices 22 to 32)
    att_months = [
        ("अप्रैल", 22), ("मई", 23), ("जुलाई", 24), ("अगस्त", 25),
        ("सितंबर", 26), ("अक्टूबर", 27), ("नवंबर", 28), ("दिसंबर", 29),
        ("जनवरी", 30), ("कुल उपस्थिति", 31), ("प्रतिशत (%)", 32)
    ]
    att_headers_html = "".join([f'<th style="padding: 6px; border: 1px solid #CBD5E1; text-align: center;">{m[0]}</th>' for m in att_months])
    att_values_html = "".join([f'<td style="padding: 6px; border: 1px solid #CBD5E1; text-align: center; font-weight: 600;">{clean_val(row[m[1]] if len(row) > m[1] else "")}</td>' for m in att_months])

    # Activities 1 to 7 (Part 1 - Indices 41 to 54)
    act_part1 = [
        ("27.08.2026", "1. Tata Building India School Essay Competition", "साहित्यिक (निबंध)", 41, 42),
        ("27.08.2026", "2. रंगोली प्रतियोगिता (Rangoli Making)", "कला एवं संस्कृति", 43, 44),
        ("27.08.2026", "3. मेहंदी प्रतियोगिता (Mehndi Design)", "कला एवं संस्कृति", 45, 46),
        ("20.08.2026", "4. राखी निर्माण प्रतियोगिता (Rakhi Making)", "क्राफ्ट एवं रचनात्मकता", 47, 48),
        ("13.08.2026", "5. चित्रकला प्रतियोगिता (Drawing)", "दृश्य कला (Fine Arts)", 49, 50),
        ("06.08.2026", "6. निबंध प्रतियोगिता (Essay Writing)", "साहित्यिक कौशल", 51, 52),
        ("30.07.2026", "7. बाल संसद गतिविधियां (Bal Sansad)", "नेतृत्व एवं सामाजिक कौशल", 53, 54),
    ]
    act_part1_rows = ""
    for dt, title, cat, p_idx, r_idx in act_part1:
        part_text = clean_val(row[p_idx] if len(row) > p_idx else "", "सक्रिय प्रतिभाग")
        rem_text = clean_val(row[r_idx] if len(row) > r_idx else "", part_text)
        act_part1_rows += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{dt}</td>
                <td style="padding: 6px; font-weight: 600; color: #1E3A8A; border: 1px solid #CBD5E1;">{title}<br><span style="font-weight: normal; color: #475569; font-size: 11px;">{part_text}</span></td>
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{cat}</td>
                <td style="padding: 6px; color: #0284C7; font-style: italic; border: 1px solid #CBD5E1;">{rem_text}</td>
                <td style="padding: 6px; text-align: center; font-weight: bold; color: #059669; border: 1px solid #CBD5E1;">5/5</td>
            </tr>
        """

    # Activities 8 to 14 (Part 2 - Indices 55 to 68)
    act_part2 = [
        ("24.07.2026", "8. कक्षा सज्जा एवं चार्ट (Class Decoration)", "रचनात्मक एवं नवाचार", 55, 56),
        ("16.07.2026", "9. भाषण प्रतियोगिता (Speech/Elocution)", "वाक कौशल व आत्मविश्वास", 57, 58),
        ("09.07.2026", "10. कहानी लेखन (Story Writing)", "साहित्यिक सृजन", 59, 60),
        ("02.07.2026", "11. आई.ई.पी. पोर्टफोलियो (IEP Portfolio)", "शैक्षणिक पोर्टफोलियो कार्य", 61, 62),
        ("25.04.2026", "12. समूह चर्चा (Group Discussion)", "संवाद एवं संप्रेषण कौशल", 63, 64),
        ("18.04.2026", "13. लेख प्रतियोगिता (Article Writing)", "वैचारिक एवं सामाजिक लेखन", 65, 66),
        ("10.04.2026", "14. मौलिक रचना (Creative Story/Writing)", "मौलिक रचनात्मकता", 67, 68),
    ]
    act_part2_rows = ""
    for dt, title, cat, p_idx, r_idx in act_part2:
        part_text = clean_val(row[p_idx] if len(row) > p_idx else "", "सक्रिय प्रतिभाग")
        rem_text = clean_val(row[r_idx] if len(row) > r_idx else "", part_text)
        act_part2_rows += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{dt}</td>
                <td style="padding: 6px; font-weight: 600; color: #1E3A8A; border: 1px solid #CBD5E1;">{title}<br><span style="font-weight: normal; color: #475569; font-size: 11px;">{part_text}</span></td>
                <td style="padding: 6px; text-align: center; border: 1px solid #CBD5E1;">{cat}</td>
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
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f8fafc; padding: 10px; color: #1e293b; line-height: 1.35; }}
        .page {{ max-width: 850px; margin: 0 auto 25px auto; background: #ffffff; border: 2px solid #1E3A8A; border-radius: 8px; padding: 22px; box-shadow: 0 4px 10px rgba(0,0,0,0.06); }}
        .header {{ text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 10px; margin-bottom: 14px; }}
        .sec-title {{ color: #1E3A8A; font-weight: bold; font-size: 13.5px; margin-top: 14px; margin-bottom: 6px; display: flex; align-items: center; gap: 6px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
        @media print {{
            body {{ background: none; padding: 0; }}
            .page {{ box-shadow: none; margin: 0; border: 2px solid #000; page-break-after: always; }}
        }}
    </style>
</head>
<body>

    <!-- ========================================== PAGE 1 ========================================== -->
    <div class="page">
        <div class="header">
            <h2 style="margin: 0; color: #1E3A8A; font-size: 21px; text-transform: uppercase; letter-spacing: 0.5px;">ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT, SONEBHADRA (UP)</h2>
            <h3 style="margin: 3px 0 0 0; color: #059669; font-size: 16px;">छात्र संपूर्ण पोर्टफोलियो एवं सतत आंतरिक मूल्यांकन रिकॉर्ड</h3>
            <div style="font-size: 12.5px; color: #475569; margin-top: 3px;">सत्र: 2026 - 2027 | कक्षा: {cls_name}</div>
            <div style="display: inline-block; background: #1E3A8A; color: white; padding: 2px 12px; border-radius: 10px; font-size: 11px; margin-top: 5px; font-weight: 600;">भाग 1 : व्यक्तिगत विवरण, पहचान एवं मासिक परीक्षा रिकॉर्ड</div>
        </div>

        <div style="display: flex; gap: 15px; margin-bottom: 10px;">
            <table style="width: 74%; border: 1px solid #CBD5E1;">
                <tr style="background: #F1F5F9;"><td style="padding: 5px; font-weight: bold; width: 28%; border: 1px solid #CBD5E1;">छात्र का नाम:</td><td style="padding: 5px; color: #1E3A8A; font-weight: bold; font-size: 13px; border: 1px solid #CBD5E1;">{st_name}</td><td style="padding: 5px; font-weight: bold; width: 22%; border: 1px solid #CBD5E1;">अनुक्रमांक (Roll No.):</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">{roll_no}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">कक्षा (Class):</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{cls_name}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">10th Roll No.:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{roll_10th}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">S.R. No.:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{sr_no}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">PEN Number:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{pen_no}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">आधार संख्या (Aadhar):</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{aadhar_no}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">जन्म तिथि (D.O.B.):</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{dob}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">पिता का नाम:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{f_name}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">माता का नाम:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{m_name}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">लिंग / धर्म (Gender/Rel.):</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{gender} / {religion}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">जाति / वर्ग (Caste/Cat.):</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{caste} / {cat}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">संपर्क सूत्र (Mobile):</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{mobile}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">ईमेल (Email):</td><td style="padding: 5px; border: 1px solid #CBD5E1; word-break: break-all;">{email}</td></tr>
                <tr><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">अभिभावक व्यवसाय:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{occupation}</td><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">E.Code / Dept.:</td><td style="padding: 5px; border: 1px solid #CBD5E1;">{ecode} / {dept}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 5px; font-weight: bold; border: 1px solid #CBD5E1;">स्थायी पता (Address):</td><td colspan="3" style="padding: 5px; border: 1px solid #CBD5E1;">{address}</td></tr>
            </table>
            
            <div style="width: 26%; border: 2px dashed #94A3B8; border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; background: #F8FAFC; padding: 8px; text-align: center;">
                {img_tag}
                <div style="font-weight: bold; font-size: 12.5px; color: #1E3A8A; margin-top: 6px;">{st_name}</div>
                <div style="font-size: 11px; color: #64748B;">रोल नंबर: {roll_no} | {cls_name}</div>
                <div style="font-size: 10px; color: #059669; margin-top: 4px; border: 1px solid #059669; padding: 2px 6px; border-radius: 8px;">सत्यापित विद्यार्थी</div>
            </div>
        </div>

        <div class="sec-title">📊 मासिक परीक्षा मूल्यांकन (Monthly Test Evaluation Record):</div>
        <table style="border: 1px solid #CBD5E1; text-align: center;">
            <tr style="background: #1E3A8A; color: white; font-weight: 600;">
                <th style="padding: 6px; border: 1px solid #CBD5E1;">हिंदी (20)</th>
                <th style="padding: 6px; border: 1px solid #CBD5E1;">अंग्रेजी (20)</th>
                <th style="padding: 6px; border: 1px solid #CBD5E1;">गणित (20)</th>
                <th style="padding: 6px; border: 1px solid #CBD5E1;">भौतिक विज्ञान (20)</th>
                <th style="padding: 6px; border: 1px solid #CBD5E1;">रसायन विज्ञान (20)</th>
                <th style="padding: 6px; border: 1px solid #CBD5E1; background: #059669;">कुल प्राप्तांक (100)</th>
            </tr>
            <tr style="background: #F8FAFC; font-size: 13px; font-weight: 600;">
                <td style="padding: 7px; border: 1px solid #CBD5E1;">{m_hindi}</td>
                <td style="padding: 7px; border: 1px solid #CBD5E1;">{m_eng}</td>
                <td style="padding: 7px; border: 1px solid #CBD5E1;">{m_maths}</td>
                <td style="padding: 7px; border: 1px solid #CBD5E1;">{m_phy}</td>
                <td style="padding: 7px; border: 1px solid #CBD5E1;">{m_che}</td>
                <td style="padding: 7px; border: 1px solid #CBD5E1; color: #059669; font-weight: bold; font-size: 14px;">{m_total}</td>
            </tr>
        </table>

        <div class="sec-title">🎯 शैक्षणिक लक्ष्य एवं संकल्प (Academic Vision & Career Goals):</div>
        <div style="display: flex; gap: 10px;">
            <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #3B82F6; padding: 7px 10px; border-radius: 4px; font-size: 12px;">
                <strong style="color: #1E3A8A;">📌 अल्पकालिक लक्ष्य (Short-Term Goal 2026-27):</strong><br>
                {short_goal}
            </div>
            <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #059669; padding: 7px 10px; border-radius: 4px; font-size: 12px;">
                <strong style="color: #059669;">🎯 दीर्घकालिक लक्ष्य (Long-Term Goal - Career):</strong><br>
                {long_goal}
            </div>
        </div>

        <div class="sec-title">💡 स्व-चिंतन एवं सुधार क्षेत्र (Student Self-Reflection & Strengths):</div>
        <div style="background: #F8FAFC; border-left: 4px solid #10B981; padding: 8px 12px; border-radius: 4px; font-size: 12px; color: #334155;">
            {reflection}
        </div>
    </div>

    <!-- ========================================== PAGE 2 ========================================== -->
    <div class="page">
        <div class="header">
            <h3 style="margin: 0; color: #1E3A8A; font-size: 18px; text-transform: uppercase;">सह-पाठ्यचर्या, उपस्थिति एवं गतिविधि मूल्यांकन प्रपत्र (भाग - 1)</h3>
            <div style="font-size: 12px; color: #475569; margin-top: 3px;">विद्यार्थी: <strong>{st_name}</strong> | अनुक्रमांक: <strong>{roll_no}</strong> | कक्षा: <strong>{cls_name}</strong></div>
            <div style="display: inline-block; background: #059669; color: white; padding: 2px 12px; border-radius: 10px; font-size: 11px; margin-top: 5px; font-weight: 600;">भाग 2 : माह-वार उपस्थिति चार्ट एवं प्रमुख प्रतियोगिताएं (1 से 7)</div>
        </div>

        <div class="sec-title">📅 माह-वार उपस्थिति विवरण (Monthly Attendance Progression 2026-27):</div>
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

    <!-- ========================================== PAGE 3 ========================================== -->
    <div class="page">
        <div class="header">
            <h3 style="margin: 0; color: #1E3A8A; font-size: 18px; text-transform: uppercase;">सह-पाठ्यचर्या, गतिविधि मूल्यांकन एवं सत्यापन (भाग - 2)</h3>
            <div style="font-size: 12px; color: #475569; margin-top: 3px;">विद्यार्थी: <strong>{st_name}</strong> | अनुक्रमांक: <strong>{roll_no}</strong> | कक्षा: <strong>{cls_name}</strong></div>
            <div style="display: inline-block; background: #D97706; color: white; padding: 2px 12px; border-radius: 10px; font-size: 11px; margin-top: 5px; font-weight: 600;">भाग 3 : गतिविधियां (8 से 14), आंतरिक मूल्यांकन रूब्रिक्स व अंतिम प्रमाणीकरण</div>
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
                    <div style="color: #64748B; font-size: 11px;">दिनांक: 07-09-2026</div>
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
st.title("🎓 Complete Student Portfolio Generator (All Fields Included)")
st.caption("Aditya Birla Intermediate College | 100% Data Field Coverage (Col A to BQ)")

# Sidebar: File Management
st.sidebar.header("📁 फ़ाइल प्रबंधन")

# 1. Master Excel Upload
excel_file = st.sidebar.file_uploader("1. मास्टर एक्सेल शीट अपलोड करें (.xlsx)", type=["xlsx"])
if excel_file:
    with open(SAVED_EXCEL_PATH, "wb") as f:
        f.write(excel_file.getbuffer())
    st.sidebar.success("✅ एक्सेल फ़ाइल सेव हो गई!")

# 2. Bulk Photo Upload
st.sidebar.subheader("📸 बल्क फ़ोटो अपलोड")
multi_photos = st.sidebar.file_uploader(
    "एक साथ सभी छात्र फ़ोटो चुनें (JPEG/JPG/PNG)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True,
    help="फ़ोटो का नाम रोल नंबर पर होना चाहिए (उदा: 39.jpg)"
)
if multi_photos:
    count = 0
    for p in multi_photos:
        filename = os.path.basename(p.name)
        save_dest = os.path.join(PHOTOS_DIR, filename)
        with open(save_dest, "wb") as f:
            f.write(p.getbuffer())
        count += 1
    st.sidebar.success(f"✅ {count} फ़ोटो सफलतापूर्वक लोड हुईं!")

if not os.path.exists(SAVED_EXCEL_PATH):
    st.info("👈 कृपया बाएँ साइडबार से मास्टर एक्सेल फ़ाइल अपलोड करें।")
    st.stop()

@st.cache_data
def load_data(path):
    return pd.read_excel(path, header=0)

df = load_data(SAVED_EXCEL_PATH)

tab1, tab2, tab3 = st.tabs(["👤 सिंगल पोर्टफोलियो", "📦 पूरी क्लास का बल्क डाउनलोड (ZIP)", "📷 एक फ़ोटो अपलोड"])

# ----------------- TAB 1: Single Portfolio -----------------
with tab1:
    st.subheader("व्यक्तिगत छात्र पोर्टफोलियो प्रीव्यू एवं डाउनलोड")
    roll_col = df.columns[0]
    name_col = df.columns[9] if len(df.columns) > 9 else df.columns[1]

    student_options = {f"Roll: {clean_val(row[roll_col])} - {clean_val(row[name_col])}": idx for idx, row in df.iterrows()}
    selected_label = st.selectbox("विद्यार्थी का चयन करें:", list(student_options.keys()))

    if selected_label:
        s_idx = student_options[selected_label]
        student_row = df.iloc[s_idx].tolist()
        roll_no = clean_val(student_row[0])
        student_name = clean_val(student_row[9])
        
        c1, c2 = st.columns([3, 1])
        with c1:
            st.write(f"**नाम:** {student_name} | **रोल नंबर:** {roll_no} | **कक्षा:** {clean_val(student_row[1])}")
            st.write(f"**PEN No:** {clean_val(student_row[4])} | **Aadhar:** {clean_val(student_row[5])}")
        with c2:
            b64_img = get_image_base64(roll_no)
            if b64_img:
                st.markdown(f'<img src="{b64_img}" width="95" style="border-radius:6px; border:1px solid #1E3A8A;"/>', unsafe_allow_html=True)
            else:
                st.warning("फ़ोटो लिंक नहीं है")

        html_code = generate_full_portfolio_html(student_row)
        file_name = f"Complete_Portfolio_Roll_{roll_no}_{student_name}.html"

        st.download_button(
            label=f"⬇️ डाउनलोड 3-पेज पोर्टफोलियो ({file_name})",
            data=html_code,
            file_name=file_name,
            mime="text/html"
        )

        with st.expander("👁️ 3-पेज पोर्टफोलियो का लाइव प्रीव्यू देखें"):
            st.components.v1.html(html_code, height=950, scrolling=True)

# ----------------- TAB 2: Bulk Generation (ZIP) -----------------
with tab2:
    st.subheader("पूरी कक्षा के सभी 3-पेज पोर्टफोलियो एक क्लिक में डाउनलोड करें")
    st.write(f"शीट में कुल विद्यार्थी: **{len(df)}**")

    if st.button("🚀 सभी छात्रों के 3-पेज पोर्टफोलियो जनरेट करें (ZIP)"):
        progress_bar = st.progress(0)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for idx, row in df.iterrows():
                s_list = row.tolist()
                r_no = clean_val(s_list[0])
                s_nm = clean_val(s_list[9])
                s_html = generate_full_portfolio_html(s_list)
                arc_name = f"Portfolio_Roll_{r_no}_{s_nm}.html"
                zip_file.writestr(arc_name, s_html)
                progress_bar.progress((idx + 1) / len(df))

        st.success("✅ सभी छात्रों के पोर्टफोलियो तैयार हैं!")
        st.download_button(
            label="⬇️ सभी 3-पेज पोर्टफोलियो डाउनलोड करें (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="All_Class_Portfolios_Complete.zip",
            mime="application/zip"
        )

# ----------------- TAB 3: Single Photo Upload -----------------
with tab3:
    st.subheader("किसी छात्र की फ़ोटो अपडेट करें")
    s_roll = st.text_input("छात्र का रोल नंबर:")
    s_file = st.file_uploader("फ़ोटो फ़ाइल चुनें (JPEG/PNG):", type=["jpg", "jpeg", "png"], key="tab3_p")

    if s_roll and s_file:
        if st.button("💾 फ़ोटो सेव करें"):
            ext = os.path.splitext(s_file.name)[1].lower()
            save_path = os.path.join(PHOTOS_DIR, f"{s_roll.strip()}{ext}")
            with open(save_path, "wb") as f:
                f.write(s_file.getbuffer())
            st.success(f"✅ रोल नंबर {s_roll} के लिए फ़ोटो सफलतापूर्वक अपडेट हो गई!")
