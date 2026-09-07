import streamlit as st
import pandas as pd
import os
import zipfile
from PIL import Image
import io

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# ----------------- Directories Setup -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_DIR = os.path.join(DATA_DIR, "excel")
PHOTOS_DIR = os.path.join(DATA_DIR, "photos")
OUTPUT_DIR = os.path.join(DATA_DIR, "output_pdfs")

for d in [EXCEL_DIR, PHOTOS_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

SAVED_EXCEL_PATH = os.path.join(EXCEL_DIR, "master_data.xlsx")

st.set_page_config(page_title="Student Portfolio Generator", layout="wide")


# ----------------- PDF Generator Engine -----------------
def generate_student_portfolio_pdf(student_data, photo_path, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        leading=18,
        alignment=1,  # Center
        textColor=colors.HexColor("#1A365D"),
        fontName="Helvetica-Bold"
    )
    sub_title = ParagraphStyle(
        'SubTitleStyle',
        parent=styles['Normal'],
        fontSize=10,
        leading=12,
        alignment=1,
        textColor=colors.HexColor("#4A5568")
    )
    section_head = ParagraphStyle(
        'SectionHead',
        fontSize=11,
        leading=13,
        textColor=colors.HexColor("#FFFFFF"),
        fontName="Helvetica-Bold"
    )
    cell_bold = ParagraphStyle('CBold', fontSize=8, leading=10, fontName="Helvetica-Bold")
    cell_norm = ParagraphStyle('CNorm', fontSize=8, leading=10, fontName="Helvetica")

    def section_banner(title):
        t = Table([[Paragraph(f"<b>{title}</b>", section_head)]], colWidths=[545])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#2B6CB0")),
            ('PADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        return t

    def get_val(col_idx):
        if col_idx < len(student_data):
            val = student_data[col_idx]
            if pd.isna(val) or val is None:
                return "—"
            return str(val).strip()
        return "—"

    story = []

    # ================= PAGE 1 =================
    story.append(Paragraph("STUDENT COMPREHENSIVE PORTFOLIO", title_style))
    story.append(Paragraph("Academic Session & Performance Record", sub_title))
    story.append(Spacer(1, 10))

    # Profile Header with Photo
    if photo_path and os.path.exists(photo_path):
        try:
            st_img = RLImage(photo_path, width=1.1 * inch, height=1.3 * inch)
        except Exception:
            st_img = Paragraph("<i>No Photo</i>", cell_norm)
    else:
        st_img = Paragraph("<br/><br/><i>Photo Box</i>", cell_norm)

    p_table_data = [
        [Paragraph("<b>Student Name:</b>", cell_bold), Paragraph(get_val(9), cell_norm),
         Paragraph("<b>Roll No:</b>", cell_bold), Paragraph(get_val(0), cell_norm), st_img],
        [Paragraph("<b>Class:</b>", cell_bold), Paragraph(get_val(1), cell_norm),
         Paragraph("<b>S.R. No:</b>", cell_bold), Paragraph(get_val(2), cell_norm), ""],
        [Paragraph("<b>Father's Name:</b>", cell_bold), Paragraph(get_val(10), cell_norm),
         Paragraph("<b>Mother's Name:</b>", cell_bold), Paragraph(get_val(11), cell_norm), ""],
        [Paragraph("<b>DOB:</b>", cell_bold), Paragraph(f"{get_val(6)}/{get_val(7)}/{get_val(8)}", cell_norm),
         Paragraph("<b>Gender / Caste:</b>", cell_bold), Paragraph(f"{get_val(12)} / {get_val(13)}", cell_norm), ""],
        [Paragraph("<b>Aadhar No:</b>", cell_bold), Paragraph(get_val(5), cell_norm),
         Paragraph("<b>PEN No:</b>", cell_bold), Paragraph(get_val(4), cell_norm), ""],
        [Paragraph("<b>Contact:</b>", cell_bold), Paragraph(get_val(17), cell_norm),
         Paragraph("<b>Email:</b>", cell_bold), Paragraph(get_val(18), cell_norm), ""],
        [Paragraph("<b>Address:</b>", cell_bold), Paragraph(get_val(16), cell_norm),
         Paragraph("<b>Parent Dept/Emp:</b>", cell_bold), Paragraph(f"{get_val(20)} - {get_val(21)}", cell_norm), ""]
    ]

    p_table = Table(p_table_data, colWidths=[90, 140, 95, 135, 85])
    p_table.setStyle(TableStyle([
        ('SPAN', (4, 0), (4, 6)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (3, -1), 0.25, colors.lightgrey),
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
        ('ALIGN', (4, 0), (4, 6), 'CENTER'),
        ('VALIGN', (4, 0), (4, 6), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(p_table)
    story.append(Spacer(1, 15))

    # Academic Marks Table (Cols 33 to 38: AH to AM)
    story.append(section_banner("MONTHLY TEST EVALUATION (MARKS)"))
    story.append(Spacer(1, 4))

    marks_data = [
        [Paragraph("<b>Hindi (20)</b>", cell_bold), Paragraph("<b>English (20)</b>", cell_bold),
         Paragraph("<b>Maths (20)</b>", cell_bold), Paragraph("<b>Physics (20)</b>", cell_bold),
         Paragraph("<b>Chemistry (20)</b>", cell_bold), Paragraph("<b>Total (100)</b>", cell_bold)],
        [Paragraph(get_val(33), cell_norm), Paragraph(get_val(34), cell_norm),
         Paragraph(get_val(35), cell_norm), Paragraph(get_val(36), cell_norm),
         Paragraph(get_val(37), cell_norm), Paragraph(f"<b>{get_val(38)}</b>", cell_bold)]
    ]
    marks_table = Table(marks_data, colWidths=[90, 90, 90, 90, 90, 95])
    marks_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(marks_table)
    story.append(Spacer(1, 15))

    # Aspirations & Focus
    story.append(section_banner("STUDENT GOALS & OBJECTIVES"))
    story.append(Spacer(1, 4))
    goals_data = [
        [Paragraph("<b>Short-Term Goals:</b>", cell_bold), Paragraph(get_val(39), cell_norm)],
        [Paragraph("<b>Long-Term Goals:</b>", cell_bold), Paragraph(get_val(40), cell_norm)]
    ]
    goals_table = Table(goals_data, colWidths=[120, 425])
    goals_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('PADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(goals_table)

    story.append(PageBreak())

    # ================= PAGE 2 =================
    story.append(Paragraph("ATTENDANCE & CO-CURRICULAR PARTICIPATION (PART 1)", title_style))
    story.append(Spacer(1, 10))

    # Attendance Records (Cols 22 to 32: W to AG)
    story.append(section_banner("ATTENDANCE PROGRESSION"))
    story.append(Spacer(1, 4))

    att_headers = ["Apr", "May", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Total", "%"]
    att_vals = [get_val(i) for i in range(22, 33)]
    att_table_data = [
        [Paragraph(f"<b>{h}</b>", cell_bold) for h in att_headers],
        [Paragraph(v, cell_norm) for v in att_vals]
    ]
    att_table = Table(att_table_data, colWidths=[49] * 11)
    att_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(att_table)
    story.append(Spacer(1, 15))

    # Co-curricular Activities (1 to 7)
    story.append(section_banner("CO-CURRICULAR ACTIVITIES & ASSESSMENTS (1 TO 7)"))
    story.append(Spacer(1, 4))

    act_list_1 = [
        ("1. Tata Essay Competition", get_val(41), get_val(42)),
        ("2. Rangoli Competition", get_val(43), get_val(44)),
        ("3. Mehndi Competition", get_val(45), get_val(46)),
        ("4. Rakhi Making Competition", get_val(47), get_val(48)),
        ("5. Drawing Competition", get_val(49), get_val(50)),
        ("6. Essay Writing Competition", get_val(51), get_val(52)),
        ("7. Bal Sansad Activities", get_val(53), get_val(54)),
    ]

    act_data_1 = [[Paragraph("<b>Activity Name</b>", cell_bold), Paragraph("<b>Participation / Role</b>", cell_bold),
                   Paragraph("<b>Remarks / Grade</b>", cell_bold)]]
    for name, part, rem in act_list_1:
        act_data_1.append([Paragraph(name, cell_bold), Paragraph(part, cell_norm), Paragraph(rem, cell_norm)])

    act_table_1 = Table(act_data_1, colWidths=[180, 185, 180])
    act_table_1.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#EDF2F7")),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(act_table_1)

    story.append(PageBreak())

    # ================= PAGE 3 =================
    story.append(Paragraph("CO-CURRICULAR PARTICIPATION (PART 2) & REFLECTIONS", title_style))
    story.append(Spacer(1, 10))

    story.append(section_banner("CO-CURRICULAR ACTIVITIES & ASSESSMENTS (8 TO 14)"))
    story.append(Spacer(1, 4))

    act_list_2 = [
        ("8. Class Decoration", get_val(55), get_val(56)),
        ("9. Speech / Elocution", get_val(57), get_val(58)),
        ("10. Story Writing", get_val(59), get_val(60)),
        ("11. IEP Portfolio Work", get_val(61), get_val(62)),
        ("12. Group Discussions", get_val(63), get_val(64)),
        ("13. Article Contributions", get_val(65), get_val(66)),
        ("14. Creative Story / Writing", get_val(67), get_val(68)),
    ]

    act_data_2 = [[Paragraph("<b>Activity Name</b>", cell_bold), Paragraph("<b>Participation / Role</b>", cell_bold),
                   Paragraph("<b>Remarks / Grade</b>", cell_bold)]]
    for name, part, rem in act_list_2:
        act_data_2.append([Paragraph(name, cell_bold), Paragraph(part, cell_norm), Paragraph(rem, cell_norm)])

    act_table_2 = Table(act_data_2, colWidths=[180, 185, 180])
    act_table_2.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#EDF2F7")),
        ('PADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(act_table_2)
    story.append(Spacer(1, 15))

    # Reflection Box (BQ)
    story.append(section_banner("STUDENT REFLECTION / OVERALL OBSERVATION"))
    story.append(Spacer(1, 4))
    reflection_text = get_val(69)
    ref_table = Table(
        [[Paragraph(reflection_text if reflection_text != "—" else "No reflections recorded for this session.",
                    cell_norm)]], colWidths=[545])
    ref_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
        ('PADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(ref_table)
    story.append(Spacer(1, 45))

    # Signature Block
    sig_data = [
        [Paragraph("___________________________<br/><b>Student Signature</b>", cell_bold),
         Paragraph("___________________________<br/><b>Class Teacher Signature</b>", cell_bold),
         Paragraph("___________________________<br/><b>Principal Signature</b>", cell_bold)]
    ]
    sig_table = Table(sig_data, colWidths=[180, 185, 180])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(sig_table)

    doc.build(story)


# ----------------- Streamlit UI -----------------
st.title("🎓 Smart Portfolio Generator")
st.caption("3-Page Automated Student Portfolio Builder (Bulk Photos & Single Photo Match)")

# Sidebar: File Management
st.sidebar.header("📁 Data Management")

# 1. Master Excel Upload
excel_file = st.sidebar.file_uploader("Upload Master Excel Sheet (.xlsx)", type=["xlsx"])
if excel_file:
    with open(SAVED_EXCEL_PATH, "wb") as f:
        f.write(excel_file.getbuffer())
    st.sidebar.success("✅ Master Excel saved permanently!")

# 2. Photos Bulk Upload (ZIP or multiple images)
st.sidebar.subheader("📸 Bulk Photo Upload")
bulk_zip = st.sidebar.file_uploader("Upload Photos (ZIP file)", type=["zip"],
                                    help="Photos should be named by Roll No (e.g., 1.jpg, 2.png)")
if bulk_zip:
    with zipfile.ZipFile(bulk_zip, 'r') as z:
        for file_info in z.infolist():
            ext = os.path.splitext(file_info.filename)[1].lower()
            if ext in ['.jpg', '.jpeg', '.png']:
                filename = os.path.basename(file_info.filename)
                if filename:
                    extracted_path = os.path.join(PHOTOS_DIR, filename)
                    with open(extracted_path, "wb") as f:
                        f.write(z.read(file_info))
    st.sidebar.success("✅ Bulk photos extracted and matched!")

# Load Data
if not os.path.exists(SAVED_EXCEL_PATH):
    st.info("👈 Please upload the Master Excel sheet from the sidebar to begin.")
    st.stop()


@st.cache_data
def load_data(path):
    df = pd.read_excel(path, header=0)
    return df


df = load_data(SAVED_EXCEL_PATH)


# Helper function to find photo by Roll No
def get_student_photo_path(roll_no):
    r_str = str(roll_no).strip()
    for ext in ['.jpg', '.jpeg', '.png', '.JPG', '.PNG']:
        candidate = os.path.join(PHOTOS_DIR, f"{r_str}{ext}")
        if os.path.exists(candidate):
            return candidate
    return None


tab1, tab2, tab3 = st.tabs(["👤 Single Student Portfolio", "📦 Bulk Generation (All)", "📷 Individual Photo Upload"])

# ----------------- TAB 1: Single Student -----------------
with tab1:
    st.subheader("Generate & Preview Student Portfolio")
    roll_col = df.columns[0]
    name_col = df.columns[9] if len(df.columns) > 9 else df.columns[1]

    student_options = {f"Roll: {row[roll_col]} - {row[name_col]}": idx for idx, row in df.iterrows()}
    selected_label = st.selectbox("Select Student", list(student_options.keys()))

    if selected_label:
        s_idx = student_options[selected_label]
        student_row = df.iloc[s_idx].tolist()
        roll_no = str(student_row[0]).strip()
        photo_path = get_student_photo_path(roll_no)

        col_info, col_img = st.columns([3, 1])
        with col_info:
            st.write(f"**Name:** {student_row[9] if len(student_row) > 9 else 'N/A'}")
            st.write(f"**Roll No:** {roll_no}")
            st.write(f"**Class:** {student_row[1] if len(student_row) > 1 else 'N/A'}")
        with col_img:
            if photo_path:
                st.image(photo_path, width=110, caption="Assigned Photo")
            else:
                st.warning("No photo found for this Roll No.")

        if st.button("📄 Generate 3-Page Portfolio"):
            pdf_name = f"Portfolio_Roll_{roll_no}.pdf"
            pdf_path = os.path.join(OUTPUT_DIR, pdf_name)
            generate_student_portfolio_pdf(student_row, photo_path, pdf_path)

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

            st.success("✅ Portfolio Generated Successfully!")
            st.download_button(
                label="⬇️ Download 3-Page PDF",
                data=pdf_bytes,
                file_name=pdf_name,
                mime="application/pdf"
            )

# ----------------- TAB 2: Bulk Generation -----------------
with tab2:
    st.subheader("Generate All Portfolios in One Click")
    st.write(f"Total students detected in Excel: **{len(df)}**")

    if st.button("🚀 Generate All Portfolios (ZIP)"):
        progress_bar = st.progress(0)
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for idx, row in df.iterrows():
                s_list = row.tolist()
                r_no = str(s_list[0]).strip()
                p_path = get_student_photo_path(r_no)
                out_path = os.path.join(OUTPUT_DIR, f"Portfolio_Roll_{r_no}.pdf")

                generate_student_portfolio_pdf(s_list, p_path, out_path)
                zip_file.write(out_path, arcname=f"Portfolio_Roll_{r_no}.pdf")
                progress_bar.progress((idx + 1) / len(df))

        st.success("✅ All 3-Page portfolios generated successfully!")
        st.download_button(
            label="⬇️ Download All Portfolios (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="All_Student_Portfolios.zip",
            mime="application/zip"
        )

# ----------------- TAB 3: Individual Photo Upload -----------------
with tab3:
    st.subheader("Upload or Replace Single Student Photo")
    single_roll = st.text_input("Enter Student Roll No:")
    single_photo = st.file_uploader("Select Photo (JPG / PNG)", type=["jpg", "jpeg", "png"], key="single_photo")

    if single_roll and single_photo:
        if st.button("Save Photo"):
            ext = os.path.splitext(single_photo.name)[1]
            dest = os.path.join(PHOTOS_DIR, f"{single_roll.strip()}{ext}")
            with open(dest, "wb") as f:
                f.write(single_photo.getbuffer())
            st.success(f"✅ Photo assigned to Roll No: {single_roll} permanently!")