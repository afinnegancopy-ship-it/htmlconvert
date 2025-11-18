import streamlit as st
from docx import Document
from openpyxl import Workbook
import re
from datetime import datetime
from io import BytesIO
from docx.oxml.ns import qn

# ============================
# Helper functions (same logic as your script)
# ============================

def run_is_bold(run):
    if run.bold is True:
        return True
    if run.bold is None:
        rPr = run._element.rPr
        if rPr is not None and rPr.find(qn('w:b')) is not None:
            return True
    return False

def paragraph_is_bold(paragraph):
    if paragraph.style is not None:
        if paragraph.style.font.bold is True:
            return True
    return False

def is_bullet_paragraph(paragraph):
    style_name = paragraph.style.name.lower()
    if 'list' in style_name or 'bullet' in style_name or 'number' in style_name:
        return True
    if paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None:
        return True
    return False

def paragraph_to_html(paragraph):
    html = ""
    text = paragraph.text.strip()

    manual_bullet_match = re.match(r'^[\u2022\u00B7\-]\s+(.*)', text)
    is_bullet = is_bullet_paragraph(paragraph) or manual_bullet_match

    if is_bullet:
        html += "<li>"
        if manual_bullet_match:
            text = manual_bullet_match.group(1)
    else:
        html += "<p>"

    strong_phrases = [
        "Description:", "How To Use:", "Set Contains:", 
        "Key Notes:", "Fit & Fabric", "Product Details"
    ]

    for run in paragraph.runs:
        run_text = run.text
        if run_is_bold(run) or paragraph_is_bold(paragraph):
            run_text = f"<b>{run_text}</b>"
        for phrase in strong_phrases:
            if phrase in run_text:
                run_text = run_text.replace(phrase, f"<strong>{phrase}</strong>")
        html += run_text

    if is_bullet:
        html += "</li>"
    else:
        html += "</p>"

    return html

def docx_to_html_blocks(docx_file):
    doc = Document(docx_file)
    html_blocks = {}
    current_id = None
    current_html = []
    inside_list = False

    for para in doc.paragraphs:
        text = para.text.strip()
        if re.fullmatch(r'\d{8,}', text):
            if current_id and current_html:
                if inside_list:
                    current_html.append("</ul>")
                    inside_list = False
                html_blocks[current_id] = ''.join(current_html).strip()
                current_html = []
            current_id = text
        else:
            manual_bullet_match = re.match(r'^[\u2022\u00B7\-]\s+', text)
            is_bullet = is_bullet_paragraph(para) or manual_bullet_match

            if is_bullet and not inside_list:
                current_html.append("<ul>")
                inside_list = True
            elif not is_bullet and inside_list:
                current_html.append("</ul>")
                inside_list = False

            current_html.append(paragraph_to_html(para))

    if current_id and current_html:
        if inside_list:
            current_html.append("</ul>")
        html_blocks[current_id] = ''.join(current_html).strip()

    return html_blocks

def export_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "HTML Export"
    ws.append(["ID", "HTML"])

    for key, value in data.items():
        ws.append([key, value])

    # Save to BytesIO for Streamlit download
    excel_io = BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

# ============================
# Streamlit UI
# ============================

st.title("Word to HTML Excel Converter")
st.write("Upload a Word document (.docx) and convert its content into HTML blocks inside an Excel file.")

uploaded_file = st.file_uploader("Choose a Word (.docx) file", type=["docx"])

if uploaded_file:
    st.success(f"File uploaded: {uploaded_file.name}")
    if st.button("Convert to Excel"):
        with st.spinner("Converting..."):
            html_data = docx_to_html_blocks(uploaded_file)
            excel_file = export_to_excel(html_data)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{uploaded_file.name.split('.')[0]}_{timestamp}.xlsx"

        st.download_button(
            label="Download Excel",
            data=excel_file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Conversion complete!")
