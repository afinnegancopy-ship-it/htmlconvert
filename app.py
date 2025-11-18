import streamlit as st
from docx import Document
from openpyxl import Workbook
import re
from datetime import datetime
from io import BytesIO

# === FUNCTIONS ===
def paragraph_to_html(paragraph):
    html = ""
    text = paragraph.text.strip()
    manual_bullet_match = re.match(r'^[\u2022\u00B7\-]\s+(.*)', text)
    
    if paragraph.style.name.startswith('List') or manual_bullet_match:
        html += "<li>"
        if manual_bullet_match:
            text = manual_bullet_match.group(1)
    else:
        html += "<p>"

    for run in paragraph.runs:
        run_text = run.text
        if run.bold:
            run_text = f"<b>{run_text}</b>"
        html += run_text

    if paragraph.style.name.startswith('List') or manual_bullet_match:
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
        # Match a ten-digit number as product ID
        if re.fullmatch(r'\d{10}', text):
            if current_id and current_html:
                if inside_list:
                    current_html.append("</ul>")
                    inside_list = False
                html_blocks[current_id] = ''.join(current_html).strip()
                current_html = []
            current_id = text
        else:
            manual_bullet_match = re.match(r'^[\u2022\u00B7\-]\s+', text)
            is_bullet = para.style.name.startswith('List') or manual_bullet_match

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

    # Save workbook to BytesIO for Streamlit download
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# === STREAMLIT APP ===
st.title("Word to HTML Exporter")
st.write("Upload a Word (.docx) file to convert its contents into HTML with proper <b> and <li> tags, separated by 10-digit product IDs.")

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

if uploaded_file:
    st.info("Processing your document...")
    data = docx_to_html_blocks(uploaded_file)
    
    if data:
        st.success(f"Found {len(data)} product IDs!")
        excel_data = export_to_excel(data)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"html_export_{timestamp}.xlsx"
        st.download_button(
            label="Download Excel file",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No product IDs found. Make sure your Word document contains 10-digit numbers as IDs.")
