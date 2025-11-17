import streamlit as st
from docx import Document
from openpyxl import Workbook
from io import BytesIO
import re
from datetime import datetime

# --- HTML Conversion Functions ---
def paragraph_to_html(paragraph):
    """Convert a paragraph to HTML, preserving bold formatting."""
    html = ""
    for run in paragraph.runs:
        run_text = run.text
        if run.bold:
            run_text = f"<b>{run_text}</b>"
        html += run_text
    return html

def paragraphs_to_html(paragraphs):
    """Convert a list of paragraphs to HTML with merged non-bullet paragraphs and proper <ul><li> for bullets."""
    html_content = []
    inside_list = False
    pending_paragraphs = []

    for para in paragraphs:
        text = para.text.strip()
        is_list_item = para.style.name.startswith('List') or re.match(r'^[•\-\*]\s+', text)

        if is_list_item:
            # Flush pending non-bullet paragraphs
            if pending_paragraphs:
                merged = "<br>".join(pending_paragraphs)
                html_content.append(f"<p>{merged}</p>")
                pending_paragraphs = []

            if not inside_list:
                html_content.append("<ul>")
                inside_list = True

            # Clean bullet character
            item_html = ""
            for run in para.runs:
                run_text = run.text
                run_text = run_text.replace("•", "").replace("-", "").replace("*", "").strip()
                if run.bold:
                    run_text = f"<b>{run_text}</b>"
                item_html += run_text
            html_content.append(f"<li>{item_html}</li>")

        else:
            if inside_list:
                html_content.append("</ul>")
                inside_list = False
            pending_paragraphs.append(paragraph_to_html(para))

    # Flush remaining paragraphs
    if pending_paragraphs:
        merged = "<br>".join(pending_paragraphs)
        html_content.append(f"<p>{merged}</p>")
    if inside_list:
        html_content.append("</ul>")

    return ''.join(html_content)

# --- Streamlit App ---
st.title("Word to HTML Excel Converter")
st.write("Upload a Word (.docx) file to convert its content to HTML and export as Excel.")

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

if uploaded_file is not None:
    try:
        # Load Word document from uploaded file
        doc = Document(uploaded_file)
        if not doc.paragraphs:
            st.error("The Word document is empty!")
        else:
            # First paragraph = code
            code = doc.paragraphs[0].text.strip()
            # Rest = HTML content
            html_content = paragraphs_to_html(doc.paragraphs[1:])

            # Create Excel in memory
            wb = Workbook()
            ws = wb.active
            ws.title = "HTML Export"
            ws.append(["ID", "HTML"])
            ws.append([code, html_content])

            output = BytesIO()
            wb.save(output)
            output.seek(0)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_filename = f"{uploaded_file.name.replace('.docx','')}_{timestamp}.xlsx"

            st.success("✅ Conversion successful!")
            st.download_button(
                label="Download Excel",
                data=output,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Error: {e}")
