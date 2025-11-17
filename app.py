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

def split_products(paragraphs):
    """Split paragraphs into products starting with a 10-digit number."""
    products = []
    current_product = None
    current_paras = []

    for para in paragraphs:
        text = para.text.strip()
        if re.match(r'^\d{10}$', text):  # New product detected
            if current_product:
                html_content = paragraphs_to_html(current_paras)
                products.append((current_product, html_content))
            current_product = text
            current_paras = []
        else:
            current_paras.append(para)

    # Add last product
    if current_product:
        html_content = paragraphs_to_html(current_paras)
        products.append((current_product, html_content))

    return products

# --- Streamlit App ---
st.title("Word to HTML Excel Converter")
st.write("Upload a Word (.docx) file. Each product must start with a 10-digit number.")

uploaded_file = st.file_uploader("Choose a Word document", type="docx")

if uploaded_file is not None:
    try:
        doc = Document(uploaded_file)
        if not doc.paragraphs:
            st.error("The Word document is empty!")
        else:
            products = split_products(doc.paragraphs)
            if not products:
                st.error("No 10-digit product IDs found in the document.")
            else:
                # Create Excel in memory
                wb = Workbook()
                ws = wb.active
                ws.title = "HTML Export"
                ws.append(["ID", "HTML"])

                for code, html in products:
                    ws.append([code, html])

                output = BytesIO()
                wb.save(output)
                output.seek(0)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                excel_filename = f"{uploaded_file.name.replace('.docx','')}_{timestamp}.xlsx"

                st.success(f"✅ Conversion successful! {len(products)} products found.")
                st.download_button(
                    label="Download Excel",
                    data=output,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ Error: {e}")

