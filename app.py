import streamlit as st
import fitz  # PyMuPDF
import xlsxwriter
from PIL import Image
import io
import os
import sys
from datetime import datetime

st.set_page_config(page_title="ID Card Extractor", page_icon="📇", layout="centered")

# --- CUSTOM CSS ---
st.markdown("""
    <style>
        .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #FF4B4B; color: white; }
        h1 { font-size: 2.2rem; }
    </style>
""", unsafe_allow_html=True)

st.title("📇 PMJAY ID Card Extractor")
st.write("Extract ID cards, pair them with a back-side image, and save to Excel.")

# --- 1. SETTINGS ---
# Set the target width for the images in Excel (in pixels).
# 400px is a good balance for visibility and file size.
TARGET_WIDTH_PX = 400 

def get_id_card_image(file_bytes):
    """Extracts the ID card and returns it as PIL Image."""
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        page = doc[0]
        # Your specific coordinates
        rect = fitz.Rect(106, 95, 410, 248)
        # High quality zoom (4x)
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=rect)
        img_data = pix.tobytes("png")
        doc.close()
        return Image.open(io.BytesIO(img_data))
    except Exception:
        return None

# --- 2. FILE UPLOADERS ---
# Load backside.png from the same folder
if getattr(sys, 'frozen', False):
    # Running as PyInstaller bundle
    base_path = sys._MEIPASS
else:
    # Running as script
    base_path = os.path.dirname(os.path.abspath(__file__))

back_image_path = os.path.join(base_path, "backside.png")
if os.path.exists(back_image_path):
    st.info("✅ Using backside.png from project folder")
else:
    st.error("❌ backside.png not found in project folder")
    st.stop()

uploaded_pdfs = st.file_uploader("📂 Upload PDF Files (Bulk)", type="pdf", accept_multiple_files=True)

# --- 3. PROCESSING ---
if uploaded_pdfs and st.button("🚀 Start Bulk Extraction"):
    
    # Prepare the Back Image
    back_img_original = Image.open(back_image_path)
    
    # Setup Output
    output_folder = os.path.join(os.getcwd(), "Extracted_ID_Reports")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = os.path.join(output_folder, f"IDs_{timestamp}.xlsx")
    
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet("IDs")
    
    # We need to process the first image to determine dimensions
    # Then we set the column widths ONCE for the whole sheet
    sample_img = get_id_card_image(uploaded_pdfs[0].read())
    uploaded_pdfs[0].seek(0) # Reset pointer
    
    if sample_img:
        # Calculate Scale to fit TARGET_WIDTH_PX
        w_orig, h_orig = sample_img.size
        scale_factor = TARGET_WIDTH_PX / w_orig
        target_height_px = int(h_orig * scale_factor)
        
        # Set Column Widths (A and B) to exactly TARGET_WIDTH_PX
        # xlsxwriter.set_column_pixels is available in newer versions
        worksheet.set_column_pixels(0, 1, TARGET_WIDTH_PX) 
        
        st.info(f"Target Cell Size: {TARGET_WIDTH_PX}x{target_height_px} px")
    else:
        st.error("Could not read first PDF to set dimensions.")
        st.stop()

    # Progress UI
    bar = st.progress(0)
    status = st.empty()
    
    for i, pdf_file in enumerate(uploaded_pdfs):
        status.text(f"Processing: {pdf_file.name}")
        
        # 1. Get Front Image
        front_pil = get_id_card_image(pdf_file.read())
        
        if front_pil:
            row = i  # 0-indexed, no header row as requested
            
            # 2. Resize Back Image to match Front Image exactly
            back_pil_resized = back_img_original.resize(front_pil.size)
            
            # 3. Convert to bytes for XlsxWriter
            front_bytes = io.BytesIO()
            front_pil.save(front_bytes, format="PNG")
            
            back_bytes = io.BytesIO()
            back_pil_resized.save(back_bytes, format="PNG")
            
            # 4. Set Row Height
            worksheet.set_row_pixels(row, target_height_px)
            
            # 5. Insert Images (Col A = Front, Col B = Back)
            # We use x_scale/y_scale to make the high-res image fit our calculated cell size
            opts = {
                'image_data': front_bytes, 
                'x_scale': scale_factor, 
                'y_scale': scale_factor,
                'object_position': 1 # Move and size with cells
            }
            worksheet.insert_image(row, 0, pdf_file.name, opts)
            
            opts['image_data'] = back_bytes
            worksheet.insert_image(row, 1, "backside.png", opts)
            
        bar.progress((i + 1) / len(uploaded_pdfs))
        
    workbook.close()
    bar.empty()
    status.empty()
    st.balloons()
    st.success("✅ Extraction Complete!")
    st.warning(f"📂 **File Saved:** `{excel_path}`")
    
    # --- DOWNLOAD EXCEL FILE ---
    with open(excel_path, "rb") as f:
        excel_data = f.read()
    
    st.download_button(
        label="📥 Download Excel File",
        data=excel_data,
        file_name=f"IDs_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if os.name == 'nt': # Windows only button
        if st.button("Open Folder"):
            os.startfile(output_folder)

# --- FOOTER ---
st.markdown("---")
st.markdown(
    "For any suggestions/improvements, contact Vivek Kumar Kamal on [Mail](mailto:vivek.creo@gmail.com) or [LinkedIn](https://www.linkedin.com/in/vivekkumarkamal/).",
    unsafe_allow_html=True
)