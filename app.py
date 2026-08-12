import sys, os, subprocess

# Install missing libraries into the writable /tmp folder to bypass server permissions
custom_lib_dir = "/tmp/my_lib"
if custom_lib_dir not in sys.path:
    sys.path.insert(0, custom_lib_dir)

try:
    import docx
    import openpyxl
except ImportError:
    subprocess.check_call([
        sys.executable, "-m", "pip", "install", 
        "--target", custom_lib_dir, 
        "python-docx==1.1.0", "openpyxl"
    ])
    import docx
    import openpyxl

import streamlit as st
import zipfile, shutil
from engine import process_and_shuffle

st.set_page_config(page_title="Major Test Paper Generator", layout="centered")

st.title("🎓 Major Test Paper Shuffler Portal")
st.write("Upload your Excel question bank to generate randomized paper sets and EvalBee OMR Keys.")

# Initialize session state memory so outputs stay on screen
if "zip_data" not in st.session_state:
    st.session_state.zip_data = None
if "zip_name" not in st.session_state:
    st.session_state.zip_name = None

# Form Inputs
with st.form("exam_form"):
    uploaded_file = st.file_uploader("Upload Question Paper Excel File (.xlsx)", type=["xlsx"])
    
    exam_title = st.text_input("Exam Title", value="MAJOR TEST - 08 (CLASS XI)")
    
    col1, col2 = st.columns(2)
    with col1:
        time_limit = st.text_input("Time Limit", value="3 Hours")
    with col2:
        max_marks = st.number_input("Maximum Marks", value=300)
        
    num_sets = st.selectbox("Number of Shuffled Sets", options=[2, 4, 6], index=1)
    
    submit_button = st.form_submit_button("Generate & Package Test Papers")

# Process when submit button is clicked
if submit_button:
    if uploaded_file is None:
        st.error("Please upload an Excel file first.")
    else:
        with st.spinner("Shuffling questions, generating Word papers & EvalBee keys..."):
            work_dir = "temp_run"
            output_dir = os.path.join(work_dir, "output")
            os.makedirs(work_dir, exist_ok=True)
            
            excel_path = os.path.join(work_dir, uploaded_file.name)
            with open(excel_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            # Run Shuffling Process
            files = process_and_shuffle(excel_path, num_sets, exam_title, time_limit, max_marks, output_dir)
            
            # Zip everything
            zip_path = os.path.join(work_dir, "Exam_Paper_Package.zip")
            with zipfile.ZipFile(zip_path, 'w') as zf:
                for f in files:
                    zf.write(f, os.path.basename(f))
                    
            # Save file data to memory so it doesn't disappear on rerun
            with open(zip_path, "rb") as zf:
                st.session_state.zip_data = zf.read()
                st.session_state.zip_name = "Exam_Paper_Package.zip"

# Display download button if package exists in memory
if st.session_state.zip_data is not None:
    st.success("🎉 All paper sets and EvalBee Answer Key generated successfully!")
    st.download_button(
        label="⬇️ Download Exam Package (.zip)",
        data=st.session_state.zip_data,
        file_name=st.session_state.zip_name,
        mime="application/zip"
    )
