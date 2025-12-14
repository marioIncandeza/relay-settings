import streamlit as st
import os
import shutil
import tempfile
import zipfile
import pythoncom  # Required for xlwings in some threaded environments
from rdb import gen_settings, RELAY_CONFIG, RELAY_REGION_METADATA

st.set_page_config(page_title="SEL Settings Generator", layout="wide")

def create_zip(source_dir, output_filename):
    """Zips the contents of source_dir into output_filename"""
    with zipfile.ZipFile(output_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Archive name should be relative to source_dir so we don't include full path
                arcname = os.path.relpath(file_path, source_dir)
                zf.write(file_path, arcname)

def main():
    st.title("SEL Settings Generator")
    st.markdown("Generate RDB files from Excel settings tables.")

    # --- Sidebar Configuration ---
    st.sidebar.header("Configuration")
    
    # Relay Type Selection
    item_labels = {k: v['label'] for k, v in RELAY_CONFIG.items()}
    selected_key = st.sidebar.selectbox(
        "Select Relay Type",
        options=list(RELAY_CONFIG.keys()),
        format_func=lambda x: item_labels[x]
    )
    
    relay_config = RELAY_CONFIG[selected_key]
    
    # Region Selection (if applicable)
    excluded_regions = []
    if selected_key in RELAY_REGION_METADATA:
        st.sidebar.subheader("Regions")
        region_meta = RELAY_REGION_METADATA[selected_key]
        
        # Region Controls
        col1, col2 = st.sidebar.columns(2)
        all_on = col1.button("Select All")
        all_off = col2.button("Deselect All")
        
        # We use session state to track region toggles to allow buttons to work
        # Initialize state if needed or if key changed
        if "region_states" not in st.session_state or st.session_state.get("last_key") != selected_key:
            st.session_state.region_states = {label: True for label in region_meta["labels"]}
            st.session_state.last_key = selected_key
            
        if all_on:
            for k in st.session_state.region_states: st.session_state.region_states[k] = True
        if all_off:
            for k in st.session_state.region_states: st.session_state.region_states[k] = False
            
        # Draw checkboxes
        for label in region_meta["labels"]:
            st.session_state.region_states[label] = st.sidebar.checkbox(
                label, 
                value=st.session_state.region_states[label]
            )
            
        # Compile excluded list
        excluded_regions = [
            region_meta["shorthand"][label]
            for label, checked in st.session_state.region_states.items() 
            if not checked
        ]

    # Comment Toggle
    include_comments = st.sidebar.checkbox("Include Comments in RDB", value=True)

    # --- Main Area ---
    
    # 1. Excel Input
    st.subheader("1. Input Settings")
    uploaded_excel = st.file_uploader("Upload Excel Settings File (.xlsx)", type=["xlsx", "xls"])

    # 2. Template Selection
    st.subheader("2. Template Source")
    template_option = st.radio(
        "Choose Template Source:",
        ("Embedded Template", "Custom Template (Zip Upload)", "Local Path (Advanced)")
    )
    
    template_path_final = None
    temp_template_dir = None  # Handle for cleanup if we extract a zip
    
    if template_option == "Embedded Template":
        # Assumes templates are stored in 'templates/<key>' relative to this script
        # Check if directory exists
        potential_path = os.path.join("templates", selected_key)
        if os.path.exists(potential_path):
            st.info(f"Using embedded template for {relay_config['label']}")
            template_path_final = os.path.abspath(potential_path)
        else:
            st.error(f"Embedded template not found at {potential_path}. Please populate the directory.")
            
    elif template_option == "Custom Template (Zip Upload)":
        uploaded_zip = st.file_uploader("Upload Template Zip", type="zip")
        if uploaded_zip:
            # We will extract this later during generation to keep it temporary
            pass
            
    elif template_option == "Local Path (Advanced)":
        local_path = st.text_input("Enter absolute path to template directory on server/local machine")
        if local_path:
            template_path_final = local_path

    # 3. Generate
    st.subheader("3. Generate")
    
    if st.button("Generate Settings", type="primary"):
        if not uploaded_excel:
            st.error("Please upload an Excel file.")
            return
            
        if template_option == "Custom Template (Zip Upload)" and not uploaded_zip:
            st.error("Please upload a Template Zip file.")
            return
            
        if template_option == "Local Path (Advanced)" and not template_path_final:
            st.error("Please specify a local path.")
            return

        # Execution Block
        with st.status("Processing...", expanded=True) as status:
            try:
                # Create a master temporary directory for this generic run
                with tempfile.TemporaryDirectory() as master_temp:
                    
                    # A. Handle Excel File
                    excel_path = os.path.join(master_temp, "input_settings.xlsx")
                    with open(excel_path, "wb") as f:
                        f.write(uploaded_excel.getbuffer())
                    
                    # B. Handle Template (if Zip)
                    if template_option == "Custom Template (Zip Upload)":
                        zip_extract_path = os.path.join(master_temp, "extracted_template")
                        with zipfile.ZipFile(uploaded_zip, 'r') as z:
                            z.extractall(zip_extract_path)
                        
                        # Check implementation: sometimes zips contain a root folder. 
                        # We need to find the folder that contains .txt files or passed as root.
                        # Logic: If the extracted content is a single folder, drill down.
                        extracted_items = os.listdir(zip_extract_path)
                        # Filter out common junk
                        valid_items = [i for i in extracted_items if not i.startswith('__') and not i.startswith('.')]
                        
                        if len(valid_items) == 1:
                            potential_root = os.path.join(zip_extract_path, valid_items[0])
                            if os.path.isdir(potential_root):
                                template_path_final = potential_root
                            else:
                                template_path_final = zip_extract_path
                        else:
                            template_path_final = zip_extract_path
                    
                    # C. Define Output Dir
                    output_dir = os.path.join(master_temp, "output_files")
                    os.makedirs(output_dir, exist_ok=True)
                    
                    # D. Run Generation
                    st.write("Initializing Excel engine...")
                    # pythoncom.CoInitialize() # May be needed for threaded contexts
                    
                    st.write("Generating files...")
                    gen_settings(
                        xl_path=excel_path,
                        template_path=template_path_final,
                        output_path=output_dir,
                        workbook_params=relay_config['params'],
                        excluded_regions=excluded_regions,
                        include_comments=include_comments
                    )
                    
                    st.write("Zipping results...")
                    zip_output_path = os.path.join(master_temp, "settings_package.zip")
                    create_zip(output_dir, zip_output_path)
                    
                    # E. Read Zip for Download
                    with open(zip_output_path, "rb") as f:
                        zip_data = f.read()
                        
                    status.update(label="Generation Complete!", state="complete", expanded=False)
                    
                    st.success("Settings generated successfully!")
                    st.download_button(
                        label="Download Generated Settings (ZIP)",
                        data=zip_data,
                        file_name=f"{selected_key}_settings.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                # st.exception(e) # Uncomment for stack trace

if __name__ == "__main__":
    main()
