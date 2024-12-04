import io
import streamlit as st
import fitz  # PyMuPDF
import openpyxl
from io import BytesIO
import zipfile

# Define preset keywords categorized by Australian states
PRESET_KEYWORDS = {
    "VIC": ["VPA", "Victorian Planning Authority", "VPP", "Victorian Planning Provision"],
    "QLD": ["Shire Planning", "City Planning"],
    "WA": ["Metropolitan Region Scheme", "Rural Living Zone"],
    "NSW": ["Urban Renewal Areas"],
}

# Define general keywords including state-specific ones
GENERAL_KEYWORDS = [
    "Activity Centre", "Amendment", "Amendments Report", "Annual Plan", "Annual Report",
    "Area Plan", "Assessments", "Broadacre", "Budget", "City Plan", "Code Amendment",
    "Concept Plan", "Corporate Business Plan", "Corporate Plan", "Council Action Plan",
    "Council Business Plan", "Council Plan", "Council Report", 
    "Development Investigation Area", "Development Plan", "Development Plan Amendment",
    "DPA", "Emerging community", "Employment land study", "Exhibition", "expansion",
    "Framework", "Framework plan", "Gateway Determination", "greenfield", "growth area", 
    "growth plan", "growth plans", "housing", "Housing Strategy",
    "Industrial land study", "infrastructure plan", "infrastructure planning", 
    "Inquiries", "Investigation area", "land use", "Land use strategy",
    "LDP", "Local Area Plan", "Local Development Area", "Local Development Plan",
    "Local Environmental Plan", "Local Planning Policy", "Local Planning Scheme",
    "Local Planning Strategy", "Local Strategic Planning Statement", "LPP", "LPS", "LSPS",
    "Major Amendment", "Major Update", "Master Plan", "Masterplan", "Neighbourhood Plan",
    "Operational Plan", "Planning Commission", "Planning Framework", "Planning Investigation Area",
    "Planning proposal", "Planning report", "Planning Scheme", "Planning Scheme Amendment",
    "Planning Strategy", "Precinct plan", "Priority Development Area",
    "Project Vision", "Rezoning", "settlement", "Strategy", "Structure Plan", "Structure Planning",
    "Study", "Territory plan", "Town Planning Scheme",
    "Township Plan", "TPS", "Urban Design Framework", "Urban growth", "Urban Release",
    "Urban renewal", "Variation", "Vision"
]

# Combine all keywords for easy access
ALL_KEYWORDS = {**PRESET_KEYWORDS, "General": GENERAL_KEYWORDS}

# Initialize session state
if 'updated_pdfs' not in st.session_state:
    st.session_state.updated_pdfs = {}
if 'csv_reports' not in st.session_state:
    st.session_state.csv_reports = {}
if 'selected_keywords' not in st.session_state:
    st.session_state.selected_keywords = set()

# Function to validate PDF files
def is_valid_pdf(file):
    try:
        file.seek(0)
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            if doc.is_encrypted:
                st.error(f"⚠️ {file.name} is encrypted. Please provide an unencrypted PDF.")
                return False
            return True
    except Exception as e:
        st.error(f"⚠️ {file.name} is not a valid PDF file. Error: {e}")
        return False

# Callback function for "Select All Keywords" checkbox
def select_all_callback():
    if st.session_state.select_all_keywords:
        # Add all keywords to selected_keywords
        st.session_state.selected_keywords = set([kw for kws in ALL_KEYWORDS.values() for kw in kws])
    else:
        # Remove all keywords from selected_keywords
        st.session_state.selected_keywords = set()

# Callback function for State checkboxes
def toggle_state_callback(state):
    state_key = f'state_{state}'
    if st.session_state[state_key]:
        # Add associated keywords
        st.session_state.selected_keywords.update(PRESET_KEYWORDS[state])
    else:
        # Remove associated keywords
        st.session_state.selected_keywords.difference_update(PRESET_KEYWORDS[state])

# Function to highlight text in PDF and return the updated PDF and keyword occurrences
def highlight_text_in_pdf(file_content, selected_keywords, original_filename):
    try:
        pdf_file = io.BytesIO(file_content)
        pdf_document = fitz.open(stream=pdf_file, filetype="pdf")
    except Exception as e:
        st.error(f"Error opening {original_filename}: {e}")
        return None, None

    # Initialize keyword_occurrences with all selected keywords
    keyword_occurrences = {keyword: [] for keyword in selected_keywords}
    keywords_found = False

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("dict")

        for keyword in selected_keywords:
            keyword_lower = keyword.lower()

            for block in text["blocks"]:
                if block["type"] == 0:  # Block is text
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text_content = span["text"]
                            lower_text = text_content.lower()

                            start = 0
                            while True:
                                start = lower_text.find(keyword_lower, start)
                                if start == -1:
                                    break

                                # Track page number for each keyword occurrence
                                if (page_num + 1) not in keyword_occurrences[keyword]:
                                    keyword_occurrences[keyword].append(page_num + 1)
                                    keywords_found = True  # At least one keyword found

                                # Highlight the keyword in the PDF
                                bbox = span["bbox"]
                                span_width = bbox[2] - bbox[0]
                                span_height = bbox[3] - bbox[1]
                                char_width = span_width / len(text_content) if len(text_content) > 0 else 1

                                keyword_bbox = fitz.Rect(
                                    bbox[0] + char_width * start,
                                    bbox[1],
                                    bbox[0] + char_width * (start + len(keyword)),
                                    bbox[3]
                                )
                                
                                keyword_bbox = keyword_bbox.intersect(fitz.Rect(0, 0, page.rect.width, page.rect.height))
                                
                                if not keyword_bbox.is_empty:
                                    highlight = page.add_highlight_annot(keyword_bbox)
                                    highlight.set_colors(stroke=(1, 0.65, 0))  # Set color to orange
                                    highlight.update()

                                start += len(keyword)

    # Save the highlighted PDF
    output_pdf = BytesIO()
    pdf_document.save(output_pdf)
    output_pdf.seek(0)

    # Close the PDF document to free resources
    pdf_document.close()

    if not keywords_found:
        return output_pdf, None  # Return the updated PDF even if no keywords are found

    return output_pdf, keyword_occurrences

# Function to generate CSV report from keyword occurrences
def generate_csv_report(keyword_occurrences):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Keywords Report"

    # Write header
    ws.append(["Keyword", "Occurrences (Page Numbers)"])

    # Write keyword occurrences
    for keyword, pages in keyword_occurrences.items():
        if pages:
            ws.append([keyword, ", ".join(map(str, pages))])

    # Auto-size columns
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length + 2

    # Save to BytesIO
    excel_output = BytesIO()
    wb.save(excel_output)
    excel_output.seek(0)

    return excel_output

# Function to hide default file uploader instructions using custom CSS
def hide_file_uploader_instructions():
    hide_ui = """
                <style>
                /* Hide the default drag and drop instructions and size limit */
                div[data-testid="stFileUploadDropzone"] > div > div > div > p {
                    display: none;
                }
                div[data-testid="stFileUploadDropzone"] > div > div > div > span {
                    display: none;
                }
                </style>
                """
    st.markdown(hide_ui, unsafe_allow_html=True)

# Main tool interface
def keyword_highlighter_page():
    st.title("📄 PDF Keyword Highlighter")

    st.write("📂 **Instructions:**")
    st.write("- **Upload PDFs:** Click the upload button and select multiple PDF files by holding `Ctrl` or `Shift`.")
    st.write("- **Select Keywords:** Choose from the predefined keywords or add your own.")
    st.write("- **Process:** Click the 'Highlight Keywords' button to start processing.")

    MAX_FILES = 20  # Maximum number of files
    MAX_TOTAL_SIZE_MB = 5000  # Maximum total upload size in MB

    # Hide default file uploader instructions
    hide_file_uploader_instructions()

    uploaded_files = st.file_uploader(
        f"📂 Choose up to {MAX_FILES} PDF files (Total size up to {MAX_TOTAL_SIZE_MB} MB)",
        type="pdf",
        accept_multiple_files=True
    )

    if uploaded_files:
        # Enforce the maximum number of files
        if len(uploaded_files) > MAX_FILES:
            st.error(f"⚠️ You can upload a maximum of {MAX_FILES} files at once.")
            uploaded_files = uploaded_files[:MAX_FILES]

        # Calculate the total size
        total_size = sum(file.size for file in uploaded_files) / (1024 * 1024)  # Convert to MB
        if total_size > MAX_TOTAL_SIZE_MB:
            st.error(f"⚠️ The total upload size exceeds {MAX_TOTAL_SIZE_MB} MB. Please reduce the number or size of files.")
            # Trim the list to fit the size limit
            allowed_size = 0
            valid_files = []
            for file in uploaded_files:
                file_size_mb = file.size / (1024 * 1024)
                if allowed_size + file_size_mb <= MAX_TOTAL_SIZE_MB:
                    valid_files.append(file)
                    allowed_size += file_size_mb
                else:
                    st.warning(f"⚠️ {file.name} exceeds the remaining upload size limit and was skipped.")
            uploaded_files = valid_files

        # Display uploaded files
        st.write(f"📥 **Uploaded {len(uploaded_files)} files:**")
        for uploaded_file in uploaded_files:
            st.write(f"- {uploaded_file.name} ({uploaded_file.size / (1024 * 1024):.2f} MB)")

        st.subheader("🔍 Select Keywords to Highlight")

        # Determine if all keywords are selected
        all_keywords_set = set([kw for kws in ALL_KEYWORDS.values() for kw in kws])
        select_all_checked = all_keywords_set.issubset(st.session_state.selected_keywords)

        # Add a "Select All" checkbox with callback
        select_all = st.checkbox("✅ Select All Keywords", value=select_all_checked, key="select_all_keywords", on_change=select_all_callback)

        # Display state tickboxes under "States:" sub-section with callbacks
        st.markdown("### States:")
        for state, keywords in PRESET_KEYWORDS.items():
            state_key = f'state_{state}'
            # Determine if all keywords for the state are selected
            state_keywords_set = set(keywords)
            state_checked = state_keywords_set.issubset(st.session_state.selected_keywords)
            is_checked = st.checkbox(f"✅ {state}", value=state_checked, key=state_key, on_change=toggle_state_callback, args=(state,))

        # General category (includes state-specific keywords) distributed across 4 columns
        with st.expander("### General Keywords", expanded=False):
            col1, col2, col3, col4 = st.columns(4)
            columns = [col1, col2, col3, col4]
            for i, keyword in enumerate(GENERAL_KEYWORDS + [kw for kws in PRESET_KEYWORDS.values() for kw in kws]):
                col = columns[i % 4]
                checkbox_key = f"General_{keyword}"
                is_checked = keyword in st.session_state.selected_keywords
                if col.checkbox(keyword, value=is_checked, key=checkbox_key):
                    st.session_state.selected_keywords.add(keyword)
                else:
                    st.session_state.selected_keywords.discard(keyword)

        # Custom keyword addition
        custom_keywords = st.text_area("✏️ Or add your own keywords (one per line):", "")
        if custom_keywords:
            custom_keywords_list = [kw.strip() for kw in custom_keywords.split('\n') if kw.strip()]
            st.session_state.selected_keywords.update(custom_keywords_list)

        # Add a checkbox for optional CSV report
        generate_csv = st.checkbox("📊 Generate CSV Report", value=False, key="generate_csv_report")

        if st.button("🚀 Highlight Keywords"):
            if not st.session_state.selected_keywords:
                st.error("⚠️ Please select or add at least one keyword.")
            else:
                # Validate uploaded files
                valid_files = []
                for file in uploaded_files:
                    # Reset the file pointer before validation
                    file.seek(0)
                    if is_valid_pdf(file):
                        valid_files.append(file)
                    else:
                        st.error(f"⚠️ {file.name} is not a valid PDF file.")
                if not valid_files:
                    st.error("⚠️ No valid PDF files to process.")
                else:
                    # Clear previous results
                    st.session_state.updated_pdfs = {}
                    st.session_state.csv_reports = {}
                    
                    total_files = len(valid_files)
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for idx, uploaded_file in enumerate(valid_files):
                        # Reset the file pointer before reading
                        uploaded_file.seek(0)
                        
                        # Read the file content once
                        file_content = uploaded_file.read()
                        if not file_content:
                            st.error(f"⚠️ {uploaded_file.name} is empty after reading.")
                            continue
                        
                        # Update status text
                        status_text.text(f"Processing file {idx + 1} of {total_files}: {uploaded_file.name}")
                        
                        # Process each PDF
                        updated_pdf, keyword_occurrences = highlight_text_in_pdf(
                            file_content, st.session_state.selected_keywords, uploaded_file.name
                        )

                        if not keyword_occurrences:
                            st.warning(f"No keywords found in **{uploaded_file.name}**.")
                            continue

                        # Store the updated PDF in session state
                        st.session_state.updated_pdfs[uploaded_file.name] = updated_pdf

                        # Generate CSV report if checkbox is selected
                        if generate_csv:
                            csv_report = generate_csv_report(keyword_occurrences)
                            st.session_state.csv_reports[uploaded_file.name] = csv_report

                        # Update progress bar
                        progress = (idx + 1) / total_files
                        progress_bar.progress(progress)

# Download section outside the main interface to persist across reruns
def download_section():
    if st.session_state.updated_pdfs:
        st.success("✅ Processing complete!")
        num_pdfs = len(st.session_state.updated_pdfs)
        
        if num_pdfs > 1:
            # More than one PDF, provide "Download All PDFs as ZIP"
            st.write("📥 **Download All Updated PDFs:**")
            # Create a zip file of all updated PDFs
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                for filename, pdf in st.session_state.updated_pdfs.items():
                    zip_file.writestr(f"highlighted_{filename}", pdf.getvalue())
            zip_buffer.seek(0)
            st.download_button(
                label="📄 Download All PDFs as ZIP",
                data=zip_buffer,
                file_name="highlighted_pdfs.zip",
                mime="application/zip",
                key="download_all_pdfs"
            )
        else:
            # Only one PDF, provide individual download button
            st.write("📥 **Download Updated PDF:**")
            for filename, pdf in st.session_state.updated_pdfs.items():
                # Provide a download button for the updated PDF
                st.download_button(
                    label=f"📄 Download {filename}",
                    data=pdf,
                    file_name=f"highlighted_{filename}",
                    mime="application/pdf",
                    key=f"download_pdf_{filename}"
                )
        
        # CSV Reports remain the same
        if st.session_state.csv_reports:
            if len(st.session_state.csv_reports) > 1:
                st.write("📊 **Download All CSV Reports:**")
                # Multiple reports, provide a ZIP file
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                    for filename, csv_report in st.session_state.csv_reports.items():
                        report_filename = f"keywords_report_{filename.replace('.pdf', '.xlsx')}"
                        zip_file.writestr(report_filename, csv_report.getvalue())
                zip_buffer.seek(0)
                st.download_button(
                    label="📄 Download All Reports as ZIP",
                    data=zip_buffer,
                    file_name="keywords_reports.zip",
                    mime="application/zip",
                    key="download_all_reports"
                )
            else:
                st.write("📊 **Download CSV Report:**")
                # Single report, provide individual download button
                for filename, csv_report in st.session_state.csv_reports.items():
                    st.download_button(
                        label=f"📄 Download {filename} Report",
                        data=csv_report,
                        file_name=f"keywords_report_{filename.replace('.pdf', '.xlsx')}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_csv_{filename}"
                    )


# Main function
def main():
    keyword_highlighter_page()
    download_section()

if __name__ == "__main__":
    main()
