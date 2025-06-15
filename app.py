import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO
from datetime import datetime
from fpdf import FPDF

# Page configuration
st.set_page_config(
    page_title="Crush Medication Review",
    page_icon="💊",
    layout="wide"
)

# Load medication data
@st.cache_data
def load_data():
    if os.path.exists("data/Crush_Med_Data_Bank_Clean.xlsx"):
        return pd.read_excel("data/Crush_Med_Data_Bank_Clean.xlsx")
    else:
        uploaded_file = st.file_uploader("Upload medication database (Crush_Med_Data_Bank_Clean.xlsx)", type="xlsx")
        if uploaded_file:
            return pd.read_excel(uploaded_file)
        else:
            st.warning("Please upload the medication database file")
            st.stop()

# Generate Word report
def generate_word_report(patient_name, dob, selected_meds, med_data):
    doc = Document()
    
    # Header
    doc.add_heading('Medication Crushability Review Report', level=1)
    doc.add_paragraph(f"**Patient:** {patient_name}\t**Date of Birth:** {dob}")
    doc.add_paragraph(f"**Report Date:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    # Important note
    note = doc.add_paragraph()
    note.add_run("IMPORTANT: Crushing tablets renders the medication unlicensed. If a liquid formulation is available, this is always the preferred option.").bold = True
    
    # Medications table
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    # Table headers
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Medication'
    hdr_cells[1].text = 'Can Be Crushed?'
    hdr_cells[2].text = 'Alternative Form'
    hdr_cells[3].text = 'Recommendation'
    
    # Table content
    for med in selected_meds:
        med_info = med_data[med_data['Drug'] == med].iloc[0]
        row_cells = table.add_row().cells
        row_cells[0].text = str(med_info['Drug'])
        row_cells[1].text = str(med_info['Can be Crushed'])
        row_cells[2].text = str(med_info['Alternative form available?'])
        row_cells[3].text = str(med_info['Recommendation'])
    
    # Professional attribution (footer)
    footer = doc.add_paragraph()
    footer.alignment = 2  # Right alignment
    footer_run = footer.add_run("Crush Med Review Tool developed by: Célio Lara Júnior - Pharmacist - PSI 10003245")
    footer_run.italic = True
    footer_run.font.size = Pt(9)
    
    return doc

# Generate PDF report
def generate_pdf_report(patient_name, dob, selected_meds, med_data):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Header
    pdf.cell(200, 10, txt="Medication Crushability Review Report", ln=1, align='C')
    pdf.cell(200, 10, txt=f"Patient: {patient_name}", ln=1)
    pdf.cell(200, 10, txt=f"Date of Birth: {dob}", ln=1)
    pdf.cell(200, 10, txt=f"Report Date: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=1)
    
    # Important note
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(200, 10, txt="IMPORTANT: Crushing tablets renders the medication unlicensed. If a liquid formulation is", ln=1)
    pdf.cell(200, 10, txt="available, this is always the preferred option.", ln=1)
    pdf.ln(5)
    
    # Medications table
    pdf.set_font("Arial", size=10)
    col_widths = [50, 30, 50, 70]  # Column widths
    
    # Table headers
    headers = ["Medication", "Crushable?", "Alternative", "Recommendation"]
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, txt=header, border=1)
    pdf.ln()
    
    # Table content
    for med in selected_meds:
        med_info = med_data[med_data['Drug'] == med].iloc[0]
        data = [
            str(med_info['Drug']),
            str(med_info['Can be Crushed']),
            str(med_info['Alternative form available?']),
            str(med_info['Recommendation'])
        ]
        for i, item in enumerate(data):
            pdf.cell(col_widths[i], 10, txt=item, border=1)
        pdf.ln()
    
    # Professional attribution (footer)
    pdf.set_y(-20)
    pdf.set_font('Arial', 'I', 8)
    pdf.cell(0, 10, txt="Crush Med Review Tool developed by: Célio Lara Júnior - Pharmacist - PSI 10003245", ln=1, align='R')
    
    return pdf.output(dest='S').encode('latin1')

# Main app interface
def main():
    st.title("💊 Medication Crushability Review")
    st.markdown("""
    This application generates medication review reports for patients with swallowing difficulties,
    maintaining the standard template required for Irish healthcare documentation.
    """)
    
    # Load data
    med_data = load_data()
    med_list = med_data['Drug'].dropna().unique().tolist()
    
    # Patient information
    with st.expander("Patient Information", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            patient_name = st.text_input("Full Name*")
        with col2:
            dob = st.date_input("Date of Birth*")
    
    # Medication selection
    with st.expander("Medication Selection", expanded=True):
        search_term = st.text_input("Search medication:")
        
        if search_term:
            filtered_meds = [m for m in med_list if search_term.lower() in str(m).lower()]
        else:
            filtered_meds = med_list
        
        selected_meds = st.multiselect(
            "Select medications for review:",
            filtered_meds,
            help="Type to filter the medication list"
        )
        
        if selected_meds:
            st.dataframe(
                med_data[med_data['Drug'].isin(selected_meds)].reset_index(drop=True),
                use_container_width=True
            )
    
    # Generate reports
    if st.button("Generate Full Report"):
        if not patient_name or not dob or not selected_meds:
            st.error("Please complete all required fields (*)")
        else:
            # Generate Word report
            doc = generate_word_report(patient_name, dob.strftime('%d/%m/%Y'), selected_meds, med_data)
            word_buffer = BytesIO()
            doc.save(word_buffer)
            
            # Generate PDF report
            pdf_bytes = generate_pdf_report(patient_name, dob.strftime('%d/%m/%Y'), selected_meds, med_data)
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="⬇️ Download Word Report",
                    data=word_buffer.getvalue(),
                    file_name=f"Crushability_Review_{patient_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with col2:
                st.download_button(
                    label="⬇️ Download PDF Report",
                    data=pdf_bytes,
                    file_name=f"Crushability_Review_{patient_name.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )

if __name__ == "__main__":
    main()
