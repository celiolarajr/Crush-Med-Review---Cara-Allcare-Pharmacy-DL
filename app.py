import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO
from datetime import datetime, date
from fpdf import FPDF
import os

def clean_text(text):
    if not isinstance(text, str):
        text = str(text)
    replacements = {
        "–": "-", "—": "-", "’": "'", "‘": "'",
        "“": '"', "”": '"', "é": "e", "á": "a",
        "í": "i", "ó": "o", "ú": "u", "â": "a",
        "ê": "e", "î": "i", "ô": "o", "û": "u",
        "ã": "a", "õ": "o", "ç": "c", "É": "E",
        "Á": "A", "Í": "I", "Ó": "O", "Ú": "U",
        "Â": "A", "Ê": "E", "Î": "I", "Ô": "O",
        "Û": "U", "Ã": "A", "Õ": "O", "Ç": "C"
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text

st.set_page_config(
    page_title="Crush Medication Review",
    page_icon="💊",
    layout="wide"
)

@st.cache_data
def load_data():
    if os.path.exists("Crush_Med_Data_Bank_Clean.xlsx"):
        return pd.read_excel("Crush_Med_Data_Bank_Clean.xlsx")
    elif os.path.exists("data/Crush_Med_Data_Bank_Clean.xlsx"):
        return pd.read_excel("data/Crush_Med_Data_Bank_Clean.xlsx")
    else:
        uploaded_file = st.file_uploader("Upload medication database (Crush_Med_Data_Bank_Clean.xlsx)", type="xlsx")
        if uploaded_file:
            return pd.read_excel(uploaded_file)
        else:
            st.warning("Please upload the medication database file")
            st.stop()

def generate_word_report(patient_name, dob, selected_meds, med_data):
    doc = Document()
    
    # Formato DD/MM/YYYY
    dob_formatted = dob if isinstance(dob, str) else dob.strftime('%d/%m/%Y')
    report_date = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    doc.add_heading('Medication Crushability Review Report', level=1)
    doc.add_paragraph(f"**Patient:** {patient_name}\t**Date of Birth:** {dob_formatted}")
    doc.add_paragraph(f"**Report Date:** {report_date}")
    
    note = doc.add_paragraph()
    note.add_run("IMPORTANT: Crushing tablets renders the medication unlicensed. If a liquid formulation is available, this is always the preferred option.").bold = True
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Medication'
    hdr_cells[1].text = 'Can Be Crushed?'
    hdr_cells[2].text = 'Alternative Form'
    hdr_cells[3].text = 'Recommendation'
    
    for med in selected_meds:
        med_info = med_data[med_data['Drug'] == med].iloc[0]
        row_cells = table.add_row().cells
        row_cells[0].text = str(med_info['Drug'])
        row_cells[1].text = str(med_info['Can be Crushed'])
        row_cells[2].text = str(med_info['Alternative form available?'])
        row_cells[3].text = str(med_info['Recommendation'])
    
    footer = doc.add_paragraph()
    footer.alignment = 2
    footer_run = footer.add_run("Crush Med Review Tool developed by Célio Lara Júnior - Pharmacist - PSI 10003245")
    footer_run.italic = True
    footer_run.font.size = Pt(9)
    
    return doc

def generate_pdf_report(patient_name, dob, selected_meds, med_data):
    pdf = FPDF()
    pdf.add_page()
    
    # Configuração da fonte
    pdf.set_font("Arial", size=12)
    
    # Formato DD/MM/YYYY
    dob_formatted = dob if isinstance(dob, str) else dob.strftime('%d/%m/%Y')
    report_date = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    # Cabeçalho
    pdf.cell(200, 10, txt=clean_text("Medication Crushability Review Report"), ln=1, align='C')
    pdf.cell(200, 10, txt=clean_text(f"Patient: {patient_name}"), ln=1)
    pdf.cell(200, 10, txt=clean_text(f"Date of Birth: {dob_formatted}"), ln=1)
    pdf.cell(200, 10, txt=clean_text(f"Report Date: {report_date}"), ln=1)
    
    # Nota importante
    pdf.set_font("Arial", 'B', 10)
    pdf.multi_cell(0, 10, txt=clean_text("IMPORTANT: Crushing tablets renders the medication unlicensed. If a liquid formulation is available, this is always the preferred option."))
    pdf.ln(5)
    
    # Tabela de medicamentos
    pdf.set_font("Arial", size=10)
    col_widths = [50, 30, 50, 70]  # Larguras das colunas
    
    # Cabeçalhos
    headers = ["Medication", "Crushable?", "Alternative", "Recommendation"]
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, txt=clean_text(header), border=1)
    pdf.ln()
    
    # Conteúdo
    for med in selected_meds:
        med_info = med_data[med_data['Drug'] == med].iloc[0]
        data = [
            str(med_info['Drug']),
            str(med_info['Can be Crushed']),
            str(med_info['Alternative form available?']),
            str(med_info['Recommendation'])
        ]
        for i, item in enumerate(data):
            pdf.cell(col_widths[i], 10, txt=clean_text(item), border=1)
        pdf.ln()
    
    # Rodapé
    pdf.set_y(-20)
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, txt=clean_text("Crush Med Review Tool developed by Celio Lara Junior - Pharmacist - PSI 10003245"), ln=1, align='R')
    
    # CORREÇÃO: Removido o .encode() desnecessário
    return pdf.output(dest='S')  # Já retorna bytes

def main():
    st.title("💊 Medication Crushability Review")
    st.markdown("""
    This application generates medication review reports for patients with swallowing difficulties.
    """)
    
    med_data = load_data()
    med_list = med_data['Drug'].dropna().unique().tolist()
    
    with st.expander("Patient Information", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            patient_name = st.text_input("Full Name*")
        with col2:
            min_date = date(1900, 1, 1)
            max_date = date.today()
            dob = st.date_input(
                "Date of Birth* (DD/MM/YYYY)",
                min_value=min_date,
                max_value=max_date,
                value=date(1980, 1, 1),
                format="DD/MM/YYYY"
            )
    
    with st.expander("Medication Selection", expanded=True):
        search_term = st.text_input("Search medication:")
        
        filtered_meds = [m for m in med_list if search_term.lower() in str(m).lower()] if search_term else med_list
        
        selected_meds = st.multiselect(
            "Select medications for review (no limit):",
            filtered_meds,
            help="Select all relevant medications"
        )
        
        if selected_meds:
            st.dataframe(
                med_data[med_data['Drug'].isin(selected_meds)].reset_index(drop=True),
                use_container_width=True,
                height=min(300, len(selected_meds) * 35 + 35)
            )
    
    if st.button("Generate Full Report"):
        if not patient_name or not dob or not selected_meds:
            st.error("Please complete all required fields (*)")
        else:
            with st.spinner("Generating reports..."):
                doc = generate_word_report(patient_name, dob, selected_meds, med_data)
                word_buffer = BytesIO()
                doc.save(word_buffer)
                
                pdf_bytes = generate_pdf_report(patient_name, dob, selected_meds, med_data)
                
                st.success("Reports generated successfully!")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label=f"⬇️ Download Word ({len(selected_meds)} meds)",
                        data=word_buffer.getvalue(),
                        file_name=f"Crushability_Review_{patient_name.replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                with col2:
                    st.download_button(
                        label=f"⬇️ Download PDF ({len(selected_meds)} meds)",
                        data=pdf_bytes,
                        file_name=f"Crushability_Review_{patient_name.replace(' ', '_')}.pdf",
                        mime="application/pdf"
                    )

if __name__ == "__main__":
    main()
