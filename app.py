import streamlit as st
import pandas as pd
from docx import Document
import io
import os
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Crush Med Review",
    page_icon="💊",
    layout="wide"
)

# Carregar dados
@st.cache_data
def load_data():
    if os.path.exists("data/Crush_Med_Data_Bank_Clean.xlsx"):
        return pd.read_excel("data/Crush_Med_Data_Bank_Clean.xlsx")
    else:
        uploaded_file = st.file_uploader("Upload do banco de dados (Crush_Med_Data_Bank_Clean.xlsx)", type="xlsx")
        if uploaded_file:
            return pd.read_excel(uploaded_file)
        else:
            st.warning("Por favor, faça upload do arquivo de dados")
            st.stop()

# Gerar relatório
def generate_report(resident_name, dob, selected_meds, med_data):
    doc = Document()
    doc.add_heading('Relatório de Revisão de Medicação', level=1)
    doc.add_paragraph(f"**Paciente:** {resident_name}\t**Data de Nascimento:** {dob}")
    doc.add_paragraph(f"**Data do Relatório:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    p_note = doc.add_paragraph()
    p_note.add_run("NOTA: Triturar comprimidos torna o medicamento não licenciado. Se uma forma líquida estiver disponível, esta é sempre a opção preferida.").bold = True
    
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Medicamento'
    hdr_cells[1].text = 'Pode Triturar?'
    hdr_cells[2].text = 'Forma Alternativa'
    hdr_cells[3].text = 'Recomendação'
    
    for med in selected_meds:
        med_info = med_data[med_data['Drug'] == med].iloc[0]
        row_cells = table.add_row().cells
        row_cells[0].text = str(med_info['Drug'])
        row_cells[1].text = str(med_info['Can be Crushed'])
        row_cells[2].text = str(med_info['Alternative form available?'])
        row_cells[3].text = str(med_info['Recommendation'])
    
    doc.add_paragraph("\n\n")
    doc.add_paragraph("Farmacêutico Responsável: ___________________________ Data: ___/___/_______")
    doc.add_paragraph("Enfermeiro Responsável: _____________________________ Data: ___/___/_______")
    doc.add_paragraph("Médico Responsável: _________________________________ Data: ___/___/_______")
    
    return doc

# Interface principal
def main():
    st.title("💊 Crush Medication Review")
    st.markdown("Este aplicativo gera relatórios de revisão de medicamentos para pacientes com dificuldade de deglutição.")
    
    med_data = load_data()
    med_list = med_data['Drug'].dropna().unique().tolist()
    
    with st.expander("Informações do Paciente", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            resident_name = st.text_input("Nome Completo do Paciente*")
        with col2:
            dob = st.date_input("Data de Nascimento*")
    
    with st.expander("Seleção de Medicamentos", expanded=True):
        search_term = st.text_input("Pesquisar medicamento:")
        filtered_meds = [m for m in med_list if search_term.lower() in str(m).lower()] if search_term else med_list
        selected_meds = st.multiselect("Selecione os medicamentos para revisão:", filtered_meds)
        
        if selected_meds:
            st.dataframe(med_data[med_data['Drug'].isin(selected_meds)].reset_index(drop=True), use_container_width=True)
    
    if st.button("Gerar Relatório Completo"):
        if not resident_name or not dob or not selected_meds:
            st.error("Por favor, preencha todas as informações obrigatórias (*)")
        else:
            doc = generate_report(resident_name, dob.strftime('%d/%m/%Y'), selected_meds, med_data)
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button(
                label="⬇️ Download do Relatório",
                data=bio.getvalue(),
                file_name=f"Crush_Med_Review_{resident_name.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

if __name__ == "__main__":
    main()
