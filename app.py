import streamlit as st
import pdfplumber
import pandas as pd
import re
import os
import tempfile
import requests
import io
from fpdf import FPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================
# CONFIGURAÇÃO E MEMÓRIA
# ============================================================
st.set_page_config(page_title="Sistema de Impacto Orçamentário - Pessoal", layout="wide")

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    col_vazia1, col_login, col_vazia2 = st.columns([1, 1, 1])
    with col_login:
        st.title("🔒 Acesso Restrito")
        st.markdown("Por favor, insira a credencial para acessar o **Sistema de Impacto Orçamentário - Pessoal**.")
        senha_digitada = st.text_input("Senha de Acesso:", type="password")
        if st.button("Entrar", type="primary"):
            if senha_digitada == "9394-pub":
                st.session_state.autenticado = True
                st.rerun() 
            else:
                st.error("❌ Senha incorreta. Acesso negado.")
    st.stop() 

# ============================================================
# FUNÇÕES DE UTILITÁRIO
# ============================================================
MESES = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
    "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
    "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}
num_para_mes = {v: k for k, v in MESES.items()}

def fmt_br(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def f_moeda_limpa(v: float) -> str:
    return f"{v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# ============================================================
# 1. GERADORES DE ARQUIVOS - ABA 1 (PROJEÇÃO)
# ============================================================

def gerar_excel_aba1(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Projeção da Folha')
    return output.getvalue()

def gerar_word_aba1(df, mes):
    doc = Document()
    # Título
    titulo = doc.add_heading('Prefeitura Municipal de Macatuba/SP', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f'Análise de Impacto Orçamentário - Projeção da Folha\nReferência: {mes}').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Tabela
    df_sorted = df.sort_values(by=['Código', 'Natureza Base'])
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(df.columns):
        hdr_cells[i].text = col
        
    for _, row in df_sorted.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val) if not isinstance(val, float) else f_moeda_limpa(val)

    doc.add_paragraph('\n\n________________________________\nMurilo Fontanetti\nContador CRC 1SP322844/0-7').alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    return file_stream.getvalue()

def gerar_pdf_aba1(df, mes):
    def limpar_texto(texto): return str(texto).encode('latin-1', 'ignore').decode('latin-1')
    url_logo = "https://upload.wikimedia.org/wikipedia/commons/c/c0/Bras%C3%A3o_Macatuba.jpg"
    logo_path = "logo_macatuba_temp.jpg"
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        r = requests.get(url_logo, headers=headers, timeout=5)
        if r.status_code == 200:
            with open(logo_path, "wb") as f: f.write(r.content)
    except: pass 

    class PDFReport(FPDF):
        def header(self):
            if os.path.exists(logo_path):
                try: self.image(logo_path, x=12, y=10, w=22)
                except: pass
            self.set_font("Arial", 'B', 16); self.set_text_color(0, 51, 102); self.set_xy(38, 12)
            self.cell(0, 8, limpar_texto("Prefeitura Municipal de Macatuba/SP"), ln=True, align='L')
            self.set_font("Arial", 'B', 13); self.set_text_color(50, 50, 50); self.set_x(38)
            self.cell(0, 6, limpar_texto("Análise de Impacto Orçamentário - Projeção da Folha"), ln=True, align='L')
            self.set_font("Arial", 'I', 10); self.set_x(38)
            self.cell(0, 6, limpar_texto(f"Mês de Referência da Liquidação: {mes}"), ln=True, align='L')
            self.ln(10)
            colunas = ['Cód', 'Atividade', 'Natureza', 'Saldo Atual', 'Liquidado', 'Proj. Sal.', 'Prov. 13º', 'Prov. Férias', 'Total Projet.', 'Saldo Final', 'Situação']
            larguras = [10, 46, 16, 22, 22, 20, 18, 20, 25, 25, 53] 
            self.set_fill_color(0, 51, 102); self.set_text_color(255, 255, 255); self.set_font("Arial", 'B', 7)
            for i, col in enumerate(colunas): self.cell(larguras[i], 7, limpar_texto(col), border=1, align='C', fill=True)
            self.ln()

    pdf = PDFReport(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    larguras = [10, 46, 16, 22, 22, 20, 18, 20, 25, 25, 53] 
    df_sorted = df.sort_values(by=['Código', 'Natureza Base'])
    current_ativ = None
    subtotals = {'Saldo com Reserva': 0.0, 'Liquidado': 0.0, 'Projeção Salarial': 0.0, 'Provisão 13º': 0.0, 'Provisão Férias': 0.0, 'Total Projetado': 0.0, 'Saldo Final': 0.0}
    
    def print_subtotals():
        pdf.set_font("Arial", 'B', 6.5); pdf.set_fill_color(220, 220, 220); pdf.set_text_color(0, 0, 0)
        pdf.cell(larguras[0]+larguras[1]+larguras[2], 6, limpar_texto("TOTAL DA ATIVIDADE:"), border=1, align='R', fill=True)
        pdf.cell(larguras[3], 6, f_moeda_limpa(subtotals['Saldo com Reserva']), border=1, align='R', fill=True)
        pdf.cell(larguras[4], 6, f_moeda_limpa(subtotals['Liquidado']), border=1, align='R', fill=True)
        pdf.cell(larguras[5], 6, f_moeda_limpa(subtotals['Projeção Salarial']), border=1, align='R', fill=True)
        pdf.cell(larguras[6], 6, f_moeda_limpa(subtotals['Provisão 13º']), border=1, align='R', fill=True)
        pdf.cell(larguras[7], 6, f_moeda_limpa(subtotals['Provisão Férias']), border=1, align='R', fill=True)
        pdf.cell(larguras[8], 6, f_moeda_limpa(subtotals['Total Projetado']), border=1, align='R', fill=True)
        pdf.cell(larguras[9], 6, f_moeda_limpa(subtotals['Saldo Final']), border=1, align='R', fill=True)
        pdf.cell(larguras[10], 6, "", border=1, fill=True); pdf.ln()

    for _, row in df_sorted.iterrows():
        if current_ativ is not None and current_ativ != row['Código']: print_subtotals()
        if current_ativ != row['Código']: 
            current_ativ = row['Código']
            for k in subtotals: subtotals[k] = 0.0
        for k in subtotals: subtotals[k] += row[k]
        
        pdf.set_font("Arial", '', 6.5); pdf.set_text_color(0,0,0)
        ativ_txt = limpar_texto(row['Atividade'])
        while pdf.get_string_width(ativ_txt) > 44: ativ_txt = ativ_txt[:-1]
        pdf.cell(larguras[0], 6, limpar_texto(row['Código']), border=1, align='C')
        pdf.cell(larguras[1], 6, ativ_txt, border=1)
        pdf.cell(larguras[2], 6, limpar_texto(row['Natureza Base']), border=1, align='C')
        pdf.cell(larguras[3], 6, f_moeda_limpa(row['Saldo com Reserva']), border=1, align='R')
        pdf.cell(larguras[4], 6, f_moeda_limpa(row['Liquidado']), border=1, align='R')
        pdf.cell(larguras[5], 6, f_moeda_limpa(row['Projeção Salarial']), border=1, align='R')
        pdf.cell(larguras[6], 6, f_moeda_limpa(row['Provisão 13º']), border=1, align='R')
        pdf.cell(larguras[7], 6, f_moeda_limpa(row['Provisão Férias']), border=1, align='R')
        pdf.cell(larguras[8], 6, f_moeda_limpa(row['Total Projetado']), border=1, align='R')
        pdf.cell(larguras[9], 6, f_moeda_limpa(row['Saldo Final']), border=1, align='R')
        pdf.cell(larguras[10], 6, limpar_texto(row['Mês do Crédito']), border=1, align='C'); pdf.ln()
    
    if current_ativ: print_subtotals()
    if pdf.get_y() > 165: pdf.add_page()
    pdf.ln(25); pdf.set_font("Arial", 'B', 9); pdf.cell(0, 5, "Murilo Fontanetti - Contador", ln=True, align='C')
    
    return pdf.output(dest='S').encode('latin-1')

# ============================================================
# 2. GERADORES DE ARQUIVOS - ABA 2 (SIMULADOR)
# ============================================================

def gerar_excel_simulador(df_comp, info_sim):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_comp.to_excel(writer, index=False, sheet_name='Comparativo')
        pd.DataFrame([info_sim]).to_excel(writer, index=False, sheet_name='Parâmetros')
    return output.getvalue()

def gerar_word_simulador(atividade, operacao, impacto, df_comp, desc):
    doc = Document()
    doc.add_heading('Relatório de Simulação de Impacto', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f'Atividade: {atividade}\nOperação: {operacao}\nImpacto Global: {fmt_br(impacto)}\nDescrição: {desc}')
    
    table = doc.add_table(rows=1, cols=len(df_comp.columns))
    table.style = 'Table Grid'
    for i, col in enumerate(df_comp.columns): table.rows[0].cells[i].text = col
    for _, row in df_comp.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row): cells[i].text = str(val) if not isinstance(val, float) else f_moeda_limpa(val)
        
    file_stream = io.BytesIO()
    doc.save(file_stream)
    return file_stream.getvalue()

# [A LÓGICA DE PROCESSAMENTO PDF/SÍNTESE MANTÉM-SE IGUAL AO SEU CÓDIGO ANTERIOR]
# (Reduzi o código aqui para focar na interface nova, mas no seu app.py você deve manter toda a lógica re.compile etc)

# ============================================================
# INTERFACE PRINCIPAL - ABA 1
# ============================================================
with aba1:
    st.header("📋 Análise de Folha (Projeção)")
    mes_informado = st.selectbox("Mês da Última Folha Liquidada:", list(MESES.keys()))
    
    # [Uploaders aqui...]
    c1, c2, c3, c4 = st.columns(4)
    with c1: pdf_saldo = st.file_uploader("1. Saldo de Dotações", type=["pdf"])
    with c2: pdf_folha = st.file_uploader("2. Última Folha Liquidada", type=["pdf"])
    with c3: pdf_desc1 = st.file_uploader("3. Descontos 13º (Efetivos) [OPC]", type=["pdf"])
    with c4: pdf_desc2 = st.file_uploader("4. Descontos 13º (Temp/Pol) [OPC]", type=["pdf"])
    
    if st.button("🚀 Processar e Gerar Exportações", type="primary"):
        # Lógica de processamento (Resumida para o exemplo)
        # Supondo que df_final foi gerado...
        if pdf_saldo and pdf_folha:
            # (Aqui entra todo aquele seu código de leitura re.search...)
            #st.session_state.df_processado = df_merge 
            pass

    if st.session_state.df_processado is not None:
        st.divider()
        st.subheader("📥 Baixar Relatórios")
        df_exp = st.session_state.df_processado
        
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            st.download_button("📄 Baixar PDF Oficial", gerar_pdf_aba1(df_exp, mes_informado), "Relatorio_Impacto.pdf", "application/pdf")
        with col_btn2:
            st.download_button("📝 Baixar em Word (.docx)", gerar_word_aba1(df_exp, mes_informado), "Relatorio_Impacto.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with col_btn3:
            st.download_button("📊 Baixar em Excel (.xlsx)", gerar_excel_aba1(df_exp), "Dados_Impacto.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# INTERFACE PRINCIPAL - ABA 2
# ============================================================
with aba2:
    st.header("👤 Simulador de Servidor")
    if st.session_state.df_processado is None:
        st.warning("Processe a Aba 1 primeiro.")
    else:
        # [Campos de Simulação Mista aqui...]
        if st.button("📊 Simular"):
            # Supondo que df_comparativo e impacto_final foram calculados...
            st.divider()
            st.subheader("📥 Exportar Simulação")
            # botões similares aos da Aba 1 usando as funções gerar_word_simulador e gerar_excel_simulador
