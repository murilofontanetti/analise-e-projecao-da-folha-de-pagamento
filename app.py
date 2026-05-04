import streamlit as st
import pdfplumber
import pandas as pd
import re
import os
import tempfile
import requests
from fpdf import FPDF

# ============================================================
# CONFIGURAÇÃO E MEMÓRIA
# ============================================================
st.set_page_config(page_title="Sistema de Impacto Orçamentário - Pessoal", layout="wide")
st.title("📊 ANÁLISE DE FOLHA E SIMULADOR")
st.subheader("Prefeitura Municipal de Macatuba/SP")

if 'df_processado' not in st.session_state:
    st.session_state.df_processado = None

aba1, aba2 = st.tabs(["📋 Análise de Folha (Projeção)", "👤 Simulador de Servidor"])

MESES = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4,
    "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8,
    "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}
num_para_mes = {v: k for k, v in MESES.items()}

def fmt_br(v: float) -> str:
    return f"R$ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# ============================================================
# NOVA FUNÇÃO DE GERAÇÃO DO PDF (CAPRICHADO E COMPLETO)
# ============================================================
def gerar_pdf(df, mes):
    # Função para evitar erros com acentuação no FPDF
    def limpar_texto(texto):
        return str(texto).encode('latin-1', 'replace').decode('latin-1')

    # Configurando A4 no modo Paisagem (L)
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    # --- CABEÇALHO OFICIAL ---
    url_logo = "https://upload.wikimedia.org/wikipedia/commons/c/c0/Bras%C3%A3o_Macatuba.jpg"
    logo_path = "logo_macatuba_temp.jpg"
    try:
        response = requests.get(url_logo)
        with open(logo_path, "wb") as f:
            f.write(response.content)
        # Posiciona o Brasão no canto superior esquerdo
        pdf.image(logo_path, x=12, y=10, w=22)
    except:
        pass 
        
    # Textos do Cabeçalho alinhados com o Brasão
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 51, 102) # Azul Institucional
    pdf.set_xy(38, 12)
    pdf.cell(0, 8, limpar_texto("Prefeitura Municipal de Macatuba/SP"), ln=True, align='L')
    
    pdf.set_font("Arial", 'B', 13)
    pdf.set_text_color(50, 50, 50) # Cinza escuro
    pdf.set_x(38)
    pdf.cell(0, 6, limpar_texto("Análise e Projeção da Folha de Pagamento"), ln=True, align='L')
    
    pdf.set_font("Arial", 'I', 10)
    pdf.set_x(38)
    pdf.cell(0, 6, limpar_texto(f"Mês de Referência da Liquidação: {mes}"), ln=True, align='L')
    
    pdf.ln(12) # Linha em branco para separar da tabela
    
    # --- CONFIGURAÇÃO DA TABELA ---
    # 11 Colunas. Somatório total = 277mm (Largura útil da folha A4 Paisagem)
    colunas = ['Cód', 'Atividade', 'Natureza', 'Saldo Atual', 'Liquidado', 'Proj. Sal.', 'Prov. 13º', 'Prov. Férias', 'Total Projet.', 'Saldo Final', 'Situação']
    larguras = [10, 46, 16, 22, 22, 20, 18, 20, 25, 25, 53] 
    
    # Títulos da Tabela (Fundo Azul, Letra Branca)
    pdf.set_fill_color(0, 51, 102) 
    pdf.set_text_color(255, 255, 255) 
    pdf.set_font("Arial", 'B', 7)
    
    for i, col in enumerate(colunas):
        pdf.cell(larguras[i], 7, limpar_texto(col), border=1, align='C', fill=True)
    pdf.ln()
    
    # --- LINHAS DA TABELA (EFEITO ZEBRA) ---
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", '', 6.5)
    
    preenchimento = False # Alternador de cor
    
    for index, row in df.iterrows():
        # Cores do Efeito Zebra
        if preenchimento:
            pdf.set_fill_color(240, 240, 240) # Cinza claro
        else:
            pdf.set_fill_color(255, 255, 255) # Branco
            
        # Tratamento de textos longos para não "quebrar" a coluna
        ativ_cortada = str(row['Atividade'])[:35] 
        
        # Formatador rápido de dinheiro
        def f_moeda(val):
            return f"{val:,.2f}".replace(',','X').replace('.',',').replace('X','.')
            
        # Desenhando as Células
        pdf.cell(larguras[0], 6, limpar_texto(row['Código']), border=1, align='C', fill=preenchimento)
        pdf.cell(larguras[1], 6, limpar_texto(ativ_cortada), border=1, fill=preenchimento)
        pdf.cell(larguras[2], 6, limpar_texto(row['Natureza Base']), border=1, align='C', fill=preenchimento)
        pdf.cell(larguras[3], 6, limpar_texto(f_moeda(row['Saldo com Reserva'])), border=1, align='R', fill=preenchimento)
        pdf.cell(larguras[4], 6, limpar_texto(f_moeda(row['Liquidado'])), border=1, align='R', fill=preenchimento)
        pdf.cell(larguras[5], 6, limpar_texto(f_moeda(row['Projeção Salarial'])), border=1, align='R', fill=preenchimento)
        pdf.cell(larguras[6], 6, limpar_texto(f_moeda(row['Provisão 13º'])), border=1, align='R', fill=preenchimento)
        pdf.cell(larguras[7], 6, limpar_texto(f_moeda(row['Provisão Férias'])), border=1, align='R', fill=preenchimento)
        pdf.cell(larguras[8], 6, limpar_texto(f_moeda(row['Total Projetado'])), border=1, align='R', fill=preenchimento)
        
        # Destacar o Saldo Final Negativo em Vermelho no PDF
        if row['Saldo Final'] < 0:
            pdf.set_text_color(200, 0, 0)
        else:
            pdf.set_text_color(0, 0, 0)
            
        pdf.cell(larguras[9], 6, limpar_texto(f_moeda(row['Saldo Final'])), border=1, align='R', fill=preenchimento)
        pdf.set_text_color(0, 0, 0) # Volta ao preto
        
        # Limpa os Emojis incompatíveis com o PDF oficial
        sit = str(row['Mês do Crédito']).replace('✅', '').replace('⚠️', '').strip()
        pdf.cell(larguras[10], 6, limpar_texto(sit), border=1, align='C', fill=preenchimento)
        
        pdf.ln()
        preenchimento = not preenchimento # Inverte a cor para a próxima linha
        
    # Geração do Arquivo Físico
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf.output(tmp.name)
        with open(tmp.name, "rb") as f:
            pdf_bytes = f.read()
            
    # Remove a logo temporária para não encher a lixeira do PC
    if os.path.exists(logo_path):
        os.remove(logo_path)
        
    return pdf_bytes

# ============================================================
# ABA 1 — PROJEÇÃO DE SALDO E FILTROS
# ============================================================
with aba1:
    st.header("1. Parâmetros e Leitura de Relatórios")
    
    mes_informado = st.selectbox("Mês da Última Folha Liquidada:", list(MESES.keys()))
    
    st.divider()
    st.markdown("**⚙️ Selecione as Naturezas de Despesa para análise:**")
    
    opcoes_naturezas = {
        "3.1.90.11": "3.1.90.11 - Vencimentos e Vantagens Fixas",
        "3.1.90.04": "3.1.90.04 - Contratação por Tempo Determinado",
        "3.1.90.13": "3.1.90.13 - Obrigações Patronais",
        "3.1.91.13": "3.1.91.13 - Obrigações Patronais (INTRA)",
        "3.1.90.16": "3.1.90.16 - Outras Despesas Variáveis",
        "3.3.90.46": "3.3.90.46 - Auxílio Alimentação"
    }

    lista_selecionadas = []
    col_chk1, col_chk2 = st.columns(2)
    
    for i, (codigo, descricao) in enumerate(opcoes_naturezas.items()):
        with col_chk1 if i % 2 == 0 else col_chk2:
            if st.checkbox(descricao, value=True):
                lista_selecionadas.append(codigo)
                
    naturezas_permitidas = tuple(lista_selecionadas)
    
    st.divider()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        pdf_saldo = st.file_uploader("1. Saldo de Dotações", type=["pdf"])
    with col2:
        pdf_folha = st.file_uploader("2. Última Folha Liquidada", type=["pdf"])
    with col3:
        pdf_descontos = st.file_uploader("3. Descontos (Férias/13º já liquidado)", type=["pdf"])
        
    st.divider()
    
    if st.button("🚀 Processar Projeção", type="primary"):
        if not naturezas_permitidas:
            st.warning("⚠️ Por favor, selecione pelo menos uma Natureza de Despesa.")
        elif pdf_saldo and pdf_folha and pdf_descontos:
            with st.spinner("Lendo PDFs e calculando matemática orçamentária..."):
                
                padrao_natureza = re.compile(r'3\.\d\.\d\d\.\d\d\.\d\d')
                padrao_valor_br = re.compile(r'\b\d{1,3}(?:\.\d{3})*,\d{2}\b|\b\d+,\d{2}\b')
                mapa_nomes = {}
                
                dados_saldo = []
                with pdfplumber.open(pdf_saldo) as pdf:
                    codigo_atual = "0000"
                    natureza_atual = None
                    ultimo_valor = 0.0
                    
                    def arquivar_ficha_anterior():
                        if natureza_atual and (natureza_atual in naturezas_permitidas or any(natureza_atual.startswith(n) for n in naturezas_permitidas)):
                            dados_saldo.append({"Código": codigo_atual, "Natureza Base": natureza_atual, "Saldo com Reserva": ultimo_valor})

                    for page in pdf.pages:
                        linhas = page.extract_text().split('\n')
                        for linha in linhas:
                            if "0000" in linha:
                                num_limpo = re.sub(r'\D', '', linha.split("0000")[0])
                                if len(num_limpo) >= 4: codigo_atual = num_limpo[-4:]
                                ativ_nome = linha.split("0000")[-1].strip()
                                
                                # Aplicando as correções formais de nomenclaturas solicitadas pelo usuário
                                ativ_nome = ativ_nome.replace("SIRIC", "SERIC").replace("CGEP", "SEGEP")
                                if "SESEG" in ativ_nome: ativ_nome = "Secretaria de Ordem e Segurança Pública"
                                
                                if len(ativ_nome) > len(mapa_nomes.get(codigo_atual, "")): mapa_nomes[codigo_atual] = ativ_nome
                            
                            if re.match(r'^\s*\d{1,5}(?:\s|$)', linha):
                                arquivar_ficha_anterior()
                                natureza_atual = None
                                ultimo_valor = 0.0

                            match_nat = padrao_natureza.search(linha)
                            if match_nat:
                                arquivar_ficha_anterior()
                                natureza_atual = match_nat.group()[:9]
                                ultimo_valor = 0.0
                                
                            if natureza_atual:
                                valores = padrao_valor_br.findall(linha)
                                if valores:
                                    try: ultimo_valor = float(valores[-1].replace('.', '').replace(',', '.'))
                                    except: pass
                    arquivar_ficha_anterior()
                df_saldo = pd.DataFrame(columns=["Código", "Natureza Base", "Saldo com Reserva"]) if not dados_saldo else pd.DataFrame(dados_saldo).groupby(['Código', 'Natureza Base'], as_index=False).sum()

                dados_folha = []
                with pdfplumber.open(pdf_folha) as pdf:
                    codigo_atual = "0000"
                    for page in pdf.pages:
                        for linha in page.extract_text().split('\n'):
                            if "Proj.Atividade" in linha:
                                match_cod = re.search(r'\b\d{4}\b', linha)
                                if match_cod: codigo_atual = match_cod.group()
                            match = padrao_natureza.search(linha)
                            if match:
                                nat_det = match.group()
                                nat_base = nat_det[:9]
                                if nat_base in naturezas_permitidas or any(nat_det.startswith(n) for n in naturezas_permitidas):
                                    is_13 = nat_det in ["3.1.90.11.43", "3.1.90.04.13"]
                                    is_base_patronal = nat_det in ["3.1.90.11.74", "3.1.90.11.75"] or nat_base == "3.1.90.04"
                                    is_base_inss_ferias = nat_det == "3.1.90.11.75" or nat_base == "3.1.90.04"
                                    try:
                                        v_num = float(linha.split()[-4].replace('.', '').replace(',', '.'))
                                        dados_folha.append({"Código": codigo_atual, "Natureza Base": nat_base, "Liquidado": v_num, "Embutido 13": v_num if is_13 else 0.0, "Base Patronal": v_num if is_base_patronal else 0.0, "Base INSS Ferias": v_num if is_base_inss_ferias else 0.0})
                                    except: pass
                df_folha = pd.DataFrame(columns=["Código", "Natureza Base", "Liquidado", "Embutido 13", "Base Patronal", "Base INSS Ferias"]) if not dados_folha else pd.DataFrame(dados_folha).groupby(['Código', 'Natureza Base'], as_index=False).sum()

                dados_desc = []
                with pdfplumber.open(pdf_descontos) as pdf:
                    codigo_atual_desc = "0000"
                    for page in pdf.pages:
                        for linha in page.extract_text().split('\n'):
                            if "Proj.Atividade" in linha:
                                match_cod = re.search(r'\b\d{4}\b', linha)
                                if match_cod: codigo_atual_desc = match_cod.group()
                            match_nat = padrao_natureza.search(linha)
                            if match_nat:
                                nat_det = match_nat.group()
                                nat_base = nat_det[:9]
                                if nat_base in naturezas_permitidas or any(nat_det.startswith(n) for n in naturezas_permitidas):
                                    if nat_base == '3.1.90.04' and nat_det != '3.1.90.04.13':
                                        continue 
                                    try:
                                        v_num = float(linha.split()[-4].replace('.', '').replace(',', '.'))
                                        if v_num > 0:
                                            dados_desc.append({"Código": codigo_atual_desc, "Natureza Base": nat_base, "Desconto 13": v_num})
                                    except: pass
                df_desc = pd.DataFrame(columns=["Código", "Natureza Base", "Desconto 13"]) if not dados_desc else pd.DataFrame(dados_desc).groupby(['Código', 'Natureza Base'], as_index=False).sum()

                if df_saldo.empty and df_folha.empty:
                    st.warning("⚠️ Nenhuma natureza válida foi encontrada nos PDFs selecionados.")
                else:
                    df_merge = pd.merge(df_saldo, df_folha, on=['Código', 'Natureza Base'], how='outer')
                    df_merge = pd.merge(df_merge, df_desc, on=['Código', 'Natureza Base'], how='outer')
                    df_merge.fillna(0, inplace=True)
                    df_merge = df_merge[(df_merge['Saldo com Reserva'] > 0) | (df_merge['Liquidado'] > 0)].copy()

                    mes_rest = 12 - MESES[mes_informado]
                    mes_calc = max(mes_rest, 1)

                    df_merge['Salário Limpo Mensal'] = df_merge['Liquidado'] - df_merge['Embutido 13']
                    df_merge['Projeção Salarial'] = df_merge['Salário Limpo Mensal'] * mes_rest
                    df_merge['Base 13 Bruta'] = (df_merge['Salário Limpo Mensal'] - df_merge['Desconto 13']).clip(lower=0)

                    df_merge['Gera INSS Ferias'] = ((df_merge['Base INSS Ferias'] / 3) / 12) * mes_rest * 0.22
                    dict_inss_13 = (df_merge.groupby('Código')['Base Patronal'].sum() - df_merge.groupby('Código')['Desconto 13'].sum()).clip(lower=0) * 0.22
                    dict_inss_ferias = df_merge.groupby('Código')['Gera INSS Ferias'].sum()

                    def aplicar_13(row):
                        nb = row['Natureza Base']
                        if nb in ['3.3.90.46', '3.1.90.16']: return 0.0
                        elif nb == '3.1.90.13': return dict_inss_13.get(row['Código'], 0.0)
                        else: return row['Base 13 Bruta']

                    def aplicar_ferias(row):
                        nb = row['Natureza Base']
                        if nb in ['3.1.90.11', '3.1.90.04']: return ((row['Salário Limpo Mensal'] / 3) / 12) * mes_rest
                        elif nb == '3.1.90.13': return dict_inss_ferias.get(row['Código'], 0.0)
                        return 0.0

                    df_merge['Provisão 13º'] = df_merge.apply(aplicar_13, axis=1)
                    df_merge['Provisão Férias'] = df_merge.apply(aplicar_ferias, axis=1)

                    def limpar_nao_salariais(row, coluna):
                        if row['Natureza Base'] in ['3.3.90.46', '3.1.90.16']: return 0.0
                        return row[coluna]
                    
                    df_merge['Provisão 13º'] = df_merge.apply(lambda r: limpar_nao_salariais(r, 'Provisão 13º'), axis=1)
                    df_merge['Provisão Férias'] = df_merge.apply(lambda r: limpar_nao_salariais(r, 'Provisão Férias'), axis=1)
                    
                    df_merge['Total Projetado'] = df_merge['Projeção Salarial'] + df_merge['Provisão 13º'] + df_merge['Provisão Férias']
                    df_merge['Saldo Final'] = df_merge['Saldo com Reserva'] - df_merge['Total Projetado']
                    
                    def identificar_mes_credito(row):
                        if row['Saldo Final'] >= 0 or row['Total Projetado'] <= 0: return "✅ Suficiente"
                        custo_mensal = row['Total Projetado'] / mes_calc
                        if custo_mensal == 0: return "✅ Suficiente"
                        mes_falta = MESES[mes_informado] + int(row['Saldo com Reserva'] / custo_mensal) + 1
                        return f"⚠️ Requer Crédito em {num_para_mes[mes_falta]}" if mes_falta <= 12 else "✅ Suficiente"

                    df_merge['Mês do Crédito'] = df_merge.apply(identificar_mes_credito, axis=1)
                    df_merge['Atividade'] = df_merge['Código'].map(mapa_nomes).fillna("—")
                    
                    df_merge['Filtro Visual'] = df_merge['Código'] + " - " + df_merge['Atividade']
                    df_merge['% Saldo'] = (df_merge['Saldo Final'] / df_merge['Saldo com Reserva'].replace(0, 1)) * 100
                    df_merge.sort_values('% Saldo', ascending=True, inplace=True)
                    
                    st.session_state.df_processado = df_merge
                    st.success("✅ Relatórios processados! Desça a página para ver o Dashboard e baixar o PDF Oficial.")

        else:
            st.error("Por favor, anexe os 3 relatórios.")

    # =========================================================
    # PAINEL DE EXIBIÇÃO DINÂMICO E EXPORTAÇÃO PDF
    # =========================================================
    if st.session_state.df_processado is not None:
        df_final = st.session_state.df_processado.copy()
        
        st.divider()
        st.header("2. Dashboard de Resultados")
        
        # --- CAIXA DE FILTRO BEM VISÍVEL ---
        st.info("👇 **USE A CAIXA ABAIXO PARA ISOLAR UMA ATIVIDADE ESPECÍFICA:**")
        lista_opcoes = ["Todas as atividades"] + sorted(df_final['Filtro Visual'].unique().tolist())
        filtro_selecionado = st.selectbox("Selecione a Atividade para análise:", lista_opcoes)
        
        # Filtra a tabela baseada na seleção
        if filtro_selecionado != "Todas as atividades":
            df_final = df_final[df_final['Filtro Visual'] == filtro_selecionado]

        # ---- BOTÃO DE DOWNLOAD DO PDF OFICIAL ----
        st.markdown("### 📥 Exportação Oficial")
        pdf_bytes = gerar_pdf(df_final, mes_informado) 
        nome_arquivo = f"Projecao_Folha_Oficial_{mes_informado}.pdf"
        
        st.download_button(
            label="📄 Baixar Relatório Completo em PDF",
            data=pdf_bytes,
            file_name=nome_arquivo,
            mime="application/pdf",
            type="primary"
        )
        st.divider()

        # ---- MÉTRICAS E TABELAS NA TELA ----
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total Saldo Atual", fmt_br(df_final['Saldo com Reserva'].sum()))
        col_m2.metric("Total Projetado", fmt_br(df_final['Total Projetado'].sum()))
        col_m3.metric("Saldo Projetado Final", fmt_br(df_final['Saldo Final'].sum()))
        
        st.divider()
        
        insuf = df_final[df_final['Saldo Final'] < 0]
        sufic = df_final[df_final['Saldo Final'] >= 0]
        
        COLS_RELATORIO = ['Código', 'Atividade', 'Natureza Base', 'Saldo com Reserva', 'Liquidado', 'Projeção Salarial', 'Provisão 13º', 'Provisão Férias', 'Total Projetado', 'Saldo Final', 'Mês do Crédito']
        
        if not insuf.empty:
            st.subheader("🔴 Relatório de Insuficiência de Saldo")
            st.dataframe(insuf[COLS_RELATORIO].round(2), use_container_width=True)
        else:
            st.info("✅ Nenhuma ficha com saldo insuficiente encontrada nesta seleção.")
            
        st.subheader("🟢 Relatório de Suficiência de Saldo")
        st.dataframe(sufic[COLS_RELATORIO].round(2), use_container_width=True)
        
        st.divider()
        st.subheader("📊 Resumo por Atividade")
        resumo = df_final.groupby(['Código', 'Atividade']).agg(
            Saldo_Total=('Saldo com Reserva', 'sum'),
            Total_Projetado=('Total Projetado', 'sum'),
            Saldo_Final=('Saldo Final', 'sum')
        ).reset_index()
        resumo['Situação'] = resumo['Saldo_Final'].apply(lambda x: "✅ Suficiente" if x >= 0 else "🔴 Insuficiente")
        st.dataframe(resumo.round(2), use_container_width=True)

# ============================================================
# ABA 2 — SIMULADOR DE ACRÉSCIMO / DECRÉSCIMO
# ============================================================
with aba2:
    st.header("👤 Simulador de Acréscimo / Decréscimo de Servidor")
    st.markdown("Informe os valores base da alteração na folha para simular o impacto até o fim do exercício.")

    c1, c2 = st.columns(2)
    with c1:
        sim_atividade  = st.text_input("Código da Atividade (ex: 2104):", max_chars=4)
        sim_operacao   = st.radio("Operação:", ["➕ Acréscimo (Nova Contratação)", "➖ Decréscimo (Exoneração/Aposentadoria)"])
        sim_mes        = st.selectbox("Mês de início da mudança:", list(MESES.keys()), key="sim_mes")

    with c2:
        st.markdown("**Valores Fixos Mensais a Simular:**")
        sim_3190_11  = st.number_input("Vencimentos - 3.1.90.11 (R$)",  min_value=0.0, step=100.0)
        sim_3190_04  = st.number_input("Subsídios/Temporários - 3.1.90.04 (R$)", min_value=0.0, step=100.0)
        sim_3390_46  = st.number_input("Auxílio Alimentação - 3.3.90.46 (R$)",      min_value=0.0, step=50.0)
        tipo_encargo = st.selectbox("Regime Previdenciário:", ["RPPS Próprio (3.1.91.13)", "INSS Geral (3.1.90.13)"])

    if st.button("📊 Processar Simulação", type="primary"):
        if not sim_atividade:
            st.error("Por favor, informe o Código da Atividade para a simulação.")
        else:
            sinal = 1 if "Acréscimo" in sim_operacao else -1
            meses_impacto = 13 - MESES[sim_mes] 
            
            total_mensal_salarial = sim_3190_11 + sim_3190_04
            encargo_mensal = total_mensal_salarial * 0.22
            total_mensal = total_mensal_salarial + sim_3390_46 + encargo_mensal
            
            projecao_mensal_total = total_mensal * meses_impacto * sinal
            prov_13 = (total_mensal_salarial / 12) * meses_impacto * sinal
            prov_ferias = ((total_mensal_salarial / 3) / 12) * meses_impacto * sinal
            encargo_extra = (abs(prov_13) + abs(prov_ferias)) * 0.22 * sinal
            impacto_final = projecao_mensal_total + prov_13 + prov_ferias + encargo_extra
            
            st.divider()
            st.subheader(f"📋 Resultado da Simulação — Atividade: {sim_atividade}")
            st.markdown(f"**Impacto considerado por:** {meses_impacto} meses (De {sim_mes} até Dezembro).")
            
            dados_simulacao = [
                {"Item": "Projeção Salários Mensais", "Valor": projecao_mensal_total - (sim_3390_46 * meses_impacto * sinal) - (encargo_mensal * meses_impacto * sinal)},
                {"Item": "Projeção Auxílio Alimentação", "Valor": sim_3390_46 * meses_impacto * sinal},
                {"Item": "Provisão 13º Proporcional", "Valor": prov_13},
                {"Item": "Provisão Férias Proporcionais", "Valor": prov_ferias},
                {"Item": f"Encargos Patronais ({tipo_encargo.split(' ')[0]})", "Valor": (encargo_mensal * meses_impacto * sinal) + encargo_extra}
            ]
            
            df_simulacao = pd.DataFrame(dados_simulacao)
            
            col_res1, col_res2 = st.columns([2, 1])
            with col_res1:
                st.dataframe(df_simulacao.style.format({"Valor": "R$ {:,.2f}"}), hide_index=True, use_container_width=True)
            with col_res2:
                status_cor = "🟢 Redução de" if sinal < 0 else "🔴 Custo de"
                st.metric(label=f"Impacto Financeiro Final ({status_cor})", value=fmt_br(impacto_final))
                
            st.info("💡 **Dica Didática:** Este valor representa a alteração real no saldo da dotação no fim do exercício caso esta movimentação seja efetivada.")