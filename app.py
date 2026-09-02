import streamlit as st
import sqlite3
import io
from datetime import date
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# =====================================================================
# INICIALIZAÇÃO DO BANCO DE DADOS
# =====================================================================
def iniciar_banco():
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cadeias_custodia_geradas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            os_num TEXT,
            empreendimento TEXT,
            matriz TEXT,
            data_geracao TEXT,
            conteudo_pdf BLOB,
            conteudo_docx BLOB
        )
    ''')
    conexao.commit()
    conexao.close()

iniciar_banco()

def deletar_cc(id_registro):
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute("DELETE FROM cadeias_custodia_geradas WHERE id = ?", (id_registro,))
    conexao.commit()
    conexao.close()

# =====================================================================
# CONFIGURAÇÃO GERAL E CSS CORPORATIVO
# =====================================================================
st.set_page_config(page_title="Sistema de Amostragem Integrado", layout="wide")
st.markdown("""
    <style>
    :root { --azul-neon: #00FFFF; --azul-neon-escuro: #00cccc; }
    .stButton>button { background-color: var(--azul-neon); color: black; font-weight: bold; border-radius: 5px; }
    .stButton>button:hover { background-color: var(--azul-neon-escuro); color: white; }
    h1, h2, h3 { color: var(--azul-neon); }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# FUNÇÕES DE GERAÇÃO COM CABEÇALHO DE REPETIÇÃO
# =====================================================================
def cabecalho_pdf(canvas, doc, titulo, empreendimento, os_num, matriz):
    """Desenha o cabeçalho fixo no topo de todas as páginas do PDF"""
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawCentredString(doc.pagesize[0] / 2.0, doc.pagesize[1] - 40, titulo)
    canvas.setFont("Helvetica", 9)
    canvas.drawString(40, doc.pagesize[1] - 65, f"Empreendimento: {empreendimento}")
    canvas.drawString(40, doc.pagesize[1] - 80, f"OS N°: {os_num} | Matriz: {matriz}")
    canvas.line(40, doc.pagesize[1] - 90, doc.pagesize[0] - 40, doc.pagesize[1] - 90)
    canvas.restoreState()

def configurar_word_paisagem(doc):
    """Converte o documento Word para orientação Paisagem (ex: EAT Fumaça/Isocinética)"""
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

def cabecalho_word(doc, titulo, empreendimento, os_num, matriz):
    """Injeta o cabeçalho no header nativo do Word para repetição automática em novas páginas"""
    header = doc.sections[0].header
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.text = f"{titulo}\nEmpreendimento: {empreendimento} | OS N°: {os_num} | Matriz: {matriz}"
    p.style.font.size = Pt(9)
    p.style.font.bold = True

# =====================================================================
# MÓDULO: CADEIAS DE CUSTÓDIA
# =====================================================================
st.title("Gestão de Cadeias de Custódia")
tab_gerar, tab_historico = st.tabs(["Gerar Documento", "Documentos Gerados"])

with tab_gerar:
    st.subheader("1. Identificação do Empreendimento")
    col1, col2, col3 = st.columns(3)
    with col1:
        empreendimento = st.text_input("Empreendimento")
        endereco = st.text_input("Endereço")
    with col2:
        cod_cliente = st.text_input("Cód. do Cliente")
        responsavel = st.text_input("Responsável")
    with col3:
        os_num = st.text_input("OS N°")
        contato = st.text_input("Contato (DDD)")

    st.subheader("2. Matriz e Roteamento Operacional")
    matriz = st.selectbox("Selecione a Matriz de Coleta", [
        "Água, Efluente e Solo (Físico-Químico)", 
        "Dosimetria de Ruído (DOS)", 
        "Emissões Atmosféricas (Isocinética - Paisagem)"
    ])
    
    # Campo Data/Hora segregados e padronizados
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_coleta = st.date_input("Data Principal da Coleta", format="DD/MM/YYYY")
    with col_d2:
        hora_coleta = st.time_input("Horário de Início (Opcional)")

    st.markdown("---")
    st.subheader("3. Detalhamento dos Pontos (Máx: 100)")
    
    # Limite configurado para suportar alto volume de amostragem
    qtd_pontos = st.number_input("Quantidade de Pontos/Amostras", min_value=1, max_value=100, step=1)
    dados_pontos = []

    # RENDERIZAÇÃO CONDICIONAL POR MATRIZ
    if matriz == "Água, Efluente e Solo (Físico-Químico)":
        orientacao_pdf = A4
        for i in range(int(qtd_pontos)):
            with st.expander(f"Ponto {i+1} - Parâmetros In Loco"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    id_p = st.text_input("ID do Ponto", key=f"a_id_{i}")
                    ph = st.text_input("pH", key=f"a_ph_{i}")
                with c2:
                    temp = st.text_input("Temperatura (°C)", key=f"a_temp_{i}")
                    cond = st.text_input("Condutividade (μS/cm)", key=f"a_cond_{i}")
                with c3:
                    od = st.text_input("OD (ppm/%)", key=f"a_od_{i}")
                    sal = st.text_input("Salinidade (ppt/%)", key=f"a_sal_{i}")
                dados_pontos.append([id_p, ph, temp, cond, od, sal])
        cabecalho_tabela = ["ID", "pH", "Temp (°C)", "Cond", "OD", "Salinidade"]

    elif matriz == "Dosimetria de Ruído (DOS)":
        orientacao_pdf = A4
        for i in range(int(qtd_pontos)):
            with st.expander(f"Colaborador {i+1} (GHE)"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    nome = st.text_input("Nome", key=f"d_nome_{i}")
                    setor = st.text_input("Setor/Cargo", key=f"d_setor_{i}")
                with c2:
                    jornada = st.text_input("Jornada", key=f"d_jor_{i}")
                    fonte = st.text_input("Fonte Geradora", key=f"d_fonte_{i}")
                with c3:
                    prot = st.selectbox("Proteção Auditiva?", ["SIM", "NÃO"], key=f"d_prot_{i}")
                    t_exp = st.text_input("Tempo Exp.", key=f"d_texp_{i}")
                dados_pontos.append([nome, setor, jornada, fonte, prot, t_exp])
        cabecalho_tabela = ["Colaborador", "Setor/Cargo", "Jornada", "Fonte Geradora", "Proteção", "Tempo Exp."]

    elif matriz == "Emissões Atmosféricas (Isocinética - Paisagem)":
        orientacao_pdf = landscape(A4)
        for i in range(int(qtd_pontos)):
            with st.expander(f"Leitura EAT {i+1}"):
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    tempo = st.text_input("Tempo (min)", key=f"e_tem_{i}")
                    dist = st.text_input("Distância (cm)", key=f"e_dis_{i}")
                with c2:
                    vol = st.text_input("Vol. Gás (m³)", key=f"e_vol_{i}")
                    dp = st.text_input("ΔP (mmH2O)", key=f"e_dp_{i}")
                with c3:
                    dh = st.text_input("ΔH (mmH2O)", key=f"e_dh_{i}")
                    vac = st.text_input("Vácuo (cmHg)", key=f"e_vac_{i}")
                with c4:
                    t_gas = st.text_input("Temp Gasômetro", key=f"e_tgas_{i}")
                    t_cham = st.text_input("Temp Chaminé", key=f"e_tcham_{i}")
                dados_pontos.append([f"Leitura {i+1}", tempo, dist, vol, dp, dh, vac, t_gas, t_cham])
        cabecalho_tabela = ["Ponto", "Tempo", "Dist.", "Vol.", "ΔP", "ΔH", "Vácuo", "T. Gas", "T. Cham"]

    st.markdown("---")
    st.subheader("4. Recepção e Inspeção")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        entregue = st.text_input("Entregue por")
        recebido = st.text_input("Recebido por")
    with col_r2:
        data_rec = st.date_input("Data Recepção", format="DD/MM/YYYY")
        hora_rec = st.time_input("Hora Recepção")

    if st.button("Gerar Documentos com Paginação Dinâmica", type="primary"):
        data_geracao = date.today().strftime("%d/%m/%Y")
        
        # ================= LÓGICA DOCX (COM CABEÇALHO REPETIDO) =================
        doc = Document()
        if orientacao_pdf == landscape(A4):
            configurar_word_paisagem(doc)
            
        cabecalho_word(doc, "CADEIA DE CUSTÓDIA", empreendimento, os_num, matriz)
        doc.add_heading(f'Matriz Operacional: {matriz}', level=1)
        
        # Inserção de tabela capaz de quebrar páginas
        t_param = doc.add_table(rows=1, cols=len(cabecalho_tabela))
        t_param.style = 'Table Grid'
        for col_idx, nome_col in enumerate(cabecalho_tabela):
            t_param.cell(0, col_idx).text = nome_col
            
        for linha_dados in dados_pontos:
            row_cells = t_param.add_row().cells
            for col_idx, dado in enumerate(linha_dados):
                row_cells[col_idx].text = str(dado)
                
        doc.add_paragraph(f"\nRecepção: Entregue por {entregue} | Recebido por {recebido}")
        doc.add_paragraph(f"Data/Hora: {data_rec.strftime('%d/%m/%Y')} às {hora_rec.strftime('%H:%M')}")
        
        buffer_docx = io.BytesIO()
        doc.save(buffer_docx)

        # ================= LÓGICA PDF (COM PAGINAÇÃO DINÂMICA REPORTLAB) =================
        buffer_pdf = io.BytesIO()
        pdf = BaseDocTemplate(buffer_pdf, pagesize=orientacao_pdf, rightMargin=30, leftMargin=30, topMargin=100, bottomMargin=30)
        
        def canvas_maker(canvas, doc):
            cabecalho_pdf(canvas, doc, "CADEIA DE CUSTÓDIA", empreendimento, os_num, matriz)

        # O PageTemplate garante que o cabeçalho apareça em TODAS as páginas criadas pelos 100 pontos
        template = PageTemplate(id='todas_paginas', frames=Frame(30, 30, orientacao_pdf[0]-60, orientacao_pdf[1]-130), onPage=canvas_maker)
        pdf.addPageTemplates([template])
        
        elementos = []
        estilos = getSampleStyleSheet()
        estilo_tabela = TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('FONTSIZE', (0,0), (-1,-1), 8 if orientacao_pdf == A4 else 9)
        ])

        dados_tabela_pdf = [cabecalho_tabela] + dados_pontos
        # Cálculo de colunas dinâmico para não quebrar a largura da folha
        largura_total = orientacao_pdf[0] - 60
        largura_col = largura_total / len(cabecalho_tabela)
        
        tabela_pdf = Table(dados_tabela_pdf, colWidths=[largura_col]*len(cabecalho_tabela), repeatRows=1)
        tabela_pdf.setStyle(estilo_tabela)
        elementos.append(tabela_pdf)
        
        elementos.append(Spacer(1, 20))
        elementos.append(Paragraph("<b>Recepção e Inspeção:</b>", estilos['Normal']))
        elementos.append(Paragraph(f"Entregue por {entregue} | Recebido por {recebido} em {data_rec.strftime('%d/%m/%Y')} às {hora_rec.strftime('%H:%M')}", estilos['Normal']))

        pdf.build(elementos)

        # ================= GRAVAÇÃO =================
        conexao = sqlite3.connect('laboratorio.db')
        cursor = conexao.cursor()
        cursor.execute('''
            INSERT INTO cadeias_custodia_geradas (os_num, empreendimento, matriz, data_geracao, conteudo_pdf, conteudo_docx)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (os_num, empreendimento, matriz, data_geracao, buffer_pdf.getvalue(), buffer_docx.getvalue()))
        conexao.commit()
        conexao.close()
        
        st.success("Documento gerado com paginação dinâmica! Verifique a aba 'Documentos Gerados'.")

with tab_historico:
    st.subheader("Cadeias de Custódia Salvas")
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute("SELECT id, os_num, empreendimento, matriz, data_geracao, conteudo_pdf, conteudo_docx FROM cadeias_custodia_geradas")
    salvos = cursor.fetchall()
    conexao.close()
    
    if salvos:
        for cc in salvos:
            id_cc, os_num_db, emp_db, matriz_db, data_db, pdf_bytes, docx_bytes = cc
            st.markdown(f"**OS N° {os_num_db}** - {emp_db} ({matriz_db}) | Gerado em: {data_db}")
            col_d1, col_d2, col_del = st.columns([2, 2, 1])
            with col_d1:
                st.download_button("📄 Baixar PDF", data=pdf_bytes, file_name=f"CC_{os_num_db}.pdf", mime="application/pdf", key=f"p_{id_cc}")
            with col_d2:
                st.download_button("📝 Baixar Word", data=docx_bytes, file_name=f"CC_{os_num_db}.docx", key=f"w_{id_cc}")
            with col_del:
                st.button("Excluir", key=f"d_{id_cc}", on_click=deletar_cc, args=(id_cc,))
            st.markdown("---")
    else:
        st.info("Nenhum registro encontrado.")