import streamlit as st
from datetime import date
import io
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# CONFIGURAÇÃO GERAL E CSS CORPORATIVO
st.set_page_config(page_title="Sistema de Amostragem Integrado", layout="wide")

st.markdown("""
    <style>
    :root {
        --azul-neon: #00FFFF;
        --azul-neon-escuro: #00cccc;
    }
    .stButton>button {
        background-color: var(--azul-neon);
        color: black;
        font-weight: bold;
        border-radius: 5px;
        border: 1px solid var(--azul-neon-escuro);
    }
    .stButton>button:hover {
        background-color: var(--azul-neon-escuro);
        color: white;
        border-color: var(--azul-neon);
    }
    h1, h2, h3 {
        color: var(--azul-neon);
    }
    .st-expander {
        border-color: var(--azul-neon) !important;
    }
    </style>
""", unsafe_allow_html=True)

# GERENCIAMENTO DE ESTADO PARA ARQUIVOS E HISTÓRICO
if 'arquivos_os' not in st.session_state:
    st.session_state['arquivos_os'] = []
if 'cadeias_geradas' not in st.session_state:
    st.session_state['cadeias_geradas'] = []

# NAVEGAÇÃO LATERAL
st.sidebar.title("Navegação do Sistema")
aba_selecionada = st.sidebar.radio(
    "Selecione o Módulo:",
    ["Ordem de Serviço", "Planos de Amostragem", "Cadeias de Custódia"]
)
st.sidebar.markdown("---")

# =====================================================================
# MÓDULO 1: ORDEM DE SERVIÇO
# =====================================================================
if aba_selecionada == "Ordem de Serviço":
    st.title("Gestão de Ordens de Serviço (Comercial)")
    st.write("Anexe as propostas comerciais enviadas para estruturação da base de dados.")
    
    arquivo_upload = st.file_uploader("Carregar Proposta/OS (Obrigatório: PDF)", type=["pdf"], accept_multiple_files=False)
    
    if st.button("Salvar Arquivo no Sistema"):
        if arquivo_upload:
            dados_arquivo = {
                "nome": arquivo_upload.name,
                "tamanho": f"{arquivo_upload.size / 1024:.2f} KB",
                "data_anexo": date.today().strftime("%d/%m/%Y"),
                "conteudo_bytes": arquivo_upload.getvalue()
            }
            st.session_state['arquivos_os'].append(dados_arquivo)
            st.success(f"Arquivo '{arquivo_upload.name}' anexado com sucesso!")
        else:
            st.warning("Selecione um arquivo PDF antes de salvar.")
            
    st.markdown("---")
    st.subheader("Histórico de Ordens de Serviço Anexadas")
    
    if st.session_state['arquivos_os']:
        for i, arq in enumerate(st.session_state['arquivos_os']):
            col_info, col_btn = st.columns([3, 1])
            with col_info:
                st.markdown(f"**{arq['nome']}** (Tamanho: {arq['tamanho']} | Data: {arq['data_anexo']})")
            with col_btn:
                st.download_button(
                    label="Baixar PDF",
                    data=arq['conteudo_bytes'],
                    file_name=arq['nome'],
                    mime="application/pdf",
                    key=f"download_os_{i}"
                )
    else:
        st.info("Nenhuma Ordem de Serviço foi anexada até o momento.")

# =====================================================================
# MÓDULO 2: PLANOS DE AMOSTRAGEM
# =====================================================================
elif aba_selecionada == "Planos de Amostragem":
    st.title("Geração de Planos de Amostragem")
    st.write("Preencha as informações basilares para emitir o documento pré-operacional.")
    
    with st.form("form_plano"):
        num_documento = st.text_input("Numeração do Documento (OS/Plano)")
        nome_empreendimento = st.text_input("Nome do Empreendimento")
        endereco_empreendimento = st.text_input("Endereço do Empreendimento")
        responsavel_empreendimento = st.text_input("Responsável (Amostrador/Supervisor)")
        
        gerar_plano = st.form_submit_button("Compilar Plano de Amostragem")
        
    if gerar_plano:
        if not num_documento or not nome_empreendimento:
            st.error("Preencha ao menos a Numeração do Documento e o Nome do Empreendimento.")
        else:
            buffer_pdf = io.BytesIO()
            c = canvas.Canvas(buffer_pdf, pagesize=A4)
            c.setFont("Helvetica-Bold", 16)
            c.drawCentredString(A4[0]/2, 800, "PLANO DE AMOSTRAGEM - PRÉ-OPERAÇÃO")
            c.setFont("Helvetica", 12)
            c.drawString(50, 760, f"Documento N°: {num_documento}")
            c.drawString(50, 740, f"Empreendimento: {nome_empreendimento}")
            c.drawString(50, 720, f"Endereço: {endereco_empreendimento}")
            c.drawString(50, 700, f"Responsável: {responsavel_empreendimento}")
            c.save()
            
            doc = Document()
            titulo_plano = doc.add_heading('PLANO DE AMOSTRAGEM', 0)
            titulo_plano.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f"Documento N°: {num_documento}")
            doc.add_paragraph(f"Empreendimento: {nome_empreendimento}")
            doc.add_paragraph(f"Endereço: {endereco_empreendimento}")
            doc.add_paragraph(f"Responsável: {responsavel_empreendimento}")
            buffer_docx = io.BytesIO()
            doc.save(buffer_docx)
            
            st.success("Plano de Amostragem gerado com sucesso!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📄 Baixar Plano (PDF)", 
                    data=buffer_pdf.getvalue(), 
                    file_name=f"Plano_{num_documento}.pdf", 
                    mime="application/pdf"
                )
            with col2:
                st.download_button(
                    label="📝 Baixar Plano (Word)", 
                    data=buffer_docx.getvalue(), 
                    file_name=f"Plano_{num_documento}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# =====================================================================
# MÓDULO 3: CADEIAS DE CUSTÓDIA
# =====================================================================
elif aba_selecionada == "Cadeias de Custódia":
    st.title("Gestão de Cadeias de Custódia")
    
    tab_gerar, tab_historico = st.tabs(["Gerar Cadeia de Custódia", "Cadeias de Custódia Geradas"])
    
    with tab_gerar:
        st.subheader("Identificação do Empreendimento")
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

        st.subheader("Identificação da Amostragem")
        col4, col5, col6 = st.columns(3)
        with col4:
            matriz = st.selectbox("Matriz da Coleta", ["Água", "Ar", "Emissão Atmosférica", "Efluente", "Solo"])
        with col5:
            if matriz == "Água":
                submatriz = st.selectbox("Submatriz (Água)", ["ACH", "ASP", "EFL"])
            else:
                submatriz = "N/A"
                st.text_input("Submatriz", value="Não aplicável", disabled=True)
        with col6:
            data_coleta = st.date_input("Data da Coleta")
            
        hoje = date.today()
        coleta_futura = hoje < data_coleta
        
        if coleta_futura:
            st.error(f"⚠️ Acesso Restrito: A Data da Coleta ({data_coleta.strftime('%d/%m/%Y')}) é no futuro. O preenchimento da inspeção e das medições in loco ficará bloqueado até o dia da coleta.")
        else:
            st.success(f"✅ Data Validada: Acesso liberado para registro de parâmetros.")

        st.markdown("---")
        st.subheader("Parâmetros das medições in loco")
        
        qtd_pontos = st.number_input("Quantidade de Pontos Amostrados", min_value=1, step=1, disabled=coleta_futura)
        dados_pontos = []
        
        for i in range(int(qtd_pontos)):
            with st.expander(f"Parâmetros In Loco - Ponto de Coleta {i+1}", expanded=not coleta_futura):
                pc_col1, pc_col2, pc_col3 = st.columns(3)
                with pc_col1:
                    id_ponto = st.text_input("Identificação do Ponto", key=f"id_{i}", disabled=coleta_futura)
                    ph = st.text_input("pH", key=f"ph_{i}", disabled=coleta_futura)
                    temp = st.text_input("Temperatura da Amostra (°C)", key=f"temp_{i}", disabled=coleta_futura)
                with pc_col2:
                    condutividade = st.text_input("Condutividade Elétrica (μS/cm)", key=f"cond_{i}", disabled=coleta_futura)
                    oxigenio = st.text_input("Oxigênio Dissolvido (ppm/%)", key=f"od_{i}", disabled=coleta_futura)
                    solidos = st.text_input("Sólidos Totais Dissolvidos (ppm)", key=f"std_{i}", disabled=coleta_futura)
                with pc_col3:
                    salinidade = st.text_input("Salinidade (ppt/%)", key=f"sal_{i}", disabled=coleta_futura)
                    resistividade = st.text_input("Resistividade (Ω cm/kΩ cm)", key=f"res_{i}", disabled=coleta_futura)
                    potencial_oxido = st.text_input("Potencial de Óxido/Redução (mV)", key=f"orp_{i}", disabled=coleta_futura)
                    
                dados_pontos.append({
                    "id": id_ponto, "ph": ph, "temp": temp, "condutividade": condutividade, 
                    "oxigenio": oxigenio, "solidos": solidos, "salinidade": salinidade, 
                    "resistividade": resistividade, "potencial_oxido": potencial_oxido
                })

        st.markdown("---")
        st.subheader("Recepção e Inspeção da Amostra")
        
        rec_col1, rec_col2, rec_col3 = st.columns(3)
        with rec_col1:
            entregue_por = st.text_input("Entregue por", disabled=coleta_futura)
            recebido_triagem = st.text_input("Recebido por (Triagem)", disabled=coleta_futura)
        with rec_col2:
            data_hora_recepcao = st.text_input("Data e Hora da Recepção", disabled=coleta_futura)
            temperatura_recepcao = st.text_input("Temperatura de Recepção", disabled=coleta_futura)
        with rec_col3:
            desvio = st.text_input("Desvio? (Sim/Não - Se sim, qual?)", disabled=coleta_futura)
            
        observacoes = st.text_area("Observações", disabled=coleta_futura)
        
        st.markdown("Envio para Ensaio")
        ensaio_col1, ensaio_col2 = st.columns(2)
        with ensaio_col1:
            fq_recebido_por = st.text_input("Recebido por (Físico-Químico)", disabled=coleta_futura)
            fq_data_hora = st.text_input("Data e Hora (Físico-Químico)", disabled=coleta_futura)
        with ensaio_col2:
            micro_recebido_por = st.text_input("Recebido por (Microbiologia)", disabled=coleta_futura)
            micro_data_hora = st.text_input("Data e Hora (Microbiologia)", disabled=coleta_futura)

        st.markdown("---")
        if not coleta_futura:
            if st.button("Gerar Documentos da Cadeia de Custódia", type="primary"):
                # ================= LÓGICA DOCX =================
                doc_cc = Document()
                titulo_cc = doc_cc.add_heading('CADEIA DE CUSTÓDIA', 0)
                titulo_cc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc_cc.add_heading('Identificação do Empreendimento', level=1)
                t_emp = doc_cc.add_table(rows=3, cols=2)
                t_emp.style = 'Table Grid'
                t_emp.cell(0,0).text = f"Empreendimento: {empreendimento}"
                t_emp.cell(0,1).text = f"Cód. Cliente: {cod_cliente}"
                t_emp.cell(1,0).text = f"Endereço: {endereco}"
                t_emp.cell(1,1).text = f"Responsável: {responsavel}"
                t_emp.cell(2,0).text = f"OS N°: {os_num}"
                t_emp.cell(2,1).text = f"Contato: {contato}"
                
                doc_cc.add_paragraph("")
                doc_cc.add_heading('Identificação da Amostragem', level=1)
                t_amos = doc_cc.add_table(rows=1, cols=3)
                t_amos.style = 'Table Grid'
                t_amos.cell(0,0).text = f"Matriz: {matriz}"
                t_amos.cell(0,1).text = f"Submatriz: {submatriz}"
                t_amos.cell(0,2).text = f"Data da Coleta: {data_coleta.strftime('%d/%m/%Y')}"
                
                doc_cc.add_paragraph("")
                doc_cc.add_heading('Parâmetros das medições in loco', level=1)
                num_cols = len(dados_pontos) + 1
                t_param = doc_cc.add_table(rows=9, cols=num_cols)
                t_param.style = 'Table Grid'
                
                t_param.cell(0,0).text = "Parâmetros / Pontos"
                for idx, pto in enumerate(dados_pontos):
                    t_param.cell(0, idx+1).text = f"Pt {idx+1}: {pto['id']}"
                
                parametros_nomes = [
                    ("pH", "ph"), ("Temp. (°C)", "temp"), ("Condutividade (μS/cm)", "condutividade"),
                    ("Oxigênio Dissolv. (%)", "oxigenio"), ("STD (ppm)", "solidos"), 
                    ("Salinidade", "salinidade"), ("Resistividade", "resistividade"), ("Pot. Óxido/Redução", "potencial_oxido")
                ]
                
                for row_idx, (nome_param, chave_param) in enumerate(parametros_nomes, start=1):
                    t_param.cell(row_idx, 0).text = nome_param
                    for col_idx, pto in enumerate(dados_pontos, start=1):
                        t_param.cell(row_idx, col_idx).text = pto[chave_param]
                
                doc_cc.add_paragraph("")
                doc_cc.add_heading('Recepção e Inspeção da Amostra', level=1)
                t_rec = doc_cc.add_table(rows=4, cols=2)
                t_rec.style = 'Table Grid'
                t_rec.cell(0,0).text = f"Entregue por: {entregue_por}"
                t_rec.cell(0,1).text = f"Recebido (Triagem): {recebido_triagem}"
                t_rec.cell(1,0).text = f"Data/Hora: {data_hora_recepcao}"
                t_rec.cell(1,1).text = f"Temperatura: {temperatura_recepcao}"
                t_rec.cell(2,0).text = f"Desvio: {desvio}"
                t_rec.cell(2,1).text = f"Observações: {observacoes}"
                t_rec.cell(3,0).text = f"Físico-Químico -> {fq_recebido_por} ({fq_data_hora})"
                t_rec.cell(3,1).text = f"Microbiologia -> {micro_recebido_por} ({micro_data_hora})"
                
                buffer_docx_cc = io.BytesIO()
                doc_cc.save(buffer_docx_cc)

                # ================= LÓGICA PDF =================
                buffer_pdf_cc = io.BytesIO()
                pdf = SimpleDocTemplate(buffer_pdf_cc, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                elementos = []
                estilos = getSampleStyleSheet()
                estilo_titulo = ParagraphStyle(name='CenterTitle', parent=estilos['Heading1'], alignment=1)
                estilo_sub = estilos['Heading3']
                estilo_tabela = TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
                    ('FONTSIZE', (0,0), (-1,-1), 9),
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey)
                ])

                elementos.append(Paragraph("CADEIA DE CUSTÓDIA", estilo_titulo))
                elementos.append(Spacer(1, 15))
                
                # Tabela Empreendimento
                elementos.append(Paragraph("Identificação do Empreendimento", estilo_sub))
                dados_emp = [
                    [f"Empreendimento: {empreendimento}", f"Cód. Cliente: {cod_cliente}"],
                    [f"Endereço: {endereco}", f"Responsável: {responsavel}"],
                    [f"OS N°: {os_num}", f"Contato: {contato}"]
                ]
                tabela_emp = Table(dados_emp, colWidths=[260, 260])
                tabela_emp.setStyle(estilo_tabela)
                elementos.append(tabela_emp)
                elementos.append(Spacer(1, 15))
                
                # Tabela Amostragem
                elementos.append(Paragraph("Identificação da Amostragem", estilo_sub))
                dados_amos = [[f"Matriz: {matriz}", f"Submatriz: {submatriz}", f"Data Coleta: {data_coleta.strftime('%d/%m/%Y')}"]]
                tabela_amos = Table(dados_amos, colWidths=[173, 173, 174])
                tabela_amos.setStyle(estilo_tabela)
                elementos.append(tabela_amos)
                elementos.append(Spacer(1, 15))
                
                # Tabela Parâmetros
                elementos.append(Paragraph("Parâmetros das medições in loco", estilo_sub))
                dados_tabela_param = [["Parâmetros / Pontos"]]
                for pto in dados_pontos:
                    dados_tabela_param[0].append(pto['id'])
                for nome_param, chave_param in parametros_nomes:
                    linha = [nome_param]
                    for pto in dados_pontos:
                        linha.append(pto[chave_param])
                    dados_tabela_param.append(linha)
                
                tabela_param = Table(dados_tabela_param)
                tabela_param.setStyle(estilo_tabela)
                elementos.append(tabela_param)
                elementos.append(Spacer(1, 15))
                
                # Tabela Recepção
                elementos.append(Paragraph("Recepção e Inspeção da Amostra", estilo_sub))
                dados_rec = [
                    [f"Entregue por: {entregue_por}", f"Recebido (Triagem): {recebido_triagem}"],
                    [f"Data/Hora: {data_hora_recepcao}", f"Temperatura: {temperatura_recepcao}"],
                    [f"Desvio: {desvio}", f"Observações: {observacoes}"],
                    [f"FQ: {fq_recebido_por} ({fq_data_hora})", f"Micro: {micro_recebido_por} ({micro_data_hora})"]
                ]
                tabela_rec = Table(dados_rec, colWidths=[260, 260])
                tabela_rec.setStyle(estilo_tabela)
                elementos.append(tabela_rec)
                
                pdf.build(elementos)

                # SALVAR NO HISTÓRICO
                st.session_state['cadeias_geradas'].append({
                    "os_num": os_num,
                    "empreendimento": empreendimento,
                    "matriz": matriz,
                    "data_geracao": date.today().strftime("%d/%m/%Y"),
                    "docx_bytes": buffer_docx_cc.getvalue(),
                    "pdf_bytes": buffer_pdf_cc.getvalue()
                })
                
                st.success("Cadeia de Custódia gerada com sucesso! Verifique a aba 'Cadeias de Custódia Geradas' para baixar os arquivos.")

    with tab_historico:
        st.subheader("Documentos Gerados na Sessão")
        if st.session_state['cadeias_geradas']:
            for i, cc_info in enumerate(st.session_state['cadeias_geradas']):
                st.markdown(f"**OS N° {cc_info['os_num']}** - {cc_info['empreendimento']} ({cc_info['matriz']}) | Gerado em: {cc_info['data_geracao']}")
                col_d1, col_d2 = st.columns([1, 1])
                with col_d1:
                    st.download_button(
                        label="📄 Baixar Cadeia de Custódia (PDF)", 
                        data=cc_info['pdf_bytes'], 
                        file_name=f"CC_{cc_info['os_num']}_{cc_info['matriz']}.pdf", 
                        mime="application/pdf",
                        key=f"dl_pdf_{i}"
                    )
                with col_d2:
                    st.download_button(
                        label="📝 Baixar Cadeia de Custódia (Word)", 
                        data=cc_info['docx_bytes'], 
                        file_name=f"CC_{cc_info['os_num']}_{cc_info['matriz']}.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_docx_{i}"
                    )
                st.markdown("---")
        else:
            st.info("Nenhuma Cadeia de Custódia foi emitida e salva no histórico durante esta sessão de acesso.")