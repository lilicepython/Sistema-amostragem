import streamlit as st
import sqlite3
import io
from datetime import date
from docxtpl import DocxTemplate

# =====================================================================
# INICIALIZAÇÃO DO BANCO DE DADOS
# =====================================================================
def iniciar_banco():
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ordens_servico_arquivos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            tamanho TEXT,
            data_anexo TEXT,
            conteudo_pdf BLOB
        )
    ''')
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

def deletar_os(id_registro):
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute("DELETE FROM ordens_servico_arquivos WHERE id = ?", (id_registro,))
    conexao.commit()
    conexao.close()

def deletar_cc(id_registro):
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute("DELETE FROM cadeias_custodia_geradas WHERE id = ?", (id_registro,))
    conexao.commit()
    conexao.close()

# =====================================================================
# CONFIGURAÇÃO GERAL E CSS
# =====================================================================
st.set_page_config(page_title="Sistema de Amostragem", layout="wide")
st.markdown("""
    <style>
    :root { --azul-neon: #00FFFF; --azul-neon-escuro: #00cccc; }
    .stButton>button { background-color: var(--azul-neon); color: black; font-weight: bold; border-radius: 5px; }
    .stButton>button:hover { background-color: var(--azul-neon-escuro); color: white; }
    h1, h2, h3 { color: var(--azul-neon); }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# NAVEGAÇÃO LATERAL
# =====================================================================
st.sidebar.title("Navegação do Sistema")
aba_selecionada = st.sidebar.radio(
    "Selecione o Módulo:",
    ["Ordem de Serviço", "Plano de Amostragem", "Cadeia de Custódia"]
)
st.sidebar.markdown("---")

# =====================================================================
# MÓDULO 1: ORDEM DE SERVIÇO
# =====================================================================
if aba_selecionada == "Ordem de Serviço":
    st.title("Gestão de Ordens de Serviço")
    
    arquivo_upload = st.file_uploader("Carregar Proposta/OS (Somente PDF)", type=["pdf"])
    
    if st.button("Salvar Arquivo"):
        if arquivo_upload:
            nome = arquivo_upload.name
            tamanho = f"{arquivo_upload.size / 1024:.2f} KB"
            data_anexo = date.today().strftime("%d/%m/%Y")
            conteudo = arquivo_upload.getvalue()
            
            conexao = sqlite3.connect('laboratorio.db')
            cursor = conexao.cursor()
            cursor.execute('INSERT INTO ordens_servico_arquivos (nome, tamanho, data_anexo, conteudo_pdf) VALUES (?, ?, ?, ?)', 
                           (nome, tamanho, data_anexo, conteudo))
            conexao.commit()
            conexao.close()
            st.success(f"Arquivo '{nome}' salvo com sucesso!")
            st.rerun()
        else:
            st.warning("Selecione um arquivo PDF.")
            
    st.markdown("---")
    st.subheader("Histórico de Arquivos Anexados")
    conexao = sqlite3.connect('laboratorio.db')
    cursor = conexao.cursor()
    cursor.execute("SELECT id, nome, tamanho, data_anexo, conteudo_pdf FROM ordens_servico_arquivos ORDER BY id DESC")
    arquivos = cursor.fetchall()
    conexao.close()
    
    if arquivos:
        for arq in arquivos:
            id_arq, nome_arq, tam_arq, data_arq, pdf_bytes = arq
            c1, c2, c3 = st.columns([6, 2, 2])
            with c1:
                st.markdown(f"**{nome_arq}** | {tam_arq} | {data_arq}")
            with c2:
                st.download_button("Baixar PDF", data=pdf_bytes, file_name=nome_arq, mime="application/pdf", key=f"dl_os_{id_arq}")
            with c3:
                st.button("Excluir", key=f"del_os_{id_arq}", on_click=deletar_os, args=(id_arq,))
            st.markdown("---")

# =====================================================================
# MÓDULO 2: PLANO DE AMOSTRAGEM
# =====================================================================
elif aba_selecionada == "Plano de Amostragem":
    st.title("Geração de Planos de Amostragem")
    
    with st.form("form_plano"):
        st.subheader("1. Informações Básicas")
        c1, c2 = st.columns(2)
        num_documento = c1.text_input("Numeração do Documento")
        nome_empreendimento = c1.text_input("Nome do Empreendimento")
        endereco_empreendimento = c2.text_input("Local de Amostragem")
        responsavel_empreendimento = c2.text_input("Contato / Responsável")
        
        st.subheader("2. Escopo e Especificações")
        objetivo = st.text_area("Objetivo do Serviço")
        especificacoes = st.text_area("Especificações do Cliente e Requisitos Legais")
        duracao = st.text_input("Duração do Serviço")
        
        st.subheader("3. Matriz e Pontos de Amostragem")
        st.write("Selecione as matrizes aplicáveis:")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        m_humano = col_m1.checkbox("C. Humano")
        m_super = col_m1.checkbox("Superficial")
        m_sub = col_m2.checkbox("Subterrânea")
        m_resid = col_m2.checkbox("Residual")
        m_sed = col_m3.checkbox("Sedimentos")
        m_solo = col_m3.checkbox("Solo")
        m_residuo = col_m4.checkbox("Resíduo")
        
        c3, c4 = st.columns(2)
        qtd_pontos_plano = c3.text_input("Quantidade de Pontos")
        id_pontos = c4.text_input("Identificação dos Pontos")
        acessibilidade = c3.text_input("Acessibilidade")
        frequencia = c4.text_input("Frequência")
        data_hora_plano = st.text_input("Data e Hora")
        
        st.subheader("4. Garantia da Validade (QA/QC)")
        col_qa1, col_qa2, col_qa3 = st.columns(3)
        qa_equip = col_qa1.checkbox("Branco de Equipamento")
        qa_campo = col_qa2.checkbox("Branco de Campo")
        qa_viagem = col_qa3.checkbox("Branco de Viagem")
        qa_amostra = col_qa1.checkbox("Branco de Amostragem")
        qa_temp = col_qa2.checkbox("Branco de Temperatura")
        qa_duplicata = col_qa3.checkbox("Duplicata de Campo")
        
        st.subheader("5. Recepção no Laboratório")
        temp_rec = st.radio("Temperatura de Recebimento", ["≤6°C", "Não se aplica", "Nenhum"], index=2)
        
        st.subheader("6. Equipe Técnica")
        st.write("Preencha os colaboradores alocados para a coleta (Deixe em branco as linhas que não usar):")
        equipe_dados = []
        for i in range(3):
            ce1, ce2, ce3 = st.columns(3)
            nome_eq = ce1.text_input(f"Nome {i+1}", key=f"eq_n_{i}")
            cargo_eq = ce2.text_input(f"Cargo {i+1}", key=f"eq_c_{i}")
            etapa_eq = ce3.text_input(f"Etapa {i+1}", key=f"eq_e_{i}")
            if nome_eq.strip() != "":
                equipe_dados.append({"NOME": nome_eq, "CARGO": cargo_eq, "ETAPA": etapa_eq})
        
        submit_plano = st.form_submit_button("Gerar Plano de Amostragem", type="primary")
    
    if submit_plano:
        if not num_documento.strip() or not nome_empreendimento.strip():
            st.error("Preencha ao menos a Numeração e o Empreendimento.")
        else:
            try:
                doc = DocxTemplate("template_plano.docx")
                contexto_plano = {
                    "NUM_DOC": num_documento,
                    "EMPREENDIMENTO": nome_empreendimento,
                    "ENDERECO": endereco_empreendimento,
                    "RESPONSAVEL": responsavel_empreendimento,
                    "OBJETIVO": objetivo,
                    "ESPECIFICACOES": especificacoes,
                    "DURACAO": duracao,
                    "QTD_PONTOS": qtd_pontos_plano,
                    "ID_PONTOS": id_pontos,
                    "ACESSIBILIDADE": acessibilidade,
                    "FREQUENCIA": frequencia,
                    "DATA_HORA": data_hora_plano,
                    "M_HUMANO": m_humano,
                    "M_SUPER": m_super,
                    "M_SUB": m_sub,
                    "M_RESID": m_resid,
                    "M_SED": m_sed,
                    "M_SOLO": m_solo,
                    "M_RESIDUO": m_residuo,
                    "QA_EQUIP": qa_equip,
                    "QA_CAMPO": qa_campo,
                    "QA_VIAGEM": qa_viagem,
                    "QA_AMOSTRA": qa_amostra,
                    "QA_TEMP": qa_temp,
                    "QA_DUPLICATA": qa_duplicata,
                    "REC_6": temp_rec == "≤6°C",
                    "REC_NA": temp_rec == "Não se aplica",
                    "EQUIPE": equipe_dados
                }
                doc.render(contexto_plano)
                
                buf_docx = io.BytesIO()
                doc.save(buf_docx)
                
                st.success("Plano de Amostragem gerado com sucesso!")
                st.download_button("📝 Baixar Plano Preenchido (Word)", data=buf_docx.getvalue(), file_name=f"Plano_{num_documento}.docx")
            except Exception as e:
                st.error(f"Erro ao gerar documento. Certifique-se de que o arquivo 'template_plano.docx' foi enviado para o servidor. Detalhe: {e}")

# =====================================================================
# MÓDULO 3: CADEIA DE CUSTÓDIA
# =====================================================================
elif aba_selecionada == "Cadeia de Custódia":
    st.title("Gestão de Cadeias de Custódia")
    tab_gerar, tab_hist = st.tabs(["Gerar Cadeia de Custódia", "Cadeias de Custódia Geradas"])
    
    with tab_gerar:
        st.subheader("Identificação do Empreendimento")
        c1, c2, c3 = st.columns(3)
        empreendimento = c1.text_input("Empreendimento")
        endereco = c1.text_input("Endereço")
        cod_cliente = c2.text_input("Cód. Cliente")
        responsavel = c2.text_input("Responsável")
        os_num = c3.text_input("OS N° / PC N°")
        contato = c3.text_input("Contato")

        st.subheader("Identificação da Amostragem")
        c_m1, c_m2, c_m3 = st.columns(3)
        opcoes_matriz = {
            "Água e Efluentes": ["ACH", "ASP", "ASB", "EFL"],
            "Resíduos e Sedimentos": ["Resíduo Sólido", "Resíduo Líquido", "Sedimento", "Solo"]
        }
        matriz = c_m1.selectbox("Matriz", list(opcoes_matriz.keys()))
        submatriz = c_m2.selectbox("Submatriz", opcoes_matriz[matriz])
        data_coleta = c_m3.date_input("Data da Coleta", format="DD/MM/YYYY")
        
        hoje = date.today()
        bloqueio_temporal = hoje < data_coleta
        
        if bloqueio_temporal:
            st.error(f"⚠️ Bloqueio Temporal: A Data informada ({data_coleta.strftime('%d/%m/%Y')}) é no futuro. O preenchimento in loco e a emissão documental estão suspensos.")
        else:
            st.success("✅ Acesso Liberado para registro in loco.")
            
        st.markdown("---")
        st.subheader("Parâmetros das medições in loco")
        qtd_pontos = st.number_input("Quantidade de Pontos (Máx: 100)", min_value=1, max_value=100, step=1, disabled=bloqueio_temporal)
        
        dados_pontos = []
        if matriz == "Água e Efluentes":
            for i in range(int(qtd_pontos)):
                with st.expander(f"Ponto {i+1}", expanded=not bloqueio_temporal):
                    p1, p2, p3 = st.columns(3)
                    id_ponto = p1.text_input("ID do Ponto", key=f"id_{i}", disabled=bloqueio_temporal)
                    ph = p1.text_input("pH", key=f"ph_{i}", disabled=bloqueio_temporal)
                    temp = p1.text_input("Temp (°C)", key=f"t_{i}", disabled=bloqueio_temporal)
                    cond = p2.text_input("Condutividade", key=f"c_{i}", disabled=bloqueio_temporal)
                    od = p2.text_input("Oxigênio Dissolvido", key=f"o_{i}", disabled=bloqueio_temporal)
                    std = p2.text_input("STD", key=f"st_{i}", disabled=bloqueio_temporal)
                    sal = p3.text_input("Salinidade", key=f"sa_{i}", disabled=bloqueio_temporal)
                    res = p3.text_input("Resistividade", key=f"re_{i}", disabled=bloqueio_temporal)
                    orp = p3.text_input("ORP", key=f"or_{i}", disabled=bloqueio_temporal)
                    
                    dados_pontos.append({
                        "ID": id_ponto, "PH": ph, "TEMP": temp, "COND": cond, 
                        "OD": od, "STD": std, "SAL": sal, "RES": res, "ORP": orp
                    })
        else:
            for i in range(int(qtd_pontos)):
                with st.expander(f"Ponto {i+1}", expanded=not bloqueio_temporal):
                    p1, p2 = st.columns(2)
                    id_ponto = p1.text_input("ID do Ponto", key=f"id_{i}", disabled=bloqueio_temporal)
                    desc = p1.text_input("Descrição", key=f"d_{i}", disabled=bloqueio_temporal)
                    tipo = p1.text_input("Tipo", key=f"t_{i}", disabled=bloqueio_temporal)
                    massa = p2.text_input("Massa (Kg)", key=f"m_{i}", disabled=bloqueio_temporal)
                    loc = p2.text_input("Localização", key=f"l_{i}", disabled=bloqueio_temporal)
                    
                    dados_pontos.append({
                        "ID": id_ponto, "DESC": desc, "TIPO": tipo, "MASSA": massa, "LOC": loc
                    })

        st.markdown("---")
        st.subheader("Recepção e Inspeção da Amostra")
        r1, r2, r3 = st.columns(3)
        entregue = r1.text_input("Entregue por", disabled=bloqueio_temporal)
        recebido = r1.text_input("Recebido por", disabled=bloqueio_temporal)
        data_rec = r2.date_input("Data Recepção (Opcional)", format="DD/MM/YYYY", value=None, disabled=bloqueio_temporal)
        hora_rec = r2.time_input("Hora Recepção (Opcional)", value=None, disabled=bloqueio_temporal)
        temp_rec = r3.text_input("Temperatura Recepção", disabled=bloqueio_temporal)
        desvio = r3.text_input("Desvio? (Sim/Não)", disabled=bloqueio_temporal)
        
        if not bloqueio_temporal:
            if st.button("Gerar Cadeia de Custódia", type="primary"):
                str_data_rec = data_rec.strftime('%d/%m/%Y') if data_rec else "___/___/___"
                str_hora_rec = hora_rec.strftime('%H:%M') if hora_rec else "___:___"

                contexto_cc = {
                    "EMPREENDIMENTO": empreendimento,
                    "ENDERECO": endereco,
                    "COD_CLIENTE": cod_cliente,
                    "RESPONSAVEL": responsavel,
                    "OS_NUM": os_num,
                    "CONTATO": contato,
                    "MATRIZ": matriz,
                    "SUBMATRIZ": submatriz,
                    "DATA_COLETA": data_coleta.strftime('%d/%m/%Y'),
                    "ENTREGUE": entregue,
                    "RECEBIDO": recebido,
                    "DATA_REC": str_data_rec,
                    "HORA_REC": str_hora_rec,
                    "TEMP_REC": temp_rec,
                    "DESVIO": desvio,
                    "PONTOS": dados_pontos
                }

                try:
                    arquivo_template = "template_cc_agua.docx" if matriz == "Água e Efluentes" else "template_cc_solidos.docx"
                    doc = DocxTemplate(arquivo_template)
                    doc.render(contexto_cc)
                    
                    buf_docx = io.BytesIO()
                    doc.save(buf_docx)
                    
                    conexao = sqlite3.connect('laboratorio.db')
                    c = conexao.cursor()
                    c.execute('INSERT INTO cadeias_custodia_geradas (os_num, empreendimento, matriz, data_geracao, conteudo_pdf, conteudo_docx) VALUES (?,?,?,?,?,?)',
                              (os_num, empreendimento, matriz, date.today().strftime("%d/%m/%Y"), None, buf_docx.getvalue()))
                    conexao.commit()
                    conexao.close()
                    
                    st.success("Cadeia de Custódia gerada a partir do Template com sucesso! Acesse a aba ao lado.")
                except Exception as e:
                    st.error(f"Erro ao processar template. Certifique-se de que os arquivos 'template_cc_agua.docx' e 'template_cc_solidos.docx' foram carregados no GitHub. Detalhe: {e}")

    with tab_hist:
        st.subheader("Cadeias de Custódia Salvas")
        conexao = sqlite3.connect('laboratorio.db')
        c = conexao.cursor()
        c.execute("SELECT id, os_num, empreendimento, matriz, data_geracao, conteudo_docx FROM cadeias_custodia_geradas")
        salvos = c.fetchall()
        conexao.close()
        
        if salvos:
            for cc in salvos:
                id_cc, os_db, emp_db, mat_db, data_db, doc_b = cc
                c1, c2, c3 = st.columns([6, 2, 2])
                with c1: st.write(f"**OS {os_db}** | {emp_db} ({mat_db}) | {data_db}")
                with c2: 
                    if doc_b:
                        st.download_button("📝 Baixar Word", doc_b, f"CC_{os_db}.docx", key=f"doc_{id_cc}")
                with c3: 
                    st.button("Excluir", key=f"del_cc_{id_cc}", on_click=deletar_cc, args=(id_cc,))
                st.markdown("---")