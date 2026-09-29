import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import time
from datetime import datetime
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# 1. Configuração da página
st.set_page_config(
    page_title="Estoque GravoMark 02.08", 
    page_icon="LOGO PNG COM FUNDO.png", 
    layout="wide"
)

# 2. CSS Ultra-Compacto e Otimizado
st.markdown(
    """
    <style>
    header, [data-testid="stHeader"], #MainMenu, footer, [data-testid="stFooter"], [data-testid="stBottom"], .stAppEmbedFooter { display: none !important; }
    [data-testid="stStatusWidget"], [data-testid="stAppViewerToolbar"], .stAppToolbar, div[class*="viewerBadge"] { display: none !important; }
    .block-container { padding-top: 0.5rem !important; padding-bottom: 0rem !important; margin-top: 0rem !important; }
    [data-testid="stSidebar"] { padding-top: 0.5rem !important; }
    [data-testid="stSidebarHeader"] { display: none !important; }
    [data-testid="stElementToolbar"] { display: none !important; }
    h1, h2, h3 { margin-top: -0.5rem !important; padding-top: 0rem !important; margin-bottom: 0.5rem !important; }
    div[data-testid="stVerticalBlock"] > div { gap: 0.4rem !important; }
    </style>
    """,
    unsafe_allow_html=True
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def conectar_google():
    creds = None
    if "google_token" in st.secrets:
        token_info = dict(st.secrets["google_token"])
        creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    elif os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None

    if not creds:
        st.error("Erro nas credenciais de acesso ao Google. Verifique o registro em Secrets.")
        st.stop()

    client_sheets = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return client_sheets, drive_service

client_sheets, drive_service = conectar_google()

# --- IDs REAIS DO GOOGLE DRIVE / SHEETS ---
SPREADSHEET_ID = "1mnR2hraUpJm5KIQRLk4JOCU35zTKgPaelr02CjxJ248"
FOLDER_ENTRADA_ID = "14b0Dp4LEEftPMIUkVxDd_0JFKGJhRbWL"
FOLDER_SAIDA_ID = "1iFmbto3DIRKW83SdON-QaMXTDrfrRmcx"

@st.cache_data(ttl=60)
def ler_dados_planilha(nome_aba, usar_formula=False):
    aba = client_sheets.open_by_key(SPREADSHEET_ID).worksheet(nome_aba)
    if usar_formula:
        return aba.get_all_records(value_render_option='FORMULA')
    return aba.get_all_records()

def formatar_codigo_peca(codigo_bruto):
    apenas_numeros = "".join(filter(str.isdigit, str(codigo_bruto)))
    if len(apenas_numeros) == 14:
        return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:9]}.{apenas_numeros[9:]}"
    return str(codigo_bruto).strip()

def extrair_numeros(texto):
    return "".join(filter(str.isdigit, str(texto)))

def buscar_link_nf_existente_no_drive(num_nf, pasta_id):
    if not num_nf or num_nf == "S/N":
        return ""
    try:
        query = f"'{pasta_id}' in parents and (name contains '{num_nf}') and trashed=false"
        resultados = drive_service.files().list(q=query, fields="files(id, name)").execute().get('files', [])
        if resultados:
            arquivo_id = resultados[0]['id']
            url_bruta = f"https://drive.google.com/file/d/{arquivo_id}/view"
            return f'=HYPERLINK("{url_bruta}"; "📄 NF {num_nf}")'
    except Exception:
        pass
    return ""

def gerar_formula_saldo(codigo, nf_origem):
    f_entrada = f'SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{codigo}"; Movimentacoes!C:C; "{nf_origem}"; Movimentacoes!B:B; "Entrada")'
    f_retorno = f'SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{codigo}"; Movimentacoes!C:C; "{nf_origem}"; Movimentacoes!B:B; "Retorno Técnica")'
    f_tec = f'SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{codigo}"; Movimentacoes!C:C; "{nf_origem}"; Movimentacoes!B:B; "Área Técnica")'
    f_saida_novo = f'SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{codigo}"; Movimentacoes!H:H; "*{nf_origem}*"; Movimentacoes!B:B; "Saida")'
    f_saida_antigo = f'SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{codigo}"; Movimentacoes!C:C; "{nf_origem}"; Movimentacoes!B:B; "Saida"; Movimentacoes!H:H; "<>*(NFE:*")'
    
    return f'={f_entrada} + {f_retorno} - {f_saida_novo} - {f_saida_antigo} - {f_tec}'

def deletar_linhas_em_lote(sheet, indices_linhas):
    if not indices_linhas:
        return
    indices_unicos = sorted(list(set(indices_linhas)), reverse=True)
    requests = []
    for idx in indices_unicos:
        requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "ROWS",
                    "startIndex": idx - 1, 
                    "endIndex": idx
                }
            }
        })
    try:
        client_sheets.open_by_key(SPREADSHEET_ID).batch_update({"requests": requests})
    except Exception:
        for idx in indices_unicos:
            sheet.delete_rows(idx)

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if "form_version" not in st.session_state:
    st.session_state["form_version"] = 0

ver = st.session_state["form_version"]

try:
    col_vazia1, col_logo, col_vazia2 = st.sidebar.columns([1, 2, 1])
    with col_logo:
        st.image("LOGO PNG.png", use_container_width=True)
except Exception:
    pass

st.sidebar.title("Navegação")
aba = st.sidebar.radio("Ir para:", ["📦 Lançar Movimentação", "📊 Consultar Estoque", "📋 Histórico", "🛠️ Área Técnica"])

if aba == "📦 Lançar Movimentação":
    
    col_titulo, col_senha = st.columns([10, 1])
    with col_titulo:
        st.subheader("Lançamento Manual em Lote")
    with col_senha:
        if not st.session_state["autenticado"]:
            with st.popover("🔑"):
                with st.form("form_login"):
                    senha = st.text_input("Senha:", type="password")
                    submit = st.form_submit_button("Desbloquear", use_container_width=True)
                    if submit:
                        if senha == "1234":
                            st.session_state["autenticado"] = True
                            st.rerun()
                        else:
                            st.error("Senha incorreta!")
        else:
            if st.button("🔒 Sair"):
                st.session_state["autenticado"] = False
                st.rerun()

    if not st.session_state["autenticado"]:
        st.info("🔒 O sistema está bloqueado. Clique no ícone de chave (🔑) no canto superior direito para acessar.")
        st.stop() 

    dados_cadastro = ler_dados_planilha("Cadastro_Pecas")
    dados_estoque_atual = ler_dados_planilha("Estoque_Atual")

    lista_estoque_geral = []
    dict_cadastrados = {}
    
    if dados_cadastro:
        for row in dados_cadastro:
            cod = str(row.get('Codigo_Peca', '')).strip().replace("'", "")
            desc = str(row.get('Descricao', '')).strip()
            if cod:
                cod_sem_pontos = extrair_numeros(cod)
                if cod_sem_pontos and cod_sem_pontos != cod:
                    item_label = f"{cod} ({cod_sem_pontos}) | {desc}"
                else:
                    item_label = f"{cod} | {desc}"
                
                lista_estoque_geral.append(item_label)
                dict_cadastrados[cod] = desc
                if cod_sem_pontos:
                    dict_cadastrados[cod_sem_pontos] = desc
                    
    lista_estoque_geral = sorted(list(set(lista_estoque_geral)))

    col_tipo, col_nf_head, col_pdf_head = st.columns([2, 3, 4])
    
    with col_tipo:
        tipo_mov = st.radio("Operação:", ["Entrada", "Saida"], horizontal=True, key=f"radio_op_{ver}")
    with col_nf_head:
        num_nf_input = st.text_input("Número da NF ou Documento*", placeholder="Obrigatório", key=f"input_nf_{ver}")
        sem_nf_check = st.checkbox("🔓 Autorizar lançamento sem NF", key=f"check_sem_nf_{ver}")
        senha_autorizacao = ""
        if sem_nf_check:
            senha_autorizacao = st.text_input("Senha de Autorização:", type="password", key=f"senha_aut_{ver}")

    with col_pdf_head:
        arquivo_nf = st.file_uploader("Anexar PDF da NF", type=["pdf"], key=f"file_up_{ver}")

    st.caption(f"💡 Preencha a tabela de {tipo_mov} abaixo (5 linhas iniciais. Clique em + para mais):")

    if tipo_mov == "Entrada":
        df_template = pd.DataFrame([
            {"Peça Cadastrada": "", "Código Peça Nova (Se não existir)": "", "Descrição Peça Nova": "", "Quantidade": 1}
            for _ in range(5)
        ])

        editor_itens = st.data_editor(
            df_template,
            num_rows="dynamic",
            use_container_width=True,
            height=250,
            key=f"editor_entrada_{ver}",
            column_config={
                "Peça Cadastrada": st.column_config.SelectboxColumn(
                    "Selecione Peça Existente",
                    options=[""] + lista_estoque_geral,
                    width="large"
                ),
                "Código Peça Nova (Se não existir)": st.column_config.TextColumn("Ou digite Código Novo", width="medium"),
                "Descrição Peça Nova": st.column_config.TextColumn("Descrição da Peça Nova", width="large"),
                "Quantidade": st.column_config.NumberColumn("Qtd", min_value=1, step=1, default=1, width="small")
            }
        )

        if st.button("💾 Registrar TODOS os Itens no Estoque", use_container_width=True, type="primary"):
            if not num_nf_input.strip() and not sem_nf_check:
                st.error("🛑 **ERRO:** O número da Nota Fiscal é OBRIGATÓRIO!")
                st.stop()
            
            if sem_nf_check and senha_autorizacao != "1234":
                st.error("🛑 **SENHA INCORRETA:** Senha de autorização inválida.")
                st.stop()

            nf_formatada = "".join(filter(str.isdigit, num_nf_input)).lstrip('0') if num_nf_input else ""
            nf_formatada = nf_formatada if nf_formatada else "S/N"

            itens_validos = []
            novas_pecas_para_cadastrar = []

            for idx, row in editor_itens.iterrows():
                peca_existente = str(row.get("Peça Cadastrada", "") or "").strip()
                cod_novo = str(row.get("Código Peça Nova (Se não existir)", "") or "").strip()
                desc_nova = str(row.get("Descrição Peça Nova", "") or "").strip()
                qtd_item = int(row.get("Quantidade", 1) or 1)

                if peca_existente and peca_existente != "None":
                    partes = peca_existente.split(" | ", 1)
                    c_fin = partes[0].split(" (")[0].strip()
                    d_fin = partes[1].strip() if len(partes) > 1 else ""
                    itens_validos.append({"codigo": c_fin, "descricao": d_fin, "qtd": qtd_item})
                elif cod_novo and cod_novo != "None" and desc_nova and desc_nova != "None":
                    c_fin = formatar_codigo_peca(cod_novo)
                    itens_validos.append({"codigo": c_fin, "descricao": desc_nova, "qtd": qtd_item})
                    if c_fin not in dict_cadastrados:
                        novas_pecas_para_cadastrar.append([f"'{c_fin}", desc_nova])

            if not itens_validos:
                st.error("Preencha pelo menos um item válido na tabela acima.")
                st.stop()

            codigos_tabela = [item['codigo'] for item in itens_validos]
            if len(codigos_tabela) != len(set(codigos_tabela)):
                st.error("🛑 **ERRO DE DUPLICIDADE:** Você inseriu a mesma peça mais de uma vez na tabela acima. Por favor, some a quantidade e faça o lançamento em uma única linha.")
                st.stop()

            sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
            dados_est_cru = sheet_est.get_all_records()
            itens_duplicados = []
            
            for item in itens_validos:
                for row_e in dados_est_cru:
                    cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").strip()
                    nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                    if cod_e == item['codigo'] and nf_e == nf_formatada:
                        itens_duplicados.append(item['codigo'])

            if itens_duplicados:
                st.error(f"🛑 **DUPLICIDADE DETECTADA NO ESTOQUE:** A peça(s) {', '.join(set(itens_duplicados))} já foi cadastrada anteriormente para a mesma NF {nf_formatada}. O sistema bloqueia essas repetições para evitar falhas no saldo.")
                st.stop()

            with st.spinner("Processando lote de entrada..."):
                data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                data_hora_segura = f"'{data_hora_bruta}"
                
                sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                sheet_cad = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Cadastro_Pecas")

                if novas_pecas_para_cadastrar:
                    sheet_cad.insert_rows(novas_pecas_para_cadastrar, row=2, value_input_option='USER_ENTERED')

                link_arquivo = ""
                if arquivo_nf is not None:
                    pasta_id = FOLDER_ENTRADA_ID
                    nome_padronizado = f"NF {nf_formatada}.pdf" if nf_formatada != "S/N" else arquivo_nf.name

                    query = f"'{pasta_id}' in parents and name='{nome_padronizado}' and trashed=false"
                    resultados = drive_service.files().list(q=query, fields="files(id)").execute().get('files', [])

                    if resultados:
                        arquivo_id = resultados[0]['id']
                    else:
                        file_metadata = {'name': nome_padronizado, 'parents': [pasta_id]}
                        media = MediaIoBaseUpload(io.BytesIO(arquivo_nf.getvalue()), mimetype=arquivo_nf.type)
                        arquivo_salvo = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                        arquivo_id = arquivo_salvo.get('id')

                    url_bruta = f"https://drive.google.com/file/d/{arquivo_id}/view"
                    link_arquivo = f'=HYPERLINK("{url_bruta}"; "📄 NF {nf_formatada}")'
                else:
                    link_arquivo = buscar_link_nf_existente_no_drive(nf_formatada, FOLDER_ENTRADA_ID)

                if not link_arquivo:
                    link_arquivo = nf_formatada

                novas_linhas_mov = []
                novas_linhas_est = []

                for item in itens_validos:
                    c_seg = f"'{item['codigo']}"
                    novas_linhas_mov.append([data_hora_segura, "Entrada", nf_formatada, c_seg, item["descricao"], item["qtd"], link_arquivo, "", ""])
                    formula_saldo = gerar_formula_saldo(item["codigo"], nf_formatada)
                    novas_linhas_est.append([c_seg, item["descricao"], formula_saldo, nf_formatada, 0, data_hora_segura])

                if novas_linhas_mov:
                    sheet_mov.insert_rows(novas_linhas_mov, row=2, value_input_option='USER_ENTERED')
                if novas_linhas_est:
                    sheet_est.insert_rows(novas_linhas_est, row=2, value_input_option='USER_ENTERED')

                st.cache_data.clear()
                st.success(f"✅ Lote registrado com sucesso! {len(itens_validos)} itens adicionados ao estoque.")
                
                st.session_state["form_version"] += 1
                time.sleep(1.2)
                st.rerun()

    else:
        opcoes_saida_disponiveis = []
        for r in dados_estoque_atual:
            c_item = str(r.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
            q_item = pd.to_numeric(r.get('Quantidade_Atual', 0), errors='coerce') or 0
            
            raw_nf = str(r.get('Nota_Fiscal', '')).strip()
            nf_label = f"NF {raw_nf}" if raw_nf and raw_nf.upper() != "S/N" else "Sem NF (S/N)"
            desc_item = str(r.get('Descricao', '')).strip()
            
            if c_item and q_item > 0:
                c_sem_pontos = extrair_numeros(c_item)
                if c_sem_pontos and c_sem_pontos != c_item:
                    rotulo_peca = f"{c_item} ({c_sem_pontos})"
                else:
                    rotulo_peca = c_item
                
                opcoes_saida_disponiveis.append(f"{rotulo_peca} | {nf_label} | Saldo: {int(q_item)} un | {desc_item}")

        if not opcoes_saida_disponiveis:
            st.warning("⚠️ Não há peças com saldo disponível para dar saída no momento.")
        else:
            df_saida_template = pd.DataFrame([
                {"Selecione a Peça e NF de Origem": "", "Quantidade de Saída": 1}
                for _ in range(5)
            ])

            editor_saida = st.data_editor(
                df_saida_template,
                num_rows="dynamic",
                use_container_width=True,
                height=250,
                key=f"editor_saida_{ver}",
                column_config={
                    "Selecione a Peça e NF de Origem": st.column_config.SelectboxColumn(
                        "Selecione Peça e Lote de Saída",
                        options=[""] + sorted(opcoes_saida_disponiveis),
                        width="large"
                    ),
                    "Quantidade de Saída": st.column_config.NumberColumn("Quantidade", min_value=1, step=1, default=1, width="small")
                }
            )

            if st.button("📤 Confirmar Lote de Saída do Estoque", use_container_width=True, type="primary"):
                if not num_nf_input.strip() and not sem_nf_check:
                    st.error("🛑 **ERRO:** Número da NF / Documento de Saída é OBRIGATÓRIO!")
                    st.stop()

                if sem_nf_check and senha_autorizacao != "1234":
                    st.error("🛑 **SENHA INCORRETA:** Senha de autorização inválida.")
                    st.stop()

                itens_saida_validos = []
                for idx, row in editor_saida.iterrows():
                    opcao_sel = str(row.get("Selecione a Peça e NF de Origem", "") or "").strip()
                    qtd_s = int(row.get("Quantidade de Saída", 1) or 1)

                    if opcao_sel and opcao_sel != "None":
                        partes_s = opcao_sel.split(" | ")
                        cod_s = partes_s[0].split(" (")[0].strip()
                        
                        str_nf_parte = partes_s[1].strip()
                        if "NF " in str_nf_parte:
                            nf_origem_lote = str_nf_parte.replace("NF ", "").strip()
                        else:
                            nf_origem_lote = "S/N"

                        saldo_str = partes_s[2].replace("Saldo:", "").replace("un", "").strip()
                        saldo_atual = int(saldo_str)

                        desc_s = partes_s[3].strip() if len(partes_s) > 3 else ""
                        itens_saida_validos.append({
                            "codigo": cod_s, 
                            "nf_origem": nf_origem_lote, 
                            "descricao": desc_s, 
                            "qtd": qtd_s, 
                            "saldo_atual": saldo_atual
                        })

                if not itens_saida_validos:
                    st.error("Selecione pelo menos uma peça para dar saída.")
                else:
                    with st.spinner("Processando lote de saída automatizada..."):
                        data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        data_hora_segura = f"'{data_hora_bruta}"

                        nf_saida_doc = num_nf_input.strip() if num_nf_input.strip() else "S/N"

                        sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                        sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")

                        # Para saber se não devemos deletar a linha do Estoque por causa da área técnica
                        dados_tec_cru = ler_dados_planilha("Area_Tecnica")
                        dict_na_tecnica_saida = {}
                        if dados_tec_cru:
                            for rt in dados_tec_cru:
                                ct = str(rt.get('Codigo_Peca', '')).replace("'", "").strip()
                                nft = str(rt.get('Nota_Fiscal', '')).strip().lstrip('0')
                                qt = pd.to_numeric(rt.get('Quantidade', 0), errors='coerce') or 0
                                dict_na_tecnica_saida[f"{ct}_{nft}"] = dict_na_tecnica_saida.get(f"{ct}_{nft}", 0) + qt

                        link_saida = ""
                        if arquivo_nf is not None:
                            nome_padronizado = f"SAIDA_{nf_saida_doc}.pdf"
                            file_metadata = {'name': nome_padronizado, 'parents': [FOLDER_SAIDA_ID]}
                            media = MediaIoBaseUpload(io.BytesIO(arquivo_nf.getvalue()), mimetype=arquivo_nf.type)
                            
                            for tentativa in range(3):
                                try:
                                    arquivo_salvo = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                                    url_bruta = f"https://drive.google.com/file/d/{arquivo_salvo.get('id')}/view"
                                    link_saida = url_bruta
                                    break
                                except Exception:
                                    time.sleep(1)

                        nfs_origem_unicas = list(set([item['nf_origem'] for item in itens_saida_validos if item['nf_origem'] != "S/N"]))
                        mapa_links_entrada = {}
                        
                        if nfs_origem_unicas:
                            condicoes = " or ".join([f"name contains '{nf}'" for nf in nfs_origem_unicas])
                            query = f"'{FOLDER_ENTRADA_ID}' in parents and ({condicoes}) and trashed=false"
                            try:
                                resultados_drive = drive_service.files().list(q=query, fields="files(id, name)").execute().get('files', [])
                                for arq in resultados_drive:
                                    nums = extrair_numeros(arq['name'])
                                    if nums:
                                        url_bruta = f"https://drive.google.com/file/d/{arq['id']}/view"
                                        mapa_links_entrada[nums] = f'=HYPERLINK("{url_bruta}"; "📄 NF {nums}")'
                            except Exception:
                                pass

                        novas_linhas_mov_s = []
                        linhas_est_deletar = []
                        novas_linhas_est_s = []

                        dados_est_cru = sheet_est.get_all_records()

                        for item_s in itens_saida_validos:
                            c_seg = f"'{item_s['codigo']}"
                            
                            link_nf_entrada_origem = mapa_links_entrada.get(item_s['nf_origem'], "")
                            texto_col_g = link_nf_entrada_origem if link_nf_entrada_origem else item_s['nf_origem']
                            
                            label_saida_com_nfe = f'Doc Saída (NFE: {item_s["nf_origem"]})'
                            if link_saida:
                                label_saida_com_nfe = f'=HYPERLINK("{link_saida}"; "📄 Doc Saída (NFE: {item_s["nf_origem"]})")'

                            novas_linhas_mov_s.append([data_hora_segura, "Saida", nf_saida_doc, c_seg, item_s["descricao"], item_s["qtd"], texto_col_g, label_saida_com_nfe, ""])

                            for idx_e, row_e in enumerate(dados_est_cru):
                                cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
                                nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                                if cod_e == item_s['codigo'] and nf_e == item_s['nf_origem'].lstrip('0'):
                                    linhas_est_deletar.append(idx_e + 2)

                            saldo_restante = item_s["saldo_atual"] - item_s["qtd"]
                            qtd_tech_lote = dict_na_tecnica_saida.get(f"{item_s['codigo']}_{item_s['nf_origem']}", 0)
                            
                            # Condição inteligente para não apagar o item da planilha se ele estiver na técnica
                            if saldo_restante > 0 or qtd_tech_lote > 0:
                                formula_saldo = gerar_formula_saldo(item_s["codigo"], item_s["nf_origem"])
                                novas_linhas_est_s.append([c_seg, item_s["descricao"], formula_saldo, item_s["nf_origem"], 0, data_hora_segura])

                        deletar_linhas_em_lote(sheet_est, linhas_est_deletar)

                        if novas_linhas_mov_s:
                            sheet_mov.insert_rows(novas_linhas_mov_s, row=2, value_input_option='USER_ENTERED')
                        if novas_linhas_est_s:
                            sheet_est.insert_rows(novas_linhas_est_s, row=2, value_input_option='USER_ENTERED')

                        st.cache_data.clear()
                        st.success(f"✅ Saída concluída! Itens totalmente zerados (físico e manutenção) foram removidos do estoque automaticamente.")
                        
                        st.session_state["form_version"] += 1
                        time.sleep(1.2)
                        st.rerun()

elif aba == "📊 Consultar Estoque":
    
    col_pesq, col_btn = st.columns([5, 1])
    with col_pesq:
        busca = st.text_input("🔍 Pesquisar por peça, código ou NF:", placeholder="Digite para filtrar...")
    with col_btn:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Atualizar", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    dados_estoque = ler_dados_planilha("Estoque_Atual")
    dados_tec = ler_dados_planilha("Area_Tecnica")

    dict_na_tecnica = {}
    if dados_tec:
        for rt in dados_tec:
            ct = str(rt.get('Codigo_Peca', '')).replace("'", "").strip()
            nft = str(rt.get('Nota_Fiscal', '')).strip().lstrip('0')
            qt = pd.to_numeric(rt.get('Quantidade', 0), errors='coerce') or 0
            chave = f"{ct}_{nft}"
            dict_na_tecnica[chave] = dict_na_tecnica.get(chave, 0) + qt

    if dados_estoque:
        df = pd.DataFrame(dados_estoque)
        df['Quantidade_Atual'] = pd.to_numeric(df['Quantidade_Atual'], errors='coerce').fillna(0)
        
        if 'Estoque_Minimo' in df.columns:
            df['Estoque_Minimo_Val'] = pd.to_numeric(df['Estoque_Minimo'], errors='coerce')
        else:
            df['Estoque_Minimo_Val'] = None

        def calc_qtd_tecnica(row):
            cod_str = str(row['Codigo_Peca']).replace("'", "").strip()
            nf_str = str(row['Nota_Fiscal']).strip().lstrip('0')
            return int(dict_na_tecnica.get(f"{cod_str}_{nf_str}", 0))

        df['Qtd_Na_Tecnica'] = df.apply(calc_qtd_tecnica, axis=1)
        
        # Filtro inteligente: Mostrar se tem saldo positivo OU se está na área técnica
        df = df[(df['Quantidade_Atual'] > 0) | (df['Qtd_Na_Tecnica'] > 0)].copy()

        if not df.empty:
            contagem_codigos = df['Codigo_Peca'].value_counts()
            codigos_com_multiplas_nfs = contagem_codigos[contagem_codigos > 1].index.tolist()

            def formatar_codigo_e_icone(row):
                cod_str = str(row['Codigo_Peca']).replace("'", "").strip()
                nf_str = str(row['Nota_Fiscal']).strip().lstrip('0')
                chave = f"{cod_str}_{nf_str}"
                
                tem_na_tec = dict_na_tecnica.get(chave, 0) > 0
                tem_mult_nfs = cod_str in codigos_com_multiplas_nfs

                prefixo = ""
                if tem_na_tec:
                    prefixo += "🛠️ "
                if tem_mult_nfs:
                    prefixo += "* "

                return f"{prefixo}{cod_str}"

            df['Codigo_Formatado'] = df.apply(formatar_codigo_e_icone, axis=1)

            if busca:
                busca_limpa = extrair_numeros(busca)
                def atende_busca(row):
                    texto_linha = " ".join([str(val) for val in row])
                    if busca.lower() in texto_linha.lower():
                        return True
                    if busca_limpa and len(busca_limpa) >= 4:
                        if busca_limpa in extrair_numeros(texto_linha):
                            return True
                    return False

                df = df[df.apply(atende_busca, axis=1)]

            df['Estoque Mínimo'] = df['Estoque_Minimo_Val'].apply(
                lambda x: "" if pd.isna(x) or x == 0 else f"{x:.1f}".rstrip('0').rstrip('.')
            )

            colunas_exibicao = ['Codigo_Formatado', 'Descricao', 'Quantidade_Atual', 'Qtd_Na_Tecnica', 'Estoque Mínimo', 'Nota_Fiscal', 'Ultima_Atualizacao']
            colunas_exibicao = [c for c in colunas_exibicao if c in df.columns]

            def destacar_estoque_minimo(row):
                estilo = [''] * len(row)
                val_min = row.get('Estoque_Minimo_Val')
                if pd.notna(val_min) and val_min > 0 and row['Quantidade_Atual'] <= val_min:
                    if 'Quantidade_Atual' in row.index:
                        idx = row.index.get_loc('Quantidade_Atual')
                        estilo[idx] = 'background-color: rgba(255, 75, 75, 0.15); color: #ff8c8c;'
                return estilo

            df_estilizado = df[colunas_exibicao].style.apply(destacar_estoque_minimo, axis=1)

            st.dataframe(
                df_estilizado,
                height=400,
                use_container_width=True,
                column_config={
                    "Codigo_Formatado": st.column_config.TextColumn("Código da Peça (🛠️=Na Técnica | *=Várias NFs)"),
                    "Descricao": st.column_config.TextColumn("Descrição", width="large"),
                    "Quantidade_Atual": st.column_config.NumberColumn("Qtd\nAtual", width="small"),
                    "Qtd_Na_Tecnica": st.column_config.NumberColumn("Na\nTécnica", width="small"),
                    "Estoque Mínimo": st.column_config.TextColumn("Estoque Mínimo", width="small"),
                    "Nota_Fiscal": st.column_config.TextColumn("Nota Fiscal")
                },
                hide_index=True
            )

            with st.expander("🛠️ Ferramentas de Manutenção (Limpeza & Correção)", expanded=False):
                st.info("💡 **Aviso:** Utilizado para varrer a planilha limpando resíduos zerados manuais e sincronizando as fórmulas.")
                
                if st.button("🔄 Limpar Zerados e Recalcular Fórmulas", type="secondary", use_container_width=True):
                    with st.spinner("Excluindo itens zerados e atualizando fórmulas em lote..."):
                        sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
                        
                        dados_cru = sheet_est.get_all_records()
                        linhas_a_deletar = []
                        
                        dados_tec_limp = ler_dados_planilha("Area_Tecnica")
                        dict_tec_limp = {}
                        if dados_tec_limp:
                            for rt in dados_tec_limp:
                                ct = str(rt.get('Codigo_Peca', '')).replace("'", "").strip()
                                nft = str(rt.get('Nota_Fiscal', '')).strip().lstrip('0')
                                qt = pd.to_numeric(rt.get('Quantidade', 0), errors='coerce') or 0
                                dict_tec_limp[f"{ct}_{nft}"] = dict_tec_limp.get(f"{ct}_{nft}", 0) + qt

                        for idx_cru, row_cru in enumerate(dados_cru):
                            qtd_atual = pd.to_numeric(row_cru.get('Quantidade_Atual', 0), errors='coerce') or 0
                            cod_cru = str(row_cru.get('Codigo_Peca', '')).replace("'", "").strip()
                            nf_cru = str(row_cru.get('Nota_Fiscal', '')).strip().lstrip('0')
                            
                            tem_na_tec = dict_tec_limp.get(f"{cod_cru}_{nf_cru}", 0) > 0
                            
                            # Só deleta na limpeza geral se o item zerou NO FÍSICO e NÃO ESTÁ na área técnica
                            if qtd_atual <= 0 and not tem_na_tec:
                                linhas_a_deletar.append(idx_cru + 2)
                        
                        if linhas_a_deletar:
                            deletar_linhas_em_lote(sheet_est, linhas_a_deletar)
                        
                        dados_atualizados = sheet_est.get_all_records()
                        updates = []
                        for idx_cru, row_cru in enumerate(dados_atualizados):
                            cod_cru = str(row_cru.get('Codigo_Peca', '')).replace("'", "").strip()
                            nf_cru = str(row_cru.get('Nota_Fiscal', '')).strip().lstrip('0')
                            
                            if cod_cru and nf_cru:
                                nova_form = gerar_formula_saldo(cod_cru, nf_cru)
                                updates.append({
                                    'range': f'C{idx_cru + 2}',
                                    'values': [[nova_form]]
                                })
                        
                        if updates:
                            sheet_est.batch_update(updates, value_input_option='USER_ENTERED')
                                
                        st.success(f"✅ Limpeza concluída! {len(linhas_a_deletar)} lotes puramente zerados foram removidos.")
                        st.cache_data.clear()
                        time.sleep(2)
                        st.rerun()

            df_exportar = df[colunas_exibicao].copy()
            buffer_excel = io.BytesIO()
            
            with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
                df_exportar.to_excel(writer, sheet_name='Estoque_Atual', index=False)

            st.download_button(
                label="📊 Baixar Estoque Formatado em Excel (.xlsx)",
                data=buffer_excel.getvalue(),
                file_name=f"Estoque_GravoMark_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        else:
            st.info("Nenhum item em estoque e nenhuma peça em manutenção no momento.")
    else:
        st.info("Estoque vazio.")

elif aba == "📋 Histórico":
    
    st.markdown("<h3 style='margin-bottom:0px;'>📋 Histórico Geral & Auditoria</h3>", unsafe_allow_html=True)
    
    dados_mov = ler_dados_planilha("Movimentacoes", usar_formula=True)

    if dados_mov:
        df_mov_raw = pd.DataFrame(dados_mov)
        chaves = list(df_mov_raw.columns)
        
        # --- BLINDAGEM CONTRA PLANILHA SEM CABEÇALHO ---
        # Se a primeira coluna tiver barra (ex: data 28/09/2026) ou a fórmula HYPERLINK, a planilha perdeu o cabeçalho.
        if len(chaves) > 0:
            primeira_col = str(chaves[0])
            tem_numeros = any(char.isdigit() for char in primeira_col)
            tem_hyperlink = "HYPERLINK" in primeira_col.upper()
            
            if (tem_numeros and "/" in primeira_col) or tem_hyperlink:
                # O cabeçalho foi deletado pelo usuário e a primeira linha de dados assumiu o lugar.
                # Recriamos a linha que sumiu:
                nova_linha = pd.DataFrame([chaves], columns=df_mov_raw.columns)
                df_mov_raw = pd.concat([nova_linha, df_mov_raw], ignore_index=True)
                
        # Forçamos cabeçalhos reais para a matriz de visualização e links (nunca mais quebra o link)
        cols_padrao = ["Data_Hora", "Tipo_Movimentacao", "Nota_Fiscal", "Codigo_Peca", "Descricao", "Quantidade", "NF_Entrada_File", "NF_Saida_File", "Area_Tecnica"]
        
        if len(df_mov_raw.columns) <= len(cols_padrao):
            df_mov_raw.columns = cols_padrao[:len(df_mov_raw.columns)]
        else:
            novas_cols = cols_padrao.copy()
            for i in range(len(cols_padrao), len(df_mov_raw.columns)):
                novas_cols.append(f"Extra_{i}")
            df_mov_raw.columns = novas_cols

        c_data = "Data_Hora"
        c_tipo = "Tipo_Movimentacao"
        c_nf = "Nota_Fiscal"
        c_cod = "Codigo_Peca"
        c_desc = "Descricao"
        c_qtd = "Quantidade"
        c_link_entrada = "NF_Entrada_File" if "NF_Entrada_File" in df_mov_raw.columns else None
        c_link_saida = "NF_Saida_File" if "NF_Saida_File" in df_mov_raw.columns else None
        c_area_tec = "Area_Tecnica" if "Area_Tecnica" in df_mov_raw.columns else None

        def safe_nf_str(val):
            if isinstance(val, float) and val.is_integer(): return str(int(val))
            return str(val).strip()

        def processar_link_dataframe(valor):
            valor = str(valor).strip()
            if not valor or valor == "None" or valor == "": return None
            if 'HYPERLINK' in valor.upper():
                match_url = re.search(r'"(http.*?)"', valor)
                match_txt = re.search(r'",\s*"(.*?)"\)', valor)
                if match_url:
                    url = match_url.group(1)
                    txt = match_txt.group(1) if match_txt else ""
                    nums = extrair_numeros(txt)
                    display = f"📄 NF {nums}" if nums else "Doc 📄"
                    sep = "&" if "?" in url else "?"
                    return f"{url}{sep}st_display={display}"
            if valor.startswith("http"):
                sep = "&" if "?" in valor else "?"
                return f"{valor}{sep}st_display=Doc 📄"
            val_clean = valor.replace(" ", "-")
            return f"https://drive.google.com/drive/search?q={val_clean}&st_display={val_clean}"

        def prepare_for_display(df):
            df_disp = df.copy()
            if c_link_entrada and c_link_entrada in df_disp.columns:
                df_disp[c_link_entrada] = df_disp[c_link_entrada].apply(processar_link_dataframe)
            if c_link_saida and c_link_saida in df_disp.columns:
                df_disp[c_link_saida] = df_disp[c_link_saida].apply(processar_link_dataframe)
            if c_area_tec and c_area_tec in df_disp.columns:
                df_disp[c_area_tec] = df_disp[c_area_tec].apply(lambda x: "" if "Origem NF:" in str(x) else str(x))
            
            cols_to_drop = ['NF_Clean', 'Cod_Clean', 'Desc_Clean']
            df_disp = df_disp.drop(columns=[c for c in cols_to_drop if c in df_disp.columns])
            return df_disp

        df_mov_raw['NF_Clean'] = df_mov_raw[c_nf].apply(safe_nf_str)
        df_mov_raw['Cod_Clean'] = df_mov_raw[c_cod].astype(str).str.strip()
        df_mov_raw['Desc_Clean'] = df_mov_raw[c_desc].astype(str).str.strip()

        with st.expander("🔍 Rastreio e Auditoria de Lotes / NF", expanded=False):
            st.caption("Filtre o extrato para auditar um Lote ou visualizar todas as peças de uma NF específica.")
            
            modo_busca = st.radio("Selecione a ordem do filtro:", 
                                  ["📦 1. Buscar a PEÇA ➔ 2. Filtrar a NF", 
                                   "📄 1. Buscar a NF ➔ 2. Filtrar a PEÇA"], horizontal=True)
            
            cod_alvo = None
            nf_alvo = None
            tipo_alvo = "Entrada"
            entradas_df = df_mov_raw[df_mov_raw[c_tipo] == 'Entrada']

            if modo_busca == "📦 1. Buscar a PEÇA ➔ 2. Filtrar a NF":
                lista_todas_p = sorted(list(set([f"{r['Cod_Clean']} | {r['Desc_Clean']}" for _, r in entradas_df.iterrows() if r['Cod_Clean']])))
                peca_sel = st.selectbox("1️⃣ Selecione o Código ou Nome da Peça:", [""] + lista_todas_p)
                
                if peca_sel:
                    cod_tmp = peca_sel.split(" | ")[0].strip()
                    nfs_disp = entradas_df[entradas_df['Cod_Clean'] == cod_tmp]
                    lista_n = sorted(nfs_disp[nfs_disp['NF_Clean'] != ""]['NF_Clean'].unique().tolist())
                    
                    nf_sel = st.selectbox("2️⃣ Selecione a NF de Entrada (Filtrada apenas desta peça selecionada):", [""] + lista_n)
                    if nf_sel:
                        cod_alvo = cod_tmp
                        nf_alvo = nf_sel

            else:
                nfs_com_tipo = []
                for _, r in df_mov_raw.iterrows():
                    nf_val = safe_nf_str(r.get(c_nf, ''))
                    t_val = str(r.get(c_tipo, '')).strip()
                    if nf_val and nf_val != "None" and nf_val != "":
                        nfs_com_tipo.append(f"{nf_val} - {t_val}")
                nfs_com_tipo = sorted(list(set(nfs_com_tipo)))
                
                nf_sel_full = st.selectbox("1️⃣ Selecione a NF e a Operação:", [""] + nfs_com_tipo)
                
                if nf_sel_full:
                    partes_nf = nf_sel_full.rsplit(" - ", 1)
                    nf_alvo_temp = partes_nf[0].strip()
                    tipo_alvo_temp = partes_nf[1].strip()
                    
                    pecas_disp = df_mov_raw[(df_mov_raw['NF_Clean'] == nf_alvo_temp) & (df_mov_raw[c_tipo] == tipo_alvo_temp)]
                    lista_p = sorted(list(set([f"{r['Cod_Clean']} | {r['Desc_Clean']}" for _, r in pecas_disp.iterrows()])))
                    
                    peca_sel = st.selectbox(f"2️⃣ Selecione a Peça (Que consta na NF {nf_alvo_temp}):", [""] + lista_p)
                    if peca_sel:
                        cod_alvo = peca_sel.split(" | ")[0].strip()
                        nf_alvo = nf_alvo_temp
                        tipo_alvo = tipo_alvo_temp

            if cod_alvo and nf_alvo:
                if tipo_alvo == "Entrada":
                    def pertence_ao_lote(row):
                        t = str(row.get(c_tipo, ''))
                        c = str(row.get(c_cod, '')).strip()
                        if c != cod_alvo: return False
                        
                        nf_raw_c = row.get(c_nf, '')
                        nf_c = safe_nf_str(nf_raw_c)
                        
                        nf_g_raw = str(row.get(c_link_entrada, '')) if c_link_entrada else ""
                        nf_h_raw = str(row.get(c_link_saida, '')) if c_link_saida else ""
                            
                        if t in ['Entrada', 'Retorno Técnica', 'Área Técnica']:
                            return nf_alvo == nf_c
                        elif t == 'Saida':
                            return (nf_alvo in nf_h_raw) or (nf_alvo in nf_g_raw) or (nf_alvo == nf_c)
                        return False
                        
                    df_lote = df_mov_raw[df_mov_raw.apply(pertence_ao_lote, axis=1)].copy()
                    
                    if not df_lote.empty:
                        st.markdown("---")
                        tot_entrada = pd.to_numeric(df_lote[df_lote[c_tipo] == 'Entrada'][c_qtd], errors='coerce').sum()
                        tot_retorno = pd.to_numeric(df_lote[df_lote[c_tipo] == 'Retorno Técnica'][c_qtd], errors='coerce').sum()
                        tot_saida = pd.to_numeric(df_lote[df_lote[c_tipo] == 'Saida'][c_qtd], errors='coerce').sum()
                        tot_tec = pd.to_numeric(df_lote[df_lote[c_tipo] == 'Área Técnica'][c_qtd], errors='coerce').sum()
                        
                        saldo_esperado = (tot_entrada + tot_retorno) - (tot_saida + tot_tec)
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric(f"📦 Entrou no Lote (NF {nf_alvo})", f"{int(tot_entrada)} un")
                        c2.metric("📤 Vendido/Saída (Deste Lote)", f"{int(tot_saida)} un")
                        c3.metric("🛠️ Em Manutenção (Deste Lote)", f"{int(tot_tec - tot_retorno)} un")
                        c4.metric("✅ Saldo Físico Correto (Deste Lote)", f"{int(saldo_esperado)} un", delta=int(saldo_esperado), delta_color="normal" if saldo_esperado>=0 else "inverse")
                        
                        st.dataframe(
                            prepare_for_display(df_lote),
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                c_data: st.column_config.TextColumn("Data e Hora"),
                                c_tipo: st.column_config.TextColumn("Operação / Status"),
                                c_nf: st.column_config.TextColumn("NF / Doc"),
                                c_cod: st.column_config.TextColumn("Código da Peça"),
                                c_desc: st.column_config.TextColumn("Descrição", width="large"),
                                c_qtd: st.column_config.NumberColumn("Qtd", width="small"),
                                c_link_entrada: st.column_config.LinkColumn("Origem/NFe", display_text=r"st_display=(.*)", width="small"),
                                c_link_saida: st.column_config.LinkColumn("Doc/NFs", display_text=r"st_display=(.*)", width="small"),
                                c_area_tec: st.column_config.TextColumn("Área Técnica / Detalhes", width="medium")
                            }
                        )
                else:
                    df_op = df_mov_raw[(df_mov_raw['NF_Clean'] == nf_alvo) & (df_mov_raw[c_tipo] == tipo_alvo) & (df_mov_raw['Cod_Clean'] == cod_alvo)].copy()
                    st.markdown(f"**Visualizando operação de {tipo_alvo} da NF {nf_alvo}:**")
                    st.dataframe(
                        prepare_for_display(df_op),
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            c_data: st.column_config.TextColumn("Data e Hora"),
                            c_tipo: st.column_config.TextColumn("Operação / Status"),
                            c_nf: st.column_config.TextColumn("NF / Doc"),
                            c_cod: st.column_config.TextColumn("Código da Peça"),
                            c_desc: st.column_config.TextColumn("Descrição", width="large"),
                            c_qtd: st.column_config.NumberColumn("Qtd", width="small"),
                            c_link_entrada: st.column_config.LinkColumn("Origem/NFe", display_text=r"st_display=(.*)", width="small"),
                            c_link_saida: st.column_config.LinkColumn("Doc/NFs", display_text=r"st_display=(.*)", width="small"),
                            c_area_tec: st.column_config.TextColumn("Área Técnica / Detalhes", width="medium")
                        }
                    )
        
        st.divider()

        # --- TABELA GERAL DO HISTÓRICO ---
        col_hist_busca_2, col_hist_btn_2 = st.columns([5, 1])
        with col_hist_busca_2:
            busca_hist = st.text_input("🔍 Pesquisar no Histórico Geral (por código, NF, descrição ou operação):", placeholder="Digite para filtrar...")
        with col_hist_btn_2:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            if st.button("🔄 Atualizar Histórico", use_container_width=True, key="btn_hist_geral"):
                st.cache_data.clear()
                st.rerun()

        if busca_hist:
            busca_limpa_h = extrair_numeros(busca_hist)
            def atende_busca_h(row):
                texto_linha = " ".join([str(val) for val in row])
                if busca_hist.lower() in texto_linha.lower():
                    return True
                if busca_limpa_h and len(busca_limpa_h) >= 4:
                    if busca_limpa_h in extrair_numeros(texto_linha):
                        return True
                return False
            df_mov_geral = df_mov_raw[df_mov_raw.apply(atende_busca_h, axis=1)]
        else:
            df_mov_geral = df_mov_raw

        st.dataframe(
            prepare_for_display(df_mov_geral.head(200)),
            height=400,
            use_container_width=True,
            column_config={
                c_data: st.column_config.TextColumn("Data e Hora"),
                c_tipo: st.column_config.TextColumn("Operação / Status"),
                c_nf: st.column_config.TextColumn("NF / Doc"),
                c_cod: st.column_config.TextColumn("Código da Peça"),
                c_desc: st.column_config.TextColumn("Descrição", width="large"),
                c_qtd: st.column_config.NumberColumn("Qtd", width="small"),
                c_link_entrada: st.column_config.LinkColumn("Origem/NFe", display_text=r"st_display=(.*)", width="small"),
                c_link_saida: st.column_config.LinkColumn("Doc/NFs", display_text=r"st_display=(.*)", width="small"),
                c_area_tec: st.column_config.TextColumn("Área Técnica / Detalhes", width="medium")
            },
            hide_index=True
        )

        st.caption("✏️️ **Anexar / Editar NF ou Link em Registro Antigo:**")

        opcoes_linhas = []
        mapa_linhas = {}

        for idx_m, rm in enumerate(dados_mov[:150]):
            num_linha = idx_m + 2
            nf_m = str(rm.get(c_nf, "")).strip()
            cod_m = str(rm.get(c_cod, "")).replace("'", "").strip()
            desc_m = str(rm.get(c_desc, "")).strip()
            tp_m = str(rm.get(c_tipo, "")).strip()

            desc_curta = (desc_m[:25] + '...') if len(desc_m) > 25 else desc_m
            rotulo = f"NF: {nf_m} | Peça: {cod_m} | {desc_curta} ({tp_m}) [L- {num_linha}]"
            
            opcoes_linhas.append(rotulo)
            mapa_linhas[rotulo] = num_linha

        if opcoes_linhas:
            col_sel_lin, col_up_nf, col_link_manual = st.columns([4, 3, 3])
            
            with col_sel_lin:
                linha_selecionada_rotulo = st.selectbox(
                    "Selecione o registro:", 
                    options=[""] + opcoes_linhas,
                    key=f"sel_reg_edit_{ver}"
                )
            
            with col_up_nf:
                arquivo_edit_nf = st.file_uploader("Upload do PDF pro Drive:", type=["pdf"], key=f"upload_edit_nf_{ver}")
            
            with col_link_manual:
                link_manual_input = st.text_input("OU cole o Link do Drive:", placeholder="https://drive.google.com/...", key=f"link_edit_nf_{ver}")

            if st.button("💾 Salvar NF no Registro Selecionado", type="primary", use_container_width=True):
                if not linha_selecionada_rotulo:
                    st.error("Selecione um registro na lista acima.")
                elif not arquivo_edit_nf and not link_manual_input.strip():
                    st.error("Faça o upload do PDF ou cole o link manual da NF.")
                else:
                    with st.spinner("Atualizando registro da NF..."):
                        num_linha = mapa_linhas[linha_selecionada_rotulo]
                        sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")

                        link_final = ""
                        if arquivo_edit_nf is not None:
                            pasta_id = FOLDER_ENTRADA_ID
                            file_metadata = {'name': arquivo_edit_nf.name, 'parents': [pasta_id]}
                            media = MediaIoBaseUpload(io.BytesIO(arquivo_edit_nf.getvalue()), mimetype=arquivo_edit_nf.type)
                            arquivo_salvo = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                            url_b = f"https://drive.google.com/file/d/{arquivo_salvo.get('id')}/view"
                            link_final = f'=HYPERLINK("{url_b}"; "📄 Abrir NF")'
                        elif link_manual_input.strip():
                            url_b = link_manual_input.strip()
                            link_final = f'=HYPERLINK("{url_b}"; "📄 Abrir NF")'

                        sheet_mov.update_cell(num_linha, 7, link_final)
                        st.cache_data.clear()
                        st.success("✅ NF/Link atualizado com sucesso no registro!")
                        
                        st.session_state["form_version"] += 1
                        time.sleep(1.2)
                        st.rerun()

    else:
        st.info("Nenhuma movimentação registrada.")

elif aba == "🛠️ Área Técnica":

    dados_estoque_atual = ler_dados_planilha("Estoque_Atual")
    dados_tec = ler_dados_planilha("Area_Tecnica")

    aba_tec_sub = st.tabs(["📤 Enviar Lote para Área Técnica", "📋 Peças na Área Técnica & Devolução por Checkbox"])

    with aba_tec_sub[0]:
        col_at1, col_at2, col_at3 = st.columns([2, 2, 2])
        with col_at1:
            pedido_os = st.text_input("Pedido / OS / Atendimento*", placeholder="Ex: OS-1052", key=f"input_os_{ver}")
        with col_at2:
            nome_tecnico = st.text_input("Técnico Responsável*", placeholder="Ex: Robinho", key=f"input_tec_{ver}")
        with col_at3:
            nome_cliente = st.text_input("Nome do Cliente (Opcional)", key=f"input_cli_{ver}")

        obs_tec = st.text_input("Observações Gerais", key=f"input_obs_{ver}")

        opcoes_tec_disp = []
        for r in dados_estoque_atual:
            c_item = str(r.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
            q_item = pd.to_numeric(r.get('Quantidade_Atual', 0), errors='coerce') or 0
            raw_nf = str(r.get('Nota_Fiscal', '')).strip()
            desc_item = str(r.get('Descricao', '')).strip()
            
            if c_item and q_item > 0:
                c_sem_pontos = extrair_numeros(c_item)
                if c_sem_pontos and c_sem_pontos != c_item:
                    rotulo_peca = f"{c_item} ({c_sem_pontos})"
                else:
                    rotulo_peca = c_item
                    
                opcoes_tec_disp.append(f"{rotulo_peca} | NF {raw_nf} | Saldo: {int(q_item)} un | {desc_item}")

        if not opcoes_tec_disp:
            st.warning("⚠️ Não há peças no estoque para transferir à Área Técnica.")
        else:
            df_tec_template = pd.DataFrame([
                {"Selecione Peça e NF no Estoque": "", "Quantidade": 1}
                for _ in range(5)
            ])

            editor_tec = st.data_editor(
                df_tec_template,
                num_rows="dynamic",
                use_container_width=True,
                height=220,
                key=f"editor_tec_{ver}",
                column_config={
                    "Selecione Peça e NF no Estoque": st.column_config.SelectboxColumn(
                        "Selecione Peça e NF de Origem",
                        options=[""] + sorted(opcoes_tec_disp),
                        width="large"
                    ),
                    "Quantidade": st.column_config.NumberColumn("Qtd Enviar", min_value=1, step=1, default=1, width="small")
                }
            )

            if st.button("🚀 Transferir Lote para Área Técnica", use_container_width=True, type="primary"):
                if not pedido_os.strip() or not nome_tecnico.strip():
                    st.error("Por favor, preencha o número da OS/Pedido e o Técnico Responsável.")
                else:
                    itens_transferir = []
                    for idx, row in editor_tec.iterrows():
                        opcao_t = str(row.get("Selecione Peça e NF no Estoque", "") or "").strip()
                        qtd_t = int(row.get("Quantidade", 1) or 1)

                        if opcao_t and opcao_t != "None":
                            partes_t = opcao_t.split(" | ")
                            cod_t = partes_t[0].split(" (")[0].strip()
                            raw_nf_t = partes_t[1].replace("NF ", "").strip()
                            
                            saldo_str = partes_t[2].replace("Saldo:", "").replace("un", "").strip()
                            saldo_atual = int(saldo_str)
                            
                            desc_t = partes_t[3].strip() if len(partes_t) > 3 else ""
                            itens_transferir.append({
                                "codigo": cod_t, 
                                "nf_origem": raw_nf_t, 
                                "descricao": desc_t, 
                                "qtd": qtd_t, 
                                "saldo_atual": saldo_atual
                            })

                    if not itens_transferir:
                        st.error("Selecione pelo menos uma peça na tabela acima.")
                    else:
                        with st.spinner("Transferindo para Área Técnica..."):
                            data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                            data_hora_segura = f"'{data_hora_bruta}"

                            sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                            sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
                            sheet_tec = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Area_Tecnica")

                            novas_linhas_mov_t = []
                            novas_linhas_tec_t = []
                            linhas_est_a_deletar = []
                            novas_linhas_est_t = []

                            dados_est_cru = sheet_est.get_all_records()

                            for item_t in itens_transferir:
                                c_seg = f"'{item_t['codigo']}"
                                info_tec = f"OS: {pedido_os} | Técnico: {nome_tecnico}"
                                novas_linhas_mov_t.append([data_hora_segura, "Área Técnica", item_t["nf_origem"], c_seg, item_t["descricao"], item_t["qtd"], "", "", info_tec])

                                for idx_e, row_e in enumerate(dados_est_cru):
                                    cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
                                    nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                                    if cod_e == item_t['codigo'] and nf_e == item_t['nf_origem'].lstrip('0'):
                                        linhas_est_a_deletar.append(idx_e + 2)

                                # SEMPRE re-insere na aba de estoque atual, permitindo que a quantidade fique 0 
                                # e mantendo a visibilidade visual (com o ícone da técnica)
                                formula_saldo = gerar_formula_saldo(item_t["codigo"], item_t["nf_origem"])
                                novas_linhas_est_t.append([c_seg, item_t["descricao"], formula_saldo, item_t["nf_origem"], 0, data_hora_segura])

                                novas_linhas_tec_t.append([c_seg, item_t["descricao"], item_t["qtd"], item_t["nf_origem"], pedido_os, nome_cliente, nome_tecnico, "Diagnóstico / Manutenção", data_hora_segura, obs_tec])

                            deletar_linhas_em_lote(sheet_est, linhas_est_a_deletar)

                            if novas_linhas_mov_t:
                                sheet_mov.insert_rows(novas_linhas_mov_t, row=2, value_input_option='USER_ENTERED')
                            if novas_linhas_est_t:
                                sheet_est.insert_rows(novas_linhas_est_t, row=2, value_input_option='USER_ENTERED')
                            if novas_linhas_tec_t:
                                sheet_tec.insert_rows(novas_linhas_tec_t, row=2, value_input_option='USER_ENTERED')

                            st.cache_data.clear()
                            st.success(f"🎉 Transferência concluída! O item ficará visível na aba Estoque com o ícone de manutenção.")
                            
                            st.session_state["form_version"] += 1
                            time.sleep(1.2)
                            st.rerun()

    with aba_tec_sub[1]:
        st.caption("☑️ Marque as caixinhas na tabela dos itens que deseja devolver e clique no botão abaixo:")

        if dados_tec:
            df_tec = pd.DataFrame(dados_tec)
            if not df_tec.empty:
                cols_texto = ["Codigo_Peca", "Descricao", "Nota_Fiscal", "Pedido_OS", "Cliente", "Tecnico", "Status", "Observacao"]
                for col in cols_texto:
                    if col in df_tec.columns:
                        df_tec[col] = df_tec[col].astype(str)

                df_tec.insert(0, "Devolver?", False)
                df_tec["_Row_Idx"] = df_tec.index + 2

                df_tec_edit = st.data_editor(
                    df_tec,
                    height=450,
                    hide_index=True,
                    use_container_width=True,
                    key=f"editor_dev_tec_{ver}",
                    column_config={
                        "Devolver?": st.column_config.CheckboxColumn("Devolver?", default=False, width="small"),
                        "Codigo_Peca": st.column_config.TextColumn("Código Peça"),
                        "Descricao": st.column_config.TextColumn("Descrição", width="large"),
                        "Quantidade": st.column_config.NumberColumn("Qtd", width="small"),
                        "Nota_Fiscal": st.column_config.TextColumn("NF Origem"),
                        "Pedido_OS": st.column_config.TextColumn("OS / Pedido"),
                        "Tecnico": st.column_config.TextColumn("Técnico"),
                        "_Row_Idx": None
                    }
                )

                if st.button("↩️ Devolver Peças Selecionadas ao Estoque", type="primary", use_container_width=True):
                    itens_para_devolver = df_tec_edit[df_tec_edit["Devolver?"] == True]

                    if itens_para_devolver.empty:
                        st.warning("Selecione pelo menos uma peça marcando a caixinha 'Devolver?'.")
                    else:
                        with st.spinner("Processando retorno das peças (Alta Performance)..."):
                            data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                            data_hora_segura = f"'{data_hora_bruta}"

                            sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                            sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
                            sheet_tec = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Area_Tecnica")

                            linhas_sheet_deletar = sorted(itens_para_devolver["_Row_Idx"].tolist(), reverse=True)

                            deletar_linhas_em_lote(sheet_tec, linhas_sheet_deletar)

                            novas_mov_dev = []
                            novas_est_dev = []
                            linhas_est_substituir = []

                            dados_est_cru = sheet_est.get_all_records()

                            for _, row_d in itens_para_devolver.iterrows():
                                c_d = str(row_d.get('Codigo_Peca', '')).replace("'", "").strip()
                                nf_d = str(row_d.get('Nota_Fiscal', '')).strip()
                                q_d = int(row_d.get('Quantidade', 1))
                                desc_d = str(row_d.get('Descricao', '')).strip()

                                c_seg = f"'{c_d}"
                                novas_mov_dev.append([data_hora_segura, "Retorno Técnica", nf_d, c_seg, desc_d, q_d, "", "", "Retornou da bancada/campo"])

                                for idx_e, row_e in enumerate(dados_est_cru):
                                    cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
                                    nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                                    if cod_e == c_d and nf_e == nf_d.lstrip('0'):
                                        linhas_est_substituir.append(idx_e + 2)

                                formula_saldo = gerar_formula_saldo(c_d, nf_d)
                                novas_est_dev.append([c_seg, desc_d, formula_saldo, nf_d, 0, data_hora_segura])

                            deletar_linhas_em_lote(sheet_est, linhas_est_substituir)

                            if novas_mov_dev:
                                sheet_mov.insert_rows(novas_mov_dev, row=2, value_input_option='USER_ENTERED')
                            if novas_est_dev:
                                sheet_est.insert_rows(novas_est_dev, row=2, value_input_option='USER_ENTERED')

                            st.cache_data.clear()
                            st.success(f"🎉 {len(itens_para_devolver)} item(ns) retornado(s) ao Estoque.")
                            
                            st.session_state["form_version"] += 1
                            time.sleep(1.2)
                            st.rerun()

            else:
                st.info("Nenhuma peça atualmente na área técnica.")
        else:
            st.info("Nenhum registro de área técnica.")
