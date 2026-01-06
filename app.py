import streamlit as st
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import time
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURAÇÕES INICIAIS ---
NOME_DA_PLANILHA = "RNCs - Qualidade Industrial"
NOME_ARQUIVO_MODELO = "Modelo - Registro de Não Conformidade.docx"

# ⚠️ CERTIFIQUE-SE DE QUE O ID ESTÁ CORRETO
ID_PASTA_DRIVE = "0AGUUHFUnBEXtUk9PVA" 

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def conectar_google_auth():
    try:
        credentials = Credentials.from_service_account_info(
            st.secrets["gsheets"],
            scopes=SCOPES
        )
        return credentials
    except Exception as e:
        st.error(f"Erro de Autenticação: {e}")
        return None

def conectar_gsheets():
    creds = conectar_google_auth()
    if creds:
        return gspread.authorize(creds)
    return None

def salvar_dados_sheets(dados):
    client = conectar_gsheets()
    if not client: return False

    try:
        sh = client.open(NOME_DA_PLANILHA)
        sheet = sh.sheet1
        dados_str = [str(item) if item is not None else "" for item in dados]
        sheet.append_row(dados_str)
        return True
    except Exception as e:
        if "200" in str(e): return True
        st.error(f"❌ Erro ao salvar no Sheets: {e}")
        return False

def upload_para_drive(buffer_arquivo, nome_arquivo):
    """Envia o arquivo em memória para o Google Drive (Compatível com Shared Drives)"""
    creds = conectar_google_auth()
    if not creds: return None

    try:
        buffer_arquivo.seek(0)
        
        # Constrói o serviço
        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': nome_arquivo,
            'parents': [ID_PASTA_DRIVE] # Certifique-se que este ID é de uma pasta em um DRIVE COMPARTILHADO
        }

        media = MediaIoBaseUpload(
            buffer_arquivo, 
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            resumable=True
        )

        # --- AQUI ESTÁ O TRUQUE PARA DRIVE COMPARTILHADO ---
        # Adicionamos supportsAllDrives=True para permitir salvar em Drives de Equipe
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink',
            supportsAllDrives=True 
        ).execute()

        return file.get('webViewLink')

    except Exception as e:
        # Tratamento de erro específico para cota
        erro_msg = str(e)
        if "storageQuotaExceeded" in erro_msg:
            st.error("❌ ERRO DE COTA: O Robô não tem espaço. Use uma pasta dentro de um 'Drive Compartilhado' (Shared Drive) e não no 'Meu Drive'.")
        else:
            st.error(f"❌ Erro no Google Drive: {erro_msg}")
        return None

def gerar_laudo_docx(contexto, imagem_bytes=None):
    try:
        doc = DocxTemplate(NOME_ARQUIVO_MODELO)
        
        # Lógica da Imagem
        if imagem_bytes:
            # width=Mm(100) define largura de 10cm
            imagem_inline = InlineImage(doc, io.BytesIO(imagem_bytes), width=Mm(100))
            contexto['foto'] = imagem_inline
        else:
            contexto['foto'] = ""

        doc.render(contexto)
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        st.error(f"Erro ao gerar documento Word: {e}")
        return None

def limpar_campos():
    campos_texto = [
        "emitente", "area_id", "nao_conf", "n_nc", "cc_origem", "setor_origem", 
        "causa", "desc_item", "cod_item", "fornecedor", "cor_tinta", 
        "cliente", "pedido", "op", "obs", "ass_lider", "ass_coord", 
        "ass_qual", "ass_refugo", "ass_gerente",
        "n_pecas_nc", "metragem_ger_nc", "peso_total_nc"
    ]
    for campo in campos_texto:
        if campo in st.session_state: st.session_state[campo] = ""
            
    campos_num = ["qtd_pecas", "metragem", "peso"]
    for campo in campos_num:
        if campo in st.session_state: st.session_state[campo] = 0
            
    if "data_nc" in st.session_state: st.session_state["data_nc"] = datetime.now()
    if "turno" in st.session_state: st.session_state["turno"] = "1º Turno"
    if "acao" in st.session_state: st.session_state["acao"] = "Retrabalhar"
    
    if "img_uploader_key" in st.session_state:
        st.session_state["img_uploader_key"] += 1

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(layout="centered", page_title="RNC - Formulário Digital Águia Sistemas", page_icon="📝")

st.markdown("""
    <style>
    .stApp { background-color: #f4f6f9; }
    [data-testid="stForm"] {
        background-color: #ffffff; padding: 2rem; border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08); border-top: 5px solid #2E3182; 
    }
    .section-header {
        background-color: #e9ecef; color: #333; padding: 8px 12px;
        font-weight: 700; text-transform: uppercase; font-size: 0.85rem;
        border-left: 4px solid #2E3182; margin-top: 20px; margin-bottom: 15px; letter-spacing: 0.5px;
    }
    .stTextInput input:disabled { background-color: #f8f9fa; color: #495057; font-weight: 600; border-color: #dee2e6; }
    label { font-size: 0.8rem !important; font-weight: 600 !important; color: #555 !important; }
    
    .download-area {
        background-color: #d4edda; padding: 15px; border-radius: 5px; 
        border: 1px solid #c3e6cb; margin-bottom: 20px; text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    if 'revisao_count' not in st.session_state: st.session_state.revisao_count = 1
    if 'sucesso_salvamento' not in st.session_state: st.session_state.sucesso_salvamento = False
    if 'buffer_download' not in st.session_state: st.session_state.buffer_download = None
    if 'nome_arquivo_download' not in st.session_state: st.session_state.nome_arquivo_download = ""
    if 'link_drive_gerado' not in st.session_state: st.session_state.link_drive_gerado = ""
    if "img_uploader_key" not in st.session_state: st.session_state.img_uploader_key = 0

    # --- ÁREA DE SUCESSO ---
    if st.session_state.sucesso_salvamento:
        limpar_campos() 
        st.markdown('<div class="download-area">', unsafe_allow_html=True)
        st.success("✅ Processo Concluído!")
        
        if st.session_state.link_drive_gerado:
            st.markdown(f"☁️ **[Abrir no Google Drive]({st.session_state.link_drive_gerado})**")
        else:
            st.warning("Upload para o Drive não retornou link (verifique logs).")

        if st.session_state.buffer_download:
            st.download_button(
                label="📥 BAIXAR DOCX (BACKUP)",
                data=st.session_state.buffer_download,
                file_name=st.session_state.nome_arquivo_download,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
        st.markdown('</div>', unsafe_allow_html=True)
        
        if st.button("Novo Registro"):
            st.session_state.sucesso_salvamento = False
            st.session_state.buffer_download = None
            st.rerun()

    num_revisao = f"{st.session_state.revisao_count:03d}"

    # --- BARRA LATERAL ---
    with st.sidebar:
        # CORREÇÃO 1: use_container_width -> width="stretch"
        try:
            st.image("image_1.png", width="stretch") 
        except:
            pass
            
        st.caption("Sistema Integrado")
        if conectar_gsheets(): st.success("BD Conectado")
        else: st.error("BD Desconectado")

    # --- FORMULÁRIO ---
    with st.form("rnc_completa", clear_on_submit=False):
        c1, c2 = st.columns([1, 3])

        with c1:
            st.image("image_1.png", width="stretch")

        with c2:
            st.markdown("###### DADOS DE CONTROLE")
            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                st.text_input("Revisão", value="0001", disabled=True)
                st.text_input("Setor", value="Qualidade Industrial", disabled=True)
            with cc2:
                st.text_input("Autor", value="Evelyn Silva", disabled=True)
                st.text_input("Doc", value="RNC - 20", disabled=True)
            with cc3:
                st.text_input("Autorizado por:", value="Melina Favaro", disabled=True)
                st.text_input("Última Atualização", value="30/09/2025", disabled=True)
        
        st.markdown("---")
        st.markdown("<h3 style='text-align: center; color: #2E3182;'>REGISTRO DE NÃO CONFORMIDADE</h3>", unsafe_allow_html=True)

        # SEÇÕES DO FORMULÁRIO
        st.markdown('<div class="section-header">INFORMAÇÕES GERAIS</div>', unsafe_allow_html=True)
        g1, g2, g3 = st.columns([1.5, 2, 1.5])
        with g1: data_nc = st.date_input("Data", value=datetime.now(), key="data_nc")
        with g2: emitente = st.text_input("Emitente", key="emitente")
        with g3: turno = st.selectbox("Turno", ["1º Turno", "2º Turno", "3º Turno", "Adm"], key="turno")
        area_id = st.text_input("Área de Identificação do Material:", key="area_id")

        st.markdown('<div class="section-header">QUALIDADE</div>', unsafe_allow_html=True)
        iqa, iqb, iqc = st.columns(3)
        with iqa: n_pecas_nc = st.text_input("Nº de Peças NC", key="n_pecas_nc")
        with iqb: metragem_ger_nc = st.text_input("Metragem geradora da NC:", key="metragem_ger_nc")
        with iqc: peso_total_nc = st.text_input("Peso Total NC (kg)", key="peso_total_nc")

        iq1, iq2 = st.columns([4, 1])
        with iq1: nao_conf = st.text_input("Não Conformidade", key="nao_conf")
        with iq2: n_nc = st.text_input("Nº NC", key="n_nc")
        
        iq3, iq4 = st.columns(2)
        with iq3: cc_origem = st.text_input("Centro Custo - Setor de Origem", key="cc_origem")
        with iq4: setor_origem = st.text_input("Setor Origem da NC", key="setor_origem")
        causa = st.text_area("Causa Raiz", height=68, key="causa")

        st.markdown('<div class="section-header">DADOS DA PEÇA</div>', unsafe_allow_html=True)
        ip1, ip2 = st.columns([3, 1])
        with ip1: desc_item = st.text_input("Descrição Item", key="desc_item")
        with ip2: cod_item = st.text_input("Item", key="cod_item")
        
        ip3, ip4, ip5 = st.columns(3)
        with ip3: qtd_pecas = st.number_input("Nº de Peças", min_value=0, key="qtd_pecas")
        with ip4: metragem = st.number_input("Metragem geradora da NC", min_value=0.0, format="%.2f", key="metragem")
        with ip5: peso = st.number_input("Peso total NC", min_value=0.0, format="%.2f", key="peso")

        st.markdown('<div class="section-header">MATÉRIA PRIMA / PROJETO</div>', unsafe_allow_html=True)
        it1, it2 = st.columns(2)
        with it1: fornecedor = st.text_input("Fornecedor da Tinta / Aço", key="fornecedor")
        with it2: cor_tinta = st.text_input("Cor da Tinta", key="cor_tinta")
        
        proj1, proj2, proj3 = st.columns([2, 1, 1])
        with proj1: cliente = st.text_input("Cliente", key="cliente")
        with proj2: pedido = st.text_input("Pedido", key="pedido")
        with proj3: op = st.text_input("Ordem de Produção", key="op")

        st.markdown('<div class="section-header">AÇÃO IMEDIATA</div>', unsafe_allow_html=True)
        acao = st.radio("Ação", ["Retrabalhar", "Liberar sob concessão", "Refugar", "Sucatear", "Reaproveitamento", "Alterar projeto"], horizontal=True, label_visibility="collapsed", key="acao")
        
        st.markdown('<div class="section-header">EVIDÊNCIAS FOTOGRÁFICAS</div>', unsafe_allow_html=True)
        
        tab_cam, tab_upl = st.tabs(["📸 Tirar Foto", "📂 Upload de Arquivo"])
        imagem_final = None 
        
        with tab_cam:
            foto_cam = st.camera_input("Capturar imagem da peça")
            if foto_cam: imagem_final = foto_cam.getvalue()
        
        with tab_upl:
            foto_upl = st.file_uploader("Escolher imagem", type=['png', 'jpg', 'jpeg'], key=f"uploader_{st.session_state.img_uploader_key}")
            if foto_upl: imagem_final = foto_upl.getvalue()

        st.markdown('<div class="section-header">OBSERVAÇÕES ADICIONAIS</div>', unsafe_allow_html=True)
        obs = st.text_area("Texto", height=100, label_visibility="collapsed", key="obs")

        st.markdown("---")
        st.markdown("#### ✅ Assinaturas")
        s1, s2, s3 = st.columns(3)
        with s1: ass_lider = st.text_input("Líder", key="ass_lider")
        with s2: ass_coord = st.text_input("Coordenação", key="ass_coord")
        with s3: ass_qual = st.text_input("Qualidade", key="ass_qual")
        
        r1, r2 = st.columns(2)
        with r1: ass_refugo = st.text_input("Solicitante Refugo", key="ass_refugo")
        with r2: ass_gerente = st.text_input("Gerência", key="ass_gerente")

        st.markdown("<br>", unsafe_allow_html=True)
        
        # CORREÇÃO 2: use_container_width -> width="stretch"
        submit_btn = st.form_submit_button("💾 REGISTRAR, GERAR LAUDO E UPLOAD", type="primary", width="stretch")
        
        if submit_btn:
            with st.spinner("⏳ Processando..."):
                contexto_docx = {
                    "data_nc": data_nc.strftime("%d/%m/%Y"),
                    "emitente": emitente, "turno": turno, "area_id": area_id,
                    "n_pecas_nc": n_pecas_nc, "metragem_ger_nc": metragem_ger_nc, "peso_total_nc": peso_total_nc,
                    "nao_conf": nao_conf, "n_nc": n_nc, "cc_origem": cc_origem, "setor_origem": setor_origem, "causa": causa,
                    "desc_item": desc_item, "cod_item": cod_item, "qtd_pecas": qtd_pecas, "metragem": metragem, "peso": peso,
                    "fornecedor": fornecedor, "cor_tinta": cor_tinta, "cliente": cliente, "pedido": pedido, "op": op,
                    "acao": acao, "obs": obs,
                    "ass_lider": ass_lider, "ass_coord": ass_coord, "ass_qual": ass_qual, "ass_refugo": ass_refugo, "ass_gerente": ass_gerente
                }
                
                buffer_doc = gerar_laudo_docx(contexto_docx, imagem_final)
                
                if buffer_doc:
                    nome_arquivo = f"RNC_{n_nc}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                    
                    link_drive = upload_para_drive(buffer_doc, nome_arquivo)
                    
                    linha_dados = [
                        data_nc, emitente, turno, area_id,
                        n_pecas_nc, metragem_ger_nc, peso_total_nc, nao_conf, n_nc, cc_origem, setor_origem, causa,
                        desc_item, cod_item, qtd_pecas, metragem, peso,
                        fornecedor, cor_tinta, cliente, pedido, op,
                        acao, obs,
                        ass_lider, ass_coord, ass_qual, ass_refugo, ass_gerente,
                        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        link_drive if link_drive else "Erro no Upload"
                    ]
                    
                    sucesso_sheets = salvar_dados_sheets(linha_dados)
                    
                    if sucesso_sheets:
                        st.session_state.revisao_count += 1
                        st.session_state.sucesso_salvamento = True
                        st.session_state.buffer_download = buffer_doc
                        st.session_state.nome_arquivo_download = nome_arquivo
                        st.session_state.link_drive_gerado = link_drive
                        st.rerun()

if __name__ == "__main__":
    main()