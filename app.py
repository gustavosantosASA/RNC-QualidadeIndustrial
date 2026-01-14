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
        with iq1: nao_conf = st.selectbox("Não Conformidade", [" ", "1 COTA CRÍTICA FORA DO ESPECIFICADO", "2 COTA GERAL FORA DO ESPECIFICADO", "3 PEÇA EXTRAVIADA PÓS REPORTE", "4 CONFERÊNCIA FALHA (QUANTIDADE)", "5 PEDIDOS MISTURADOS NA MESMA EMBALAGEM", "6 PROBLEMAS DE EMBALAGEM", "7 MATERIAL DANIFICADO DURANTE MOVIMENTAÇÃO", "8 ETIQUETA A CANETA", "9 ETIQUETA TROCADA", "10 MATERIAL SEM ETIQUETA", "11 BITOLA INCORRETA", "12 ESQUADRO FORA DO ESPECIFICADO", "13 REBARBA EXCESSIVA", "14 DESVIO DE FORMA NA CURVATURA", "15 ESTAMPO DUPLO", "16 ESTAMPO FORA DE PASSO / POSIÇÃO", "17 REBARBA EXCESSIVA NO ESTAMPO", "18 COMPRIMENTO FORA DO ESPECIFICADO", "19 DESVIO DE FORMA / REBARBA NO CORTE", "20 DESVIO DE FORMA DE TORÇÃO", "21 DESVIO NA POSIÇÃO DO ESTAMPO TRANSVERSAL", "22 DISTÂNCIA DE CORTE DO ESTAMPO DESCENTRALIZADO", "23 RAIO FORA DAS ESPECIFICAÇÕES", "24 RISCOS, MARCAS, ARRANCAMENTO NA SUPERFÍCIE", "25 PEÇA COM RISCO / AMASSADA NO MORDENTE", "27 PEÇA SEM ESTAMPO", "28 DOBRA INVERTIDA (DIREITA/ESQUERDA)", "29 FURO / ROSCA FORA DE POSIÇÃO", "30 PEÇA COM ESQUADRO FORA DO ESPECIFICADO", "31 PEÇA COM ESTAMPO PRÓXIMO A DOBRA", "33 CORDÃO DE SOLDA COM COMPRIMENTO/LARGURA FORA DO ESPECIFICADO", "35 CORDÃO DE SOLDA COM POROSIDADE", "36 CORDÃO DE SOLDA COM RESPINGOS EM EXCESSO", "37 CORDÃO DE SOLDA DESLOCADO", "38 CORDÃO DE SOLDA SEM PENETRAÇÃO", "39 PEÇA SEM SOLDA", "40 SOLDA COM ESTÉTICA FORA DO PADRÃO", "41 PEÇA COM GARRA TROCADA", "42 PEÇA FURADA NA REGIÃO DO CORDÃO", "43 TUBO COM ABAS AMASSADAS", "44 COMPONENTE SOLDADO TROCADO", "45 LARGURA FORA DO ESPECIFICADO", "46 FURO DA PEÇA COM REBARBA EXCESSIVA", "47 COMPONENTE SOLDADO FALTANTE OU FORA DE POSIÇÃO", "48 DESVIO DE FORMA NO ARAME", "49 FALHA NA SOLDA DO ARAME", "50 MALHA FORA DO ESPECIFICADO", "52 CAMADA BAIXA", "53 EXCESSO DE TINTA", "54 MARCA DE ÓLEO / FOSFATO / ÁGUA NA PEÇA", "55 MATERIAL DANIFICADO NA ESTUFA", "56 PEÇA COM COR ERRADA", "57 PINTURA QUEIMADA", "58 PINTURA SEM ADERÊNCIA", "59 PINTURA SEM CURA", "60 TINTA CONTAMINADA", "61 MATERIAL DANIFICADO POR EMBALAGEM INCORRETA", "62 FALHA NO CARREGAMENTO", "64 MATERIAL BLOQUEADO ALOCADO NO PÁTIO", "66 MATERIAL DE TERCEIROS NÃO ALOCADO NA EXPEDIÇÃO", "70 SALDO NO EXP E FÍSICO EM OUTRO SETOR", "73 PEÇA COM REBARBA EXCESSIVA", "74 PEÇA OU CHAPA ENFERRUJADA", "75 PEÇA PROGRAMADA DIVERGENTE AO DESENHO ATUAL", "76 PEÇA QUEIMADA", "77 PEÇA SOBRANDO OU FALTANDO", "79 CHAPA COM CORTE NÃO REALIZADO OU FINALIZADO", "80 PEÇA OU CHAPA COM OXIDAÇÃO EXCESSIVA", "81 CHAPA ONDULADA", "82 CHAPA DANIFICADA DURANTE O CORTE", "83 SLITTER COM EMBOBINAMENTO NÃO CONFORME", "84 SLITTER COM FALHAS NAS BORDAS", "85 SLITTER COM MEDIDA DIVERGENTE", "86 SLITTER COM OXIDAÇÃO EXCESSIVA", "87 SLITTER COM REBARBA EXCESSIVA", "88 SLITTER ONDULADA", "89 TUBO COM FALHA NA SOLDA", "90 TUBO/METALON CORTADO SEM ÂNGULO / FORA DE ÂNGULO", "91 ACESSÓRIO DE MICROPISTA MONTADO DIVERGENTE", "92 PISTA COM APERTO INSUFICIENTE", "93 TUBO OU MANCAL DE ESPESSURA DIVERGENTE", "99 ARAME SOLDADO FORA DE POSIÇÃO", "100 PEÇA COM EMENDA DE SOLDA", "101 FALTA DE PEÇAS PARA CONCLUSÃO DA ORDEM DE PRODUÇÃO DO ITEM PAI", "103 PEÇAS REPORTADAS EM OUTRAS OPERAÇÕES OU COM IMPOSSIBILIDADE DE PROGRAMAÇÃO", "104 PLACAS / PEÇAS NÃO PRODUZIDAS OU PRODUZIDAS ERRADAS", "105 PRODUÇÃO OU ENVIO DO ITEM ERRADO", "106 PRODUÇÃO OU COMPRA DE ITENS DUPLICADOS", "107 FALTA DE ACESSÓRIOS NO CLIENTE", "108 MATERIAL SEM REPORTE DA ÚLTIMA OPERAÇÃO", "109 PEÇA COM RISCOS E MARCAS DO PROCESSO", "110 PEÇA DESENVOLVIDA INCOMPATÍVEL COM A MONTAGEM", "111 PROJETO/LAYOUT INCORRETO", "112 FALTA DE INFORMAÇÃO NO LAYOUT", "113 PEÇA/MÁQUINA CHEGOU COM PROBLEMA", "114 MONTAGEM INTERNA INCORRETA", "115 PEÇA COM POSSIBILIDADE DE PRODUÇÃO INCORRETA", "116 ENCONTRADO MATERIAL DENTRO DA MÁQUINA", "117 PISO COM LARGURA MENOR DE MONTAGEM", "118 LONA DANIFICADA", "119 FALTA DE CORREÇÃO DA ORDEM (QUANTIDADE)", "120 MATERIAL ENVIADO INCORRETO", "122 PEÇA OU CARACTERÍSTICA INCOMPATÍVEL COM FERRAMENTAS", "123 PEÇA COM FURAÇÃO INCORRETA", "124 PEÇA ATRASADA NA PROGRAMAÇÃO", "126 ACESSÓRIO INCOMPATÍVEL COM NECESSIDADE", "127 CABO FICOU CURTO", "128 CABO INCORRETO", "129 ESQUEMA INCORRETO", "130 ATUALIZAÇÃO DE PROJETO / DESDOBRO", "131 PINTURA COM ESTÉTICA FORA DO PADRÃO", "132 ERRO NO PROJETO ELÉTRICO", "133 FALTA DE INFORMAÇÃO NO ESQUEMA", "134 LIGAÇÃO DO PAINEL INCORRETA", "135 ERRO NA MONTAGEM EXTERNA", "136 LIGAÇÃO ELÉTRICA DA PISTA INCORRETA", "137 FALTA DE COMPONENTE ELÉTRICO", "138 ACABAMENTOS DIFERENTES MISTURADOS NA EMBALAGEM", "139 INFORMAÇÕES INCORRETAS NO PROJETO", "141 CABO DANIFICADO", "142 PEÇAS DE DIFERENTES COMPRIMENTOS NA MESMA EMBALAGEM", "143 PEÇA PINTADA SEM O PROCESSO ANTERIOR", "144 PEÇAS COM PINTURA ZEBRADA", "146 PINTURA COM DIVERGÊNCIA NA TONALIDADE", "147 PINTURA COM ASPECTO ÁSPERO", "148 FARDO COM PESO EXCESSIVO DO SUPORTADO PELA EMPILHADEIRA", "152 LIGAÇÕES DO ESQUEMA ELÉTRICO INVERTIDAS", "156 ACESSÓRIO ELÉTRICO INCOMPATÍVEL COM NECESSIDADE NA OBRA", "157 ACESSÓRIO COM IMPOSSIBILIDADE DE VISUALIZAÇÃO", "158 MONTAGEM DE ACESSÓRIOS DIVERGENTE DO DESENHO", "159 PISTA FORA DE ESQUADRO", "161 ROLO MOTOR NO INÍCIO DA PISTA", "165 MONTAGEM DO TRANSFER COM CORREIAS POSICIONADAS INCORRETAMENTE", "167 VOLUME DANIFICADO NO TRANSPORTE", "168 MATERIAL DANIFICADO NO TRANSPORTE", "169 EMBALAGEM MAL FEITA", "170 CARGA MAL AMARRADA", "171 CORTE COM ESTÉTICA FORA DO PADRÃO", "183 TRANSPORTE SEM ACESSÓRIOS DE CARGA EXIGIDOS", "184 TRANSPORTE COM ACESSÓRIOS DE CARGAS DANIFICADOS", "186 DESCARREGAMENTO DE MATERIAIS NO CLIENTE ERRADO"], key="nao_conf")
        with iq2: n_nc = st.text_input("Nº NC", key="n_nc")
        
        iq3, iq4 = st.columns(2)
        with iq3: cc_origem = st.selectbox("Centro Custo - Setor de Origem", [" ", "11000", "11001", "11002", "11003", "11004", "11006", "11007", "11008", "11009", "11010", "11011", "11021", "11022", "11023", "11024", "11025", "11026", "11027", "11028", "11029", "11030", "11031", "17101", "17104", "17105", "17106", "21000", "21001", "21005", "21007", "21009", "21013", "21014", "21017", "21018", "21020", "21022", "21024", "21025", "21026", "21027", "21028", "21029", "21030", "21031", "21032", "21033", "27001", "27002", "27003", "27004", "27005", "27007", "31030", "31201", "31211", "31212", "31213", "31221", "31222", "31223", "31224", "31225", "31231", "31232", "31233", "31234", "31235", "31241", "31251", "31301", "31311", "31312", "31313", "31314", "31315", "31316", "31317", "31321", "31322", "31331", "31341", "31401", "31411", "31501", "31502", "31601", "31611", "31701", "31801", "31911", "31912", "37001", "37002", "37003", "37004", "37005"], key="cc_origem")
        with iq4: setor_origem = st.selectbox("Setor Origem da NC",["", "DIRETORIA", "ADMINISTRATIVO", "VENDAS - PG", "VENDAS - SP", "MONTAGEM", "COMEX - COMÉRCIO EXTERIOR", "TRANSPORTES", "MARKETING", "ENGENHARIA", "INOVAÇÃO", "EXPEDIÇÃO", "TI - TECNOLOGIA DA INFORMAÇÃO", "CONTROLADORIA", "JURÍDICO", "FINANCEIRO", "RH - RECURSOS HUMANOS", "PORTARIA", "MEIO AMBIENTE", "ENGENHARIA INDUSTRIAL", "COLABORADORES AFASTADOS", "LIMPEZA", "COMPRAS", "PROJETOS MECÂNICOS", "PROJETOS ELÉTRICOS", "DESENVOLVIMENTO", "EXPEDIÇÃO - AUTOMAÇÃO", "REFEITÓRIO", "SEGURANÇA DO TRABALHO", "P.C.P. ESTRUTURAS", "MANUTENÇÃO DE MÁQUINAS", "FERRAMENTARIA", "SAÚDE OCUPACIONAL", "GERENCIA E LIDERANÇA INDUSTRIAL", "MANUTENÇÃO ELÉTRICA", "MANUTENÇÃO PREDIAL", "QUALIDADE INDUSTRIAL", "MONTAGEM - EXTERNA", "SUPRIMENTOS", "ENGENHARIA DE OBRAS", "DESDOBRO", "INDUSTRIALIZAÇÃO", "FERRAMENTARIA - MANUTENÇÃO", "ENGENHARIA DE FERRAMENTAS", "PCM E LIDERANÇA DE MANUTENÇÃO", "EMPILHADEIRAS", "MOVIMENTAÇÃO DE BOBINAS", "EQUIPAMENTOS DE APOIO", "ENGENHARIA DE SOFTWARE", "MONTAGEM EXTERNA AUTOMAÇÃO", "P.C.P. AUTOMAÇÃO", "SUPERVISOR MECÂNICA", "SUPERVISOR TRANSPORTADOR", "SUPRIMENTOS", "ORDENS MANUTENÇÃO", "CORTE TRANSVERSAL - DIVIMEC", "LASER - BYSTRONIC", "LASER - TRUMPF 3030", "LASER - TRUMPF 1040", "DOBRADEIRA GASPARINI 110/3", "DOBRADEIRA BRAFFEMAN 130/3", "DOBRADEIRA GASPARINI 135/3", "DOBRADEIRA GASPARINI 250/4", "DOBRADEIRA GASPARINI 200/4", "SOLDA MANUAL - FÁBRICA 01", "SOLDA MANUAL - FÁBRICA 02", "SOLDA KAWASAKI", "SOLDA KAWASAKI TWIN", "SOLDA PANASONIC", "GUILHOTINA", "PRENSAS", "CORTE LONGITUDINAL - DIVIMEC", "PR 310 - TIANFON II", "PR 200 - ZIKELLI I", "PR 200 - ZIKELLI II", "PR 350 - ZIKELLI", "PR 400 - ZIKELLI", "PR 600 - ZIKELLI", "PR 300 - TIANFON I", "ESTAMPO CONTÍNUO - FÁBRICA 02", "ESTAMPO CONTÍNUO - FÁBRICA 03", "PR LONGARINAS - ZIKELLI", "SOLDA GME", "CORTE DE ARAMES", "PONTEADEIRAS", "CORTE DE TUBOS", "SERRA FITA DE PERFIS", "INJETORAS", "MONTAGEM DE REDUTORES", "DINÂMICO", "FABRICAÇÃO DE EMBALAGEM - EXTERNA", "LINHA DE PINTURA 04", "LINHA DE PINTURA 03", "MONTAGEM ELÉTRICA", "MONTAGEM MECÂNICA", "MONTAGEM MÁQUINAS", "MONTAGEM PISTAS", "MONTAGEM CAVALETES"], key="setor_origem")
        
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
        with it1: fornecedor = st.selectbox("Fornecedor da Tinta / Aço", [" ", "PRINCELUX", "WEG", "RENNER", "SUPERLACK"], key="fornecedor")
        with it2: cor_tinta = st.selectbox("Cor da Tinta", [" ", "AMARELO 1003", "AMARELO 1023", "AMARELO-LARANJA-2000", "AZUL 10B", "BEGE 2.5", "BRANCO 9003", "CINZA 7012", "CINZA 7035", "CINZA N6.5", "LARANJA 2.5", "VERDE 6013", "VERDE 2.5", "VERMELHO-RAL-3020", "LARANJA 2000", "AMARELO 1021", "ALUMÍNIO-BRANCO-METÁLICO-RAAL-9006"], key="cor_tinta")
        
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