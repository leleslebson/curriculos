"""
CV Intelligence — Extrator de Currículos
CSS mínimo e cirúrgico: apenas o que o Streamlit permite sobrescrever com segurança.
"""

import streamlit as st
import json
import pandas as pd
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Any

import PyPDF2
from docx import Document

from langchain_groq import ChatGroq
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import PydanticOutputParser
from pydantic import BaseModel, Field

from streamlit_option_menu import option_menu

# ══════════════════════════════════════════
# CONFIGURAÇÃO DA PÁGINA
# ══════════════════════════════════════════
st.set_page_config(
    page_title="CV Intelgente ",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(
    '<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">',
    unsafe_allow_html=True
)

# ══════════════════════════════════════════
# CSS — MÍNIMO E SEGURO
# Regra: só sobrescrever seletores estáveis
# do Streamlit. Nunca tocar em componentes
# que têm iframe próprio (option_menu).
# ══════════════════════════════════════════
def aplicar_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

        /* Fonte base em todo o app */
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif !important;
        }

        /* Área de conteúdo principal */
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1160px !important;
        }

        /* Botão primário */
        .stButton > button[kind="primary"] {
            background-color: #2563eb !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            font-size: 0.875rem !important;
            padding: 0.5rem 1.6rem !important;
            box-shadow: 0 1px 4px rgba(37,99,235,0.25) !important;
            transition: background 0.15s, transform 0.1s !important;
        }
        .stButton > button[kind="primary"]:hover {
            background-color: #1d4ed8 !important;
            transform: translateY(-1px) !important;
        }

        /* Botão secundário */
        .stButton > button[kind="secondary"] {
            border-radius: 6px !important;
            font-size: 0.875rem !important;
        }

        /* Barra de progresso */
        div[data-testid="stProgressBar"] > div {
            background-color: #2563eb !important;
        }

        /* Remover menu hamburger e footer */
        #MainMenu, footer { visibility: hidden; }

        /* ── Componentes HTML customizados ── */

        .cv-page-header {
            display: flex;
            align-items: flex-start;
            gap: 14px;
            margin-bottom: 1.5rem;
        }
        .cv-icon-box {
            width: 42px;
            height: 42px;
            min-width: 42px;
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #2563eb;
            font-size: 1.05rem;
        }
        .cv-page-title {
            font-size: 1.4rem;
            font-weight: 700;
            color: #111827;
            letter-spacing: -0.3px;
            line-height: 1.2;
            margin: 0;
        }
        .cv-page-subtitle {
            font-size: 0.84rem;
            color: #6b7280;
            margin-top: 4px;
        }

        .cv-section-label {
            display: flex;
            align-items: center;
            gap: 7px;
            font-size: 0.72rem;
            font-weight: 700;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.09em;
            margin: 1.2rem 0 0.7rem;
        }
        .cv-section-label::after {
            content: '';
            flex: 1;
            height: 1px;
            background: #e5e7eb;
        }

        .cv-badge {
            display: inline-flex;
            align-items: center;
            gap: 7px;
            font-size: 0.82rem;
            font-weight: 500;
            padding: 6px 13px;
            border-radius: 6px;
            margin: 5px 0;
            font-family: 'Inter', sans-serif;
        }
        .cv-badge-success { background:#f0fdf4; color:#15803d; border:1px solid #bbf7d0; }
        .cv-badge-error   { background:#fef2f2; color:#b91c1c; border:1px solid #fecaca; }
        .cv-badge-warning { background:#fffbeb; color:#b45309; border:1px solid #fde68a; }
        .cv-badge-info    { background:#eff6ff; color:#1d4ed8; border:1px solid #bfdbfe; }

        .cv-api-error {
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-left: 4px solid #dc2626;
            border-radius: 8px;
            padding: 1rem 1.2rem;
            display: flex;
            align-items: flex-start;
            gap: 12px;
            font-size: 0.875rem;
            color: #7f1d1d;
        }

        .cv-stat-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            padding: 1.1rem 1.3rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .cv-stat-label {
            font-size: 0.7rem;
            font-weight: 700;
            color: #9ca3af;
            text-transform: uppercase;
            letter-spacing: 0.08em;
        }
        .cv-stat-value {
            font-size: 1.8rem;
            font-weight: 700;
            color: #111827;
            font-family: 'JetBrains Mono', monospace;
            line-height: 1.2;
            margin-top: 4px;
        }

        .cv-about-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            padding: 1.4rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .cv-about-card h4 {
            font-size: 0.75rem !important;
            font-weight: 700 !important;
            color: #6b7280 !important;
            text-transform: uppercase !important;
            letter-spacing: 0.08em !important;
            margin: 0 0 0.8rem !important;
        }
        .cv-about-row {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.875rem;
            color: #374151;
            padding: 5px 0;
            border-bottom: 1px solid #f9fafb;
        }
        .cv-about-row:last-child { border-bottom: none; }
        .cv-about-icon {
            color: #2563eb;
            width: 18px;
            text-align: center;
            flex-shrink: 0;
        }
    </style>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════
# HELPERS HTML
# ══════════════════════════════════════════
def page_header(icon: str, title: str, subtitle: str = ""):
    sub = f'<div class="cv-page-subtitle">{subtitle}</div>' if subtitle else ""
    st.markdown(f"""
    <div class="cv-page-header">
        <div class="cv-icon-box"><i class="{icon}"></i></div>
        <div>
            <div class="cv-page-title">{title}</div>
            {sub}
        </div>
    </div>""", unsafe_allow_html=True)

def section_label(icon: str, label: str):
    st.markdown(
        f'<div class="cv-section-label"><i class="{icon}"></i> {label}</div>',
        unsafe_allow_html=True
    )

def badge_success(msg):
    st.markdown(f'<div class="cv-badge cv-badge-success"><i class="fa-solid fa-circle-check"></i> {msg}</div>', unsafe_allow_html=True)

def badge_error(msg):
    st.markdown(f'<div class="cv-badge cv-badge-error"><i class="fa-solid fa-circle-xmark"></i> {msg}</div>', unsafe_allow_html=True)

def badge_warning(msg):
    st.markdown(f'<div class="cv-badge cv-badge-warning"><i class="fa-solid fa-triangle-exclamation"></i> {msg}</div>', unsafe_allow_html=True)

def badge_info(msg):
    st.markdown(f'<div class="cv-badge cv-badge-info"><i class="fa-solid fa-circle-info"></i> {msg}</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════
# MODELO DE DADOS
# ══════════════════════════════════════════
class CurriculoData(BaseModel):
    nome: str       = Field(default="", description="Nome completo")
    cpf: str        = Field(default="", description="CPF (apenas números)")
    identidade: str = Field(default="", description="RG/Identidade")
    telefone: str   = Field(default="", description="Telefone")
    rua: str        = Field(default="", description="Rua/Logradouro")
    numero: str     = Field(default="", description="Número do endereço")
    bairro: str     = Field(default="", description="Bairro")
    cidade: str     = Field(default="", description="Cidade")
    estado: str     = Field(default="", description="Estado (sigla)")
    cargo: str      = Field(default="", description="Cargo desejado ou atual")


# ══════════════════════════════════════════
# EXTRAÇÃO DE TEXTO
# ══════════════════════════════════════════
def extrair_texto_pdf(arquivo) -> str:
    try:
        reader = PyPDF2.PdfReader(arquivo)
        return "".join([p.extract_text() or "" for p in reader.pages]).strip()
    except Exception as e:
        raise Exception(f"Falha na leitura do PDF: {e}")

def extrair_texto_word(arquivo) -> str:
    try:
        doc = Document(arquivo)
        return "\n".join([p.text for p in doc.paragraphs]).strip()
    except Exception as e:
        raise Exception(f"Falha na leitura do Word: {e}")

def extrair_texto_arquivo(arquivo) -> str:
    nome = arquivo.name.lower()
    if nome.endswith(".pdf"):
        return extrair_texto_pdf(arquivo)
    elif nome.endswith((".docx", ".doc")):
        return extrair_texto_word(arquivo)
    raise Exception(f"Formato não suportado: {nome}")


# ══════════════════════════════════════════
# INSTRUÇÕES JSON
# ══════════════════════════════════════════
def carregar_instrucoes_json() -> Dict[str, Any]:
    caminho = Path("cv.json")
    if not caminho.exists():
        padrao = {
            "instrucoes": "Extraia as informações abaixo do currículo",
            "campos": {
                "nome": "Nome completo", "cpf": "CPF (apenas números)",
                "identidade": "RG ou identidade", "telefone": "Telefone",
                "rua": "Rua", "numero": "Número", "bairro": "Bairro",
                "cidade": "Cidade", "estado": "Estado (sigla)",
                "cargo": "Cargo desejado/atual"
            }
        }
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(padrao, f, ensure_ascii=False, indent=2)
        return padrao
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


# ══════════════════════════════════════════
# PROCESSAMENTO COM IA
# ══════════════════════════════════════════
def processar_curriculo_com_ia(texto: str) -> Dict[str, str]:
    try:
        api_key = st.secrets["GROQ_API_KEY"]
    except KeyError:
        raise Exception("Chave API não configurada nos secrets.")

    instrucoes = carregar_instrucoes_json()
    parser = PydanticOutputParser(pydantic_object=CurriculoData)
    llm = ChatGroq(
        model="llama-3.3-70b-versatile",
        api_key=api_key, temperature=0, max_retries=2
    )
    prompt = ChatPromptTemplate.from_messages([
        ("system", """Você é um assistente especializado em extrair informações de currículos.
Instruções do JSON: {instrucoes_json}
IMPORTANTE:
- Campos não encontrados → deixe vazio ("").
- CPF → apenas números. Telefone → formato original. Estado → sigla (SP, RJ...).
{format_instructions}"""),
        ("user", "Currículo:\n\n{curriculo}")
    ])
    resultado = (prompt | llm | parser).invoke({
        "curriculo": texto,
        "instrucoes_json": json.dumps(instrucoes, ensure_ascii=False, indent=2),
        "format_instructions": parser.get_format_instructions()
    })
    return resultado.dict()


# ══════════════════════════════════════════
# EXPORTAÇÃO EXCEL
# ══════════════════════════════════════════
def gerar_excel(dados: List[Dict[str, str]]) -> BytesIO:
    colunas = ["nome","cpf","identidade","telefone","rua",
               "numero","bairro","cidade","estado","cargo"]
    df = pd.DataFrame(dados, columns=colunas)
    df.columns = [c.capitalize() for c in df.columns]
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Currículos")
        ws = writer.sheets["Currículos"]
        for i, col in enumerate(df.columns, 1):
            w = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[chr(64 + i)].width = min(w, 50)
    output.seek(0)
    return output


# ══════════════════════════════════════════
# MENU LATERAL
# Nota: option_menu renderiza dentro de um
# iframe próprio — os estilos são passados
# pelo dicionário `styles`, não por CSS global.
# ══════════════════════════════════════════
def menu_lateral():
    with st.sidebar:
        # Logo / identidade
        st.markdown("""
        <div style="padding:1.3rem 0.6rem 1.4rem; border-bottom:1px solid #e5e7eb; margin-bottom:0.5rem;">
            <div style="display:flex; align-items:center; gap:10px;">
                <div style="width:36px; height:36px; background:#2563eb; border-radius:8px;
                            display:flex; align-items:center; justify-content:center; flex-shrink:0;">
                    <i class="fa-solid fa-id-card-clip" style="color:#fff; font-size:0.9rem;"></i>
                </div>
                <div>
                    <div style="font-family:'Inter',sans-serif; font-size:0.9rem;
                                font-weight:700; color:#111827; letter-spacing:-0.2px;">
                        CV Inteligente
                    </div>
                    <div style="font-family:monospace; font-size:0.67rem; color:#9ca3af;">
                        Leles Lebson · v2.0
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # option_menu: estilos declarados aqui dentro do componente
        # são aplicados ao iframe interno — não conflitam com o CSS global
        escolha = option_menu(
            menu_title=None,
            options=["Dashboard", "Processar Currículos", "Sobre"],
            icons=[
                "bar-chart-line",       # Bootstrap Icons (padrão do option_menu)
                "file-earmark-arrow-up",
                "info-circle"
            ],
            default_index=1,
            styles={
                "container": {
                    "padding": "4px 0",
                    "background-color": "transparent",
                },
                "icon": {
                    "color": "#2563eb",
                    "font-size": "15px",
                },
                "nav-link": {
                    "font-size": "14px",
                    "font-weight": "400",
                    "color": "#6b7280",
                    "padding": "9px 14px",
                    "border-radius": "6px",
                    "margin": "1px 0",
                },
                "nav-link-selected": {
                    "background-color": "#eff6ff",
                    "font-weight": "600",
                    "color": "#1d4ed8",
                },
            }
        )

        # Rodapé
        st.markdown("""
        <div style="position:fixed; bottom:1.2rem; left:0; width:18rem;
                    padding:0.65rem 1rem 0; border-top:1px solid #f3f4f6;">
            <div style="font-size:0.69rem; color:#9ca3af; font-family:monospace; line-height:1.6;">
                <i class="fa-solid fa-shield-halved" style="margin-right:5px;"></i>
                Processado via Groq API.<br>Nenhum dado é retido.
            </div>
        </div>
        """, unsafe_allow_html=True)

        return escolha


# ══════════════════════════════════════════
# PÁGINAS
# ══════════════════════════════════════════
def pagina_dashboard():
    page_header(
        "fa-solid fa-chart-bar",
        "Dashboard",
        "Visão geral das extrações realizadas na sessão atual"
    )
    st.divider()

    col1, col2, col3 = st.columns(3)
    total = len(st.session_state.get("dados", []))

    with col1:
        st.markdown(f"""
        <div class="cv-stat-card">
            <div class="cv-stat-label">Currículos processados</div>
            <div class="cv-stat-value">{total}</div>
        </div>""", unsafe_allow_html=True)

    with col2:
        taxa = "100%" if total > 0 else "—"
        st.markdown(f"""
        <div class="cv-stat-card">
            <div class="cv-stat-label">Taxa de extração</div>
            <div class="cv-stat-value">{taxa}</div>
        </div>""", unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class="cv-stat-card">
            <div class="cv-stat-label">Tempo médio por arquivo</div>
            <div class="cv-stat-value">—</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    section_label("fa-solid fa-circle-info", "Instruções")
    st.info(
        "Acesse **Processar Currículos** no menu lateral para iniciar a extração de dados via IA.",
        icon=None
    )


def pagina_processar():
    page_header(
        "fa-solid fa-file-arrow-up",
        "Extração de Currículos",
        "Extraia dados estruturados de arquivos PDF e Word"
    )
    st.divider()

    section_label("fa-solid fa-folder-open", "Seleção de Arquivos")

    arquivos = st.file_uploader(
        "Arraste arquivos ou clique para selecionar (PDF, DOCX, DOC)",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True
    )

    if arquivos:
        st.caption(f"{len(arquivos)} arquivo(s) selecionado(s)")

        if st.button("Iniciar Extração", type="primary"):
            dados_extraidos = []
            barra = st.progress(0, text="Iniciando processamento...")
            status = st.empty()

            for idx, arq in enumerate(arquivos):
                pct = (idx + 1) / len(arquivos)
                barra.progress(pct, text=f"Processando {idx+1}/{len(arquivos)}: {arq.name}")
                status.markdown(
                    f'<div class="cv-badge cv-badge-info">'
                    f'<i class="fa-solid fa-rotate fa-spin"></i> {arq.name}</div>',
                    unsafe_allow_html=True
                )
                try:
                    texto = extrair_texto_arquivo(arq)
                    if len(texto.strip()) < 50:
                        badge_warning(f"{arq.name} — conteúdo insuficiente (possível imagem escaneada).")
                        continue
                    dados = processar_curriculo_com_ia(texto)
                    dados_extraidos.append(dados)
                except Exception as e:
                    badge_error(f"{arq.name}: {str(e)}")

            barra.empty()
            status.empty()

            if dados_extraidos:
                st.session_state.dados = dados_extraidos
                badge_success(f"Extração concluída — {len(dados_extraidos)} registro(s) obtido(s).")
            else:
                badge_warning("Nenhum dado pôde ser extraído dos arquivos enviados.")

    # Tabela de resultados
    if st.session_state.get("dados"):
        st.markdown("<br>", unsafe_allow_html=True)
        section_label("fa-solid fa-table", "Resultados da Extração")

        df = pd.DataFrame(st.session_state.dados)
        df_display = df.rename(columns=lambda x: x.capitalize())

        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            height=420
        )

        st.markdown("<br>", unsafe_allow_html=True)
        col_dl, _ = st.columns([2, 5])
        with col_dl:
            st.download_button(
                label="Exportar para Excel",
                data=gerar_excel(st.session_state.dados),
                file_name="curriculos_extraidos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


def pagina_sobre():
    page_header(
        "fa-solid fa-circle-info",
        "Sobre o Sistema",
        "Informações técnicas e de privacidade"
    )
    st.divider()

    col1, col2 = st.columns([3, 2], gap="large")

    with col1:
        section_label("fa-solid fa-layer-group", "Stack Tecnológico")
        st.markdown("""
        <div class="cv-about-card">
            <div class="cv-about-row">
                <i class="fa-brands fa-python cv-about-icon"></i> Python 3.11+
            </div>
            <div class="cv-about-row">
                <i class="fa-solid fa-water cv-about-icon"></i> Streamlit
            </div>
            <div class="cv-about-row">
                <i class="fa-solid fa-link cv-about-icon"></i> LangChain
            </div>
            <div class="cv-about-row">
                <i class="fa-solid fa-microchip cv-about-icon"></i> Groq · Llama 3.3 70B
            </div>
            <div class="cv-about-row">
                <i class="fa-solid fa-table-columns cv-about-icon"></i> AG Grid
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        section_label("fa-solid fa-shield-halved", "Privacidade")
        st.markdown("""
        <div class="cv-about-card">
            <p style="font-size:0.875rem; color:#374151; line-height:1.75; margin:0;">
                Os textos dos currículos são enviados temporariamente à API Groq para inferência.<br><br>
                Nenhum dado é gravado em banco de dados ou retido após o encerramento da sessão.
            </p>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════
def main():
    aplicar_css()

    # Checa API key
    if "GROQ_API_KEY" not in st.secrets:
        st.markdown("""
        <div class="cv-api-error">
            <i class="fa-solid fa-triangle-exclamation" style="color:#dc2626; font-size:1.1rem; flex-shrink:0;"></i>
            <div>
                <strong>Chave API não configurada.</strong><br>
                Adicione <code>GROQ_API_KEY</code> nos Secrets do Streamlit Cloud.
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    pagina = menu_lateral()

    if pagina == "Dashboard":
        pagina_dashboard()
    elif pagina == "Processar Currículos":
        pagina_processar()
    elif pagina == "Sobre":
        pagina_sobre()


if __name__ == "__main__":
    main()
