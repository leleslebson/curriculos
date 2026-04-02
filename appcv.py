"""
Sistema de Extração de Dados de Currículos para Streamlit Cloud
Utiliza LangChain + Groq (Llama) e exporta para Excel.
Layout moderno, sem campo de API key (usa st.secrets).
"""

import streamlit as st
import json
import os
import re
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd
from io import BytesIO

# Extração de texto
import PyPDF2
from docx import Document

# LangChain
from langchain_groq import ChatGroq
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import PydanticOutputParser
from pydantic import BaseModel, Field

# =======================
# MODELO DE DADOS
# =======================
class CurriculoData(BaseModel):
    nome: str = Field(default="", description="Nome completo")
    cpf: str = Field(default="", description="CPF (apenas números)")
    identidade: str = Field(default="", description="RG/Identidade")
    telefone: str = Field(default="", description="Telefone")
    rua: str = Field(default="", description="Rua/Logradouro")
    numero: str = Field(default="", description="Número do endereço")
    bairro: str = Field(default="", description="Bairro")
    cidade: str = Field(default="", description="Cidade")
    estado: str = Field(default="", description="Estado (sigla)")
    cargo: str = Field(default="", description="Cargo desejado ou atual")

# =======================
# FUNÇÕES DE EXTRAÇÃO DE TEXTO
# =======================
def extrair_texto_pdf(arquivo) -> str:
    try:
        reader = PyPDF2.PdfReader(arquivo)
        texto = "".join([page.extract_text() or "" for page in reader.pages])
        return texto.strip()
    except Exception as e:
        raise Exception(f"Erro ao ler PDF: {str(e)}")

def extrair_texto_word(arquivo) -> str:
    try:
        doc = Document(arquivo)
        texto = "\n".join([p.text for p in doc.paragraphs])
        return texto.strip()
    except Exception as e:
        raise Exception(f"Erro ao ler Word: {str(e)}")

def extrair_texto_arquivo(arquivo) -> str:
    nome = arquivo.name.lower()
    if nome.endswith(".pdf"):
        return extrair_texto_pdf(arquivo)
    elif nome.endswith((".docx", ".doc")):
        return extrair_texto_word(arquivo)
    else:
        raise Exception(f"Formato não suportado: {nome}")

# =======================
# INSTRUÇÕES JSON
# =======================
def carregar_instrucoes_json() -> Dict[str, Any]:
    """Carrega cv.json da raiz do projeto."""
    caminho = Path("cv.json")
    if not caminho.exists():
        # Cria um padrão caso o arquivo não exista (nunca deve ocorrer, mas seguro)
        padrao = {
            "instrucoes": "Extraia as informações abaixo do currículo",
            "campos": {
                "nome": "Nome completo",
                "cpf": "CPF (apenas números)",
                "identidade": "RG ou identidade",
                "telefone": "Telefone",
                "rua": "Rua",
                "numero": "Número",
                "bairro": "Bairro",
                "cidade": "Cidade",
                "estado": "Estado (sigla)",
                "cargo": "Cargo desejado/atual"
            }
        }
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(padrao, f, ensure_ascii=False, indent=2)
        return padrao
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)

# =======================
# PROCESSAMENTO COM IA (USANDO SECRETS)
# =======================
def processar_curriculo_com_ia(texto_curriculo: str) -> Dict[str, str]:
    """Usa a chave API do st.secrets['GROQ_API_KEY']."""
    try:
        api_key = st.secrets["GROQ_API_KEY"]
    except KeyError:
        raise Exception("Chave API não encontrada. Configure o secret GROQ_API_KEY no Streamlit Cloud.")

    instrucoes = carregar_instrucoes_json()
    parser = PydanticOutputParser(pydantic_object=CurriculoData)

    llm = ChatGroq(
        model="llama-3.3-70b-versatile",
        api_key=api_key,
        temperature=0,
        max_retries=2,
    )

    prompt = ChatPromptTemplate.from_messages([
        ("system", """Você é um assistente especializado em extrair informações de currículos.
        Instruções do JSON: {instrucoes_json}

        IMPORTANTE:
        - Campos não encontrados → deixe vazio ("").
        - CPF → apenas números.
        - Telefone → formato original.
        - Estado → sigla (SP, RJ, MG...).
        - Seja literal e preciso.

        {format_instructions}
        """),
        ("user", "Currículo:\n\n{curriculo}")
    ])

    chain = prompt | llm | parser
    resultado = chain.invoke({
        "curriculo": texto_curriculo,
        "instrucoes_json": json.dumps(instrucoes, ensure_ascii=False, indent=2),
        "format_instructions": parser.get_format_instructions()
    })
    return resultado.dict()

# =======================
# EXPORTAÇÃO EXCEL
# =======================
def gerar_excel(dados_curriculos: List[Dict[str, str]]) -> BytesIO:
    colunas = ["nome", "cpf", "identidade", "telefone", "rua", "numero", "bairro", "cidade", "estado", "cargo"]
    df = pd.DataFrame(dados_curriculos, columns=colunas)
    df.columns = [col.capitalize() for col in df.columns]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Currículos")
        ws = writer.sheets["Currículos"]
        for i, col in enumerate(df.columns, 1):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[chr(64 + i)].width = min(max_len, 50)
    output.seek(0)
    return output

# =======================
# LAYOUT MODERNO + CSS
# =======================
def aplicar_css():
    st.markdown("""
    <style>
        /* Fonte moderna */
        @import url('https://fonts.googleapis.com/css2?family=Inter:opsz@14..32&display=swap');
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        /* Título principal */
        .main-title {
            background: linear-gradient(120deg, #2b5876, #4e4376);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            font-weight: 700;
            font-size: 2.8rem;
            margin-bottom: 0;
        }
        /* Card personalizado */
        .custom-card {
            background-color: #f8f9fa;
            border-radius: 20px;
            padding: 1.2rem;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            margin-bottom: 1rem;
            border: 1px solid #e9ecef;
        }
        /* Botão primário */
        .stButton > button {
            background: linear-gradient(90deg, #4e4376, #2b5876);
            color: white;
            border: none;
            border-radius: 40px;
            padding: 0.5rem 1.8rem;
            font-weight: 600;
            transition: 0.2s;
        }
        .stButton > button:hover {
            transform: scale(1.02);
            background: linear-gradient(90deg, #2b5876, #4e4376);
            color: white;
        }
        /* Upload box */
        .stFileUploader > div {
            border: 2px dashed #4e4376;
            border-radius: 20px;
            background-color: #fef9e6;
        }
        /* Sidebar */
        [data-testid="stSidebar"] {
            background-color: #f1f3f5;
            border-right: 1px solid #dee2e6;
        }
        /* Badges */
        .badge {
            background-color: #4e4376;
            color: white;
            padding: 4px 10px;
            border-radius: 40px;
            font-size: 0.75rem;
            display: inline-block;
            margin-right: 8px;
        }
        hr {
            margin: 1rem 0;
        }
    </style>
    """, unsafe_allow_html=True)

# =======================
# AVISO LGPD (consentimento)
# =======================
def exibir_aviso_lgpd():
    with st.expander("📜 Aviso de Privacidade (LGPD)", expanded=False):
        st.markdown("""
        Os currículos enviados serão processados temporariamente por inteligência artificial (API Groq).  
        Nenhum dado é armazenado permanentemente após o término da sessão.  
        Ao utilizar este sistema, você autoriza o envio dos dados para processamento externo.  
        **Não envie dados sensíveis além dos necessários para a seleção.**  
        """)
        if not st.session_state.get("consentimento", False):
            consentido = st.checkbox("Li e concordo com o processamento dos dados conforme descrito.")
            if consentido:
                st.session_state.consentimento = True
                st.rerun()
            else:
                st.stop()  # Impede o uso até concordar

# =======================
# MAIN
# =======================
def main():
    st.set_page_config(page_title="Extrator de Currículos", page_icon="📄", layout="wide")
    aplicar_css()

    # Título elegante
    st.markdown('<p class="main-title">📄 Extrator Inteligente de Currículos</p>', unsafe_allow_html=True)
    st.markdown("Extraia dados estruturados (nome, CPF, endereço, cargo...) de PDFs e Word em segundos.")
    st.markdown("---")

    # Verificar secrets
    if "GROQ_API_KEY" not in st.secrets:
        st.error("🚨 **Erro crítico:** Chave da API Groq não configurada. O administrador do sistema deve adicionar `GROQ_API_KEY` nos secrets do Streamlit Cloud.")
        st.stop()

    # Aviso LGPD com consentimento obrigatório
    exibir_aviso_lgpd()

    # Sidebar moderna
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/curriculum.png", width=80)
        st.markdown("### ⚙️ Sobre o sistema")
        st.markdown("""
        <div class="custom-card">
        ✅ <strong>Formatos aceitos</strong><br>
        📄 PDF &nbsp;&nbsp;|&nbsp;&nbsp;📝 Word (.docx)
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### 🧠 Campos extraídos")
        cols = st.columns(2)
        campos = ["Nome", "CPF", "Identidade", "Telefone", "Rua", "Número", "Bairro", "Cidade", "Estado", "Cargo"]
        for i, campo in enumerate(campos):
            cols[i%2].markdown(f"`• {campo}`")

        st.markdown("---")
        st.caption("🔒 Processamento seguro via API Groq (Llama 3.3)")

    # Upload de arquivos
    st.subheader("📂 1. Envie os currículos")
    arquivos = st.file_uploader(
        "Arraste ou clique para selecionar",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if not arquivos:
        st.info("👆 Aguardando arquivos...")
        return

    st.success(f"{len(arquivos)} arquivo(s) carregado(s)")

    if st.button("🚀 Processar agora", type="primary", use_container_width=False):
        dados_extraidos = []
        progresso = st.progress(0, text="Iniciando...")
        status_area = st.empty()

        for idx, arquivo in enumerate(arquivos):
            status_area.info(f"📑 Processando: **{arquivo.name}**")
            try:
                texto = extrair_texto_arquivo(arquivo)
                if len(texto.strip()) < 50:
                    st.warning(f"⚠️ {arquivo.name} – texto muito curto (possivelmente imagem/scanner não legível).")
                    continue

                dados = processar_curriculo_com_ia(texto)
                dados_extraidos.append(dados)
                st.success(f"✅ {arquivo.name} concluído")
            except Exception as e:
                st.error(f"❌ {arquivo.name} – erro: {str(e)}")

            progresso.progress((idx + 1) / len(arquivos), text=f"Progresso: {idx+1}/{len(arquivos)}")

        status_area.empty()
        progresso.empty()

        if dados_extraidos:
            st.markdown("---")
            st.subheader("📊 2. Dados extraídos (prévia)")
            df_preview = pd.DataFrame(dados_extraidos)
            df_preview.columns = [c.capitalize() for c in df_preview.columns]
            st.dataframe(df_preview, use_container_width=True, height=300)

            st.subheader("📎 3. Download Excel")
            excel_buffer = gerar_excel(dados_extraidos)
            st.download_button(
                label="📥 Baixar planilha .xlsx",
                data=excel_buffer,
                file_name="curriculos_processados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("Nenhum dado foi extraído. Verifique se os currículos contêm texto legível.")

if __name__ == "__main__":
    main()
