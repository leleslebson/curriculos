# 🪪 CV Inteligente

> Extrator inteligente de currículos com IA — converta PDFs e documentos Word em dados estruturados exportáveis para Excel.

![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=flat&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.x-FF4B4B?style=flat&logo=streamlit&logoColor=white)
![LangChain](https://img.shields.io/badge/LangChain-0.x-1C3C3C?style=flat&logo=chainlink&logoColor=white)
![Groq](https://img.shields.io/badge/Groq-Llama_3.3_70B-F55036?style=flat)
![License](https://img.shields.io/badge/Licença-MIT-green?style=flat)

Demonstração APP
https://lercurriculos.streamlit.app/

---

## 📋 Sumário

- [Visão Geral](#-visão-geral)
- [Funcionalidades](#-funcionalidades)
- [Demonstração](#-demonstração)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Configuração](#-configuração)
- [Como Usar](#-como-usar)
- [Estrutura do Projeto](#-estrutura-do-projeto)
- [Stack Tecnológico](#-stack-tecnológico)
- [Privacidade e Segurança](#-privacidade-e-segurança)
- [Contribuição](#-contribuição)
- [Licença](#-licença)

---

## 🧠 Visão Geral

**CV Inteligente** é uma aplicação web construída com Streamlit que utiliza o modelo **Llama 3.3 70B** (via Groq API) para extrair automaticamente informações estruturadas de currículos em formato PDF ou Word. Os dados extraídos são exibidos em uma tabela interativa e podem ser exportados para Excel com um único clique.

Ideal para equipes de RH, recrutadores e profissionais que precisam processar grandes volumes de currículos com agilidade e precisão.

---

## ✨ Funcionalidades

- 📁 **Upload múltiplo** — processa vários arquivos simultaneamente (PDF, DOCX, DOC)
- 🤖 **Extração via IA** — usa LLM para identificar campos mesmo em layouts não padronizados
- 📊 **Dashboard de sessão** — acompanhe o total de currículos processados
- 📋 **Tabela interativa** — visualize todos os dados extraídos em uma grade editável
- 📥 **Exportação para Excel** — baixe todos os registros formatados em `.xlsx`
- ⚙️ **Instruções customizáveis** — configure os campos extraídos via arquivo `cv.json`
- 🔒 **Sem retenção de dados** — nenhuma informação é persistida após a sessão

### Campos extraídos por padrão

| Campo       | Descrição                         |
|-------------|-----------------------------------|
| `nome`      | Nome completo do candidato        |
| `cpf`       | CPF (somente números)             |
| `identidade`| RG / número de identidade         |
| `telefone`  | Telefone de contato               |
| `rua`       | Logradouro                        |
| `numero`    | Número do endereço                |
| `bairro`    | Bairro                            |
| `cidade`    | Cidade                            |
| `estado`    | Estado (sigla: SP, RJ...)         |
| `cargo`     | Cargo desejado ou atual           |

---

## 🖥️ Demonstração

```
Dashboard  →  Processar Currículos  →  Sobre
    ↓                  ↓
Métricas          Upload de PDFs/DOCXs
de sessão              ↓
                  Extração via IA
                       ↓
                  Tabela de resultados
                       ↓
                  Download Excel (.xlsx)
```

---

## 📦 Pré-requisitos

- Python **3.11** ou superior
- Uma chave de API válida da **[Groq](https://console.groq.com)**
- pip ou outro gerenciador de pacotes Python

---

## 🚀 Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/cv-inteligente.git
cd cv-inteligente
```

### 2. Crie um ambiente virtual (recomendado)

```bash
python -m venv .venv
source .venv/bin/activate      # Linux/macOS
.venv\Scripts\activate         # Windows
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

**`requirements.txt` esperado:**

```
streamlit
langchain
langchain-groq
langchain-core
pydantic
PyPDF2
python-docx
pandas
openpyxl
streamlit-option-menu
```

---

## ⚙️ Configuração

### Chave de API (obrigatório)

#### Localmente

Crie o arquivo `.streamlit/secrets.toml` na raiz do projeto:

```toml
GROQ_API_KEY = "gsk_xxxxxxxxxxxxxxxxxxxxxxxx"
```

> ⚠️ **Nunca** versione o arquivo `secrets.toml`. Ele já está no `.gitignore` padrão do Streamlit.

#### Streamlit Cloud

No painel do Streamlit Cloud, vá em **Settings → Secrets** e adicione:

```
GROQ_API_KEY = "gsk_xxxxxxxxxxxxxxxxxxxxxxxx"
```

### Personalização de campos (`cv.json`)

Na primeira execução, o arquivo `cv.json` é criado automaticamente na raiz com as instruções padrão. Edite-o para customizar os campos extraídos:

```json
{
  "instrucoes": "Extraia as informações abaixo do currículo",
  "campos": {
    "nome": "Nome completo",
    "cpf": "CPF (apenas números)",
    "cargo": "Cargo desejado/atual"
  }
}
```

---

## 🖱️ Como Usar

### Execute a aplicação

```bash
streamlit run app.py
```

Acesse em: `http://localhost:8501`

### Fluxo de uso

1. Acesse **Processar Currículos** no menu lateral
2. Faça upload de um ou mais arquivos (PDF, DOCX ou DOC)
3. Clique em **Iniciar Extração**
4. Acompanhe o progresso em tempo real
5. Visualize os dados extraídos na tabela
6. Clique em **Exportar para Excel** para baixar os resultados

---

## 🗂️ Estrutura do Projeto

```
cv-inteligente/
├── app.py                  # Aplicação principal
├── cv.json                 # Configuração de campos (gerado automaticamente)
├── requirements.txt        # Dependências Python
├── .streamlit/
│   └── secrets.toml        # Chave API (não versionar)
└── README.md
```

---

## 🛠️ Stack Tecnológico

| Tecnologia | Função |
|---|---|
| [Streamlit](https://streamlit.io) | Framework de interface web |
| [LangChain](https://www.langchain.com) | Orquestração do pipeline de IA |
| [Groq API](https://groq.com) | Inferência ultrarrápida de LLM |
| Llama 3.3 70B | Modelo de linguagem para extração |
| [PyPDF2](https://pypdf2.readthedocs.io) | Leitura de arquivos PDF |
| [python-docx](https://python-docx.readthedocs.io) | Leitura de arquivos Word |
| [Pandas](https://pandas.pydata.org) | Manipulação de dados tabulares |
| [openpyxl](https://openpyxl.readthedocs.io) | Exportação para Excel |
| [streamlit-option-menu](https://github.com/victoryhb/streamlit-option-menu) | Navegação lateral customizada |

---

## 🔒 Privacidade e Segurança

- Os textos dos currículos são enviados **temporariamente** à API Groq para inferência
- **Nenhum dado é gravado** em banco de dados ou armazenamento externo
- Todos os registros são **descartados ao encerrar a sessão**
- A chave de API é carregada via `st.secrets` e nunca exposta no código-fonte
- Recomenda-se uso em ambientes privados ou com HTTPS habilitado em produção

---

## 🤝 Contribuição

Contribuições são bem-vindas! Para contribuir:

1. Faça um fork do repositório
2. Crie uma branch para sua feature: `git checkout -b feature/minha-feature`
3. Commit suas alterações: `git commit -m 'feat: adiciona nova funcionalidade'`
4. Push para a branch: `git push origin feature/minha-feature`
5. Abra um Pull Request

Por favor, siga o padrão [Conventional Commits](https://www.conventionalcommits.org/pt-br/) nas mensagens de commit.

---

## 📄 Licença

Este projeto está licenciado sob a **MIT License**. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.

---

<div align="center">
  Feito com ❤️ e <strong>Groq + Llama 3.3</strong>
</div>
