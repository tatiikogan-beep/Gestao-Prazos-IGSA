# IGSA · Gestão de Prazos — Deploy no Streamlit Cloud

## Pré-requisitos
- Conta no GitHub (já existe: tatiikogan-bip)
- Conta no Streamlit Cloud (https://share.streamlit.io) — gratuita

---

## PASSO 1 — Criar repositório no GitHub

1. Acesse https://github.com/new
2. Nome do repositório: `Gestao-Prazos-IGSA`
3. Visibilidade: **Public** (necessário para Streamlit Cloud gratuito)
4. Clique em **Create repository**

---

## PASSO 2 — Fazer upload dos arquivos

No repositório criado, clique em **uploading an existing file** e envie:

```
app.py
requirements.txt
dados_publicados.json
.streamlit/config.toml
```

Ou via Git no terminal:
```bash
git clone https://github.com/tatiikogan-bip/Gestao-Prazos-IGSA.git
cd Gestao-Prazos-IGSA
# Copie os arquivos aqui
git add .
git commit -m "deploy inicial"
git push
```

---

## PASSO 3 — Deploy no Streamlit Cloud

1. Acesse https://share.streamlit.io
2. Clique em **New app**
3. Preencha:
   - Repository: `tatiikogan-bip/Gestao-Prazos-IGSA`
   - Branch: `main`
   - Main file path: `app.py`
4. Clique em **Deploy**
5. Aguarde ~2 minutos

O link será: `https://gestao-prazos-igsa.streamlit.app`

---

## PASSO 4 — Criar GitHub Token (para publicação automática)

1. Acesse https://github.com/settings/tokens
2. Clique em **Generate new token (classic)**
3. Nome: `Painel Prazos IGSA`
4. Permissões: marque apenas **repo** (full control)
5. Clique em **Generate token**
6. **Copie o token imediatamente** (não será exibido novamente)

Guarde o token — você vai inserir na Área Administrativa ao publicar.

---

## USO DIÁRIO

1. Acesse o link do painel
2. No menu lateral, clique em **⚙️ Área Administrativa**
3. Selecione a data de referência
4. Faça upload da planilha exportada do LegalOne
5. Revise os alertas de inconsistências
6. Clique em **PUBLICAR DADOS NO PAINEL**
7. Os dados ficam imediatamente disponíveis para todos no link público

---

## IMPORTANTE — Persistência dos dados

O Streamlit Cloud **não persiste arquivos entre deploys**.
Para garantir que os dados publicados sobrevivam a reinicializações:

- **Com GitHub Token:** ao publicar, os dados são salvos automaticamente
  no repositório (arquivo `dados_publicados.json`) e ficam permanentes.
- **Sem GitHub Token:** os dados ficam disponíveis apenas enquanto
  o app está rodando. Ao reiniciar, será necessário republicar.

**Recomendação:** sempre use o GitHub Token ao publicar.

---

## ESTRUTURA DE ARQUIVOS

```
Gestao-Prazos-IGSA/
├── app.py                    # Aplicação principal
├── requirements.txt          # Dependências Python
├── dados_publicados.json     # Dados publicados (atualizado pelo admin)
└── .streamlit/
    └── config.toml           # Tema visual IGSA
```
