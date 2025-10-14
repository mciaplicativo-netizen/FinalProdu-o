# Produção & Controle de Estoque - Starter Streamlit App

Este repositório é um **template** para um app Streamlit que gerencia produção e controle de estoque (matérias-primas).

## Como usar
1. Coloque o arquivo `Indicadores_CPP1.xlsx` (ou faça upload via sidebar) na raiz do repo.
2. Personalize os nomes das sheets e colunas no `streamlit_app.py`.
3. Faça commit no GitHub.
4. No Streamlit Cloud -> New app -> conecte ao repositório e escolha `streamlit_app.py`.

## Funcionalidades
- Visualizar estoque
- Ajustar entradas/saídas de estoque
- Criar ordens de produção que consomem MP via BOM em JSON
- Exportar CSV do estoque atual

## Próximos passos recomendados
- Persistir dados em banco (SQLite ou Postgres)
- Criar autenticação (OAuth)
- Adicionar notificações de reorder (e-mails)
