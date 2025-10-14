"""
Production & Inventory control - Streamlit starter app
Place this repository on GitHub and then deploy on Streamlit Cloud (recommended).
This starter app expects an Excel file with sheets for Materials (MP), Inventory, and Production Orders.
It includes:
- Dashboard (inventory levels, reorder alerts)
- Production Order creation (consumes materials)
- Inventory adjustments (receipts / returns)
- CSV export of current inventory

Adapt the sheet names and columns to your actual spreadsheet (example names: "MP", "Inventory", "Production")
"""

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Produção & Estoque", layout="wide")

st.title("Produção & Controle de Estoque (Starter)")

# --- Load data ---
uploaded = st.sidebar.file_uploader("Upload Excel (ou use arquivo padrão 'Indicadores CPP1.xlsx')", type=["xlsx","xls"])
use_default = False
if uploaded is None:
    try:
        default_path = "Indicadores_CPP1.xlsx"
        df_sheets = pd.read_excel(default_path, sheet_name=None)
        use_default = True
    except Exception:
        df_sheets = {}
else:
    df_sheets = pd.read_excel(uploaded, sheet_name=None)

st.sidebar.markdown("### Sheets encontradas")
for s in df_sheets.keys():
    st.sidebar.write("- " + s)

# Map likely sheet names (customize if needed)
mp_sheet = st.sidebar.selectbox("Sheet: Matérias-primas (MP)", options=list(df_sheets.keys()) + [None], index=0 if len(df_sheets)>0 else 0)
inventory_sheet = st.sidebar.selectbox("Sheet: Estoque", options=list(df_sheets.keys()) + [None], index=min(1,len(df_sheets)-1) if len(df_sheets)>1 else 0)
production_sheet = st.sidebar.selectbox("Sheet: Produção/Ordens", options=list(df_sheets.keys()) + [None], index=min(2,len(df_sheets)-1) if len(df_sheets)>2 else 0)

# Load DataFrames or create empty ones
mp_df = pd.DataFrame()
inv_df = pd.DataFrame()
prod_df = pd.DataFrame()

if mp_sheet in df_sheets:
    mp_df = df_sheets[mp_sheet].copy()
if inventory_sheet in df_sheets:
    inv_df = df_sheets[inventory_sheet].copy()
if production_sheet in df_sheets:
    prod_df = df_sheets[production_sheet].copy()

# Normalise columns if possible
st.sidebar.markdown("### Quick actions")
if st.sidebar.button("Criar snapshot CSV do estoque atual"):
    if not inv_df.empty:
        towrite = BytesIO()
        inv_df.to_csv(towrite, index=False)
        towrite.seek(0)
        st.sidebar.download_button("Download estoque.csv", towrite, file_name="estoque_snapshot.csv", mime="text/csv")
    else:
        st.sidebar.warning("Nenhum dado de estoque disponível.")

# Main UI layout
col1, col2 = st.columns([2,1])

with col1:
    st.header("1. Estoque")
    if inv_df.empty:
        st.info("Nenhum sheet de estoque identificado. Você pode criar manualmente abaixo.")
        inv_df = pd.DataFrame(columns=["mp_id","mp_nome","quantidade","unidade","local"])
    st.dataframe(inv_df)

    st.subheader("Ajuste de estoque (entrada/saída)")
    with st.form("ajuste_form"):
        mp_id = st.selectbox("MP (id/nome)", options=list(inv_df["mp_id"].astype(str).tolist()) if "mp_id" in inv_df.columns else ["--nenhum--"])
        ajuste = st.number_input("Quantidade (positivo entrada, negativo saída)", value=0.0, step=1.0)
        motivo = st.text_input("Motivo")
        submitted = st.form_submit_button("Aplicar ajuste")
        if submitted:
            # Apply change to DataFrame in session state
            if "inv_df" not in st.session_state:
                st.session_state.inv_df = inv_df.copy()
            df = st.session_state.inv_df
            if "mp_id" in df.columns and mp_id != "--nenhum--":
                mask = df["mp_id"].astype(str) == str(mp_id)
                if mask.any():
                    df.loc[mask, "quantidade"] = df.loc[mask, "quantidade"].astype(float) + float(ajuste)
                    st.success("Ajuste aplicado.")
                else:
                    st.error("MP não encontrada.")
            else:
                # create new item if no mp_id present
                new_row = {"mp_id": mp_id, "mp_nome": mp_id, "quantidade": ajuste, "unidade":"", "local":""}
                st.session_state.inv_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                st.success("Item criado e ajuste aplicado.")
            st.experimental_rerun()

with col2:
    st.header("2. Produção")
    st.info("Criar ordem de produção consumirá materiais do estoque (simulado).")
    with st.form("prod_form"):
        prod_code = st.text_input("Código do produto")
        qtd_prod = st.number_input("Quantidade a produzir", value=1, step=1)
        # For simplicity enter BOM as JSON: [{'mp_id':'M001','qty_per_product':0.5}, ...]
        bom_text = st.text_area("BOM (JSON) - lista de componentes com 'mp_id' e 'qty_per_product'", value="[]", height=120)
        submit_prod = st.form_submit_button("Criar ordem e consumir MP")
        if submit_prod:
            import json
            try:
                bom = json.loads(bom_text)
                if "inv_df" not in st.session_state:
                    st.session_state.inv_df = inv_df.copy()
                df = st.session_state.inv_df
                # Check availability
                insufficient = []
                for comp in bom:
                    mid = str(comp.get("mp_id"))
                    need = float(comp.get("qty_per_product",0))*float(qtd_prod)
                    mask = df["mp_id"].astype(str)==mid if "mp_id" in df.columns else pd.Series([False]*len(df))
                    if not mask.any():
                        insufficient.append(f"{mid} (não encontrado)")
                    else:
                        avail = float(df.loc[mask,"quantidade"].sum())
                        if avail < need:
                            insufficient.append(f"{mid} (falta {need-avail})")
                if insufficient:
                    st.error("MP insuficiente: " + "; ".join(insufficient))
                else:
                    # consume
                    for comp in bom:
                        mid = str(comp.get("mp_id"))
                        need = float(comp.get("qty_per_product",0))*float(qtd_prod)
                        mask = df["mp_id"].astype(str)==mid
                        # subtract proportionally from matching rows
                        idxs = df.loc[mask].index
                        remaining = need
                        for ix in idxs:
                            take = min(remaining, float(df.at[ix,"quantidade"]))
                            df.at[ix,"quantidade"] = float(df.at[ix,"quantidade"]) - take
                            remaining -= take
                            if remaining <= 1e-9:
                                break
                    st.success("Ordem criada e MP consumida (simulado).")
                    # record production order in session
                    if "prod_log" not in st.session_state:
                        st.session_state.prod_log = []
                    st.session_state.prod_log.append({"prod_code":prod_code,"qtd":qtd_prod,"bom":bom})
                    st.experimental_rerun()
            except Exception as e:
                st.error("Erro ao ler BOM JSON: " + str(e))

st.sidebar.markdown("---")
st.sidebar.header("Export & Persistência")
if st.sidebar.button("Exportar estoque atual para CSV"):
    df = st.session_state.get("inv_df", inv_df)
    towrite = BytesIO()
    df.to_csv(towrite, index=False)
    towrite.seek(0)
    st.sidebar.download_button("Download CSV", towrite, file_name="estoque_atual.csv", mime="text/csv")

st.sidebar.markdown("**Instruções rápidas**")
st.sidebar.caption("1) Ajuste colunas e nomes das sheets conforme sua planilha.
2) Suba este repositório no GitHub e conecte no Streamlit Cloud para deploy.
3) Para integração completa, substitua armazenamento em session_state por um banco (SQLite / Postgres).")
