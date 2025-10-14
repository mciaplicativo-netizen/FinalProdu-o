"""
Streamlit app - Produção & Estoque (sincronizado com SQLite e Excel)
- Ao iniciar: importa sheets do Excel para SQLite (se existirem)
- Atualizações via app gravam no SQLite e reaplicam as sheets no Excel
- Mapeamento automático de sheets detectadas:
    * "Estoque MP" -> estoque_mp
    * "Estoque Injetados" -> estoque_injetados
    * "Produção - injeção+ Zamac" -> producao
"""

import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
import threading
import time
import json
from io import BytesIO

APP_DIR = Path(".")
DB_PATH = APP_DIR / "database.db"
XLSX_PATH = APP_DIR / "Indicadores_CPP1.xlsx"
LOCK_PATH = APP_DIR / ".write_lock"

st.set_page_config(page_title="Produção & Estoque (Sincronizado)", layout="wide")
st.title("Produção & Estoque — SQLite ⇄ Excel (sincronizado)")

def with_lock(fn):
    def wrapper(*args, **kwargs):
        # Simple file lock
        while LOCK_PATH.exists():
            time.sleep(0.1)
        try:
            LOCK_PATH.write_text("lock")
            return fn(*args, **kwargs)
        finally:
            if LOCK_PATH.exists():
                LOCK_PATH.unlink()
    return wrapper

@with_lock
def write_excel_sheets(sheet_dfs: dict):
    """
    Escreve (ou substitui) as sheets fornecidas no arquivo Excel, preservando outras sheets.
    sheet_dfs: dict of sheet_name -> DataFrame
    """
    from openpyxl import load_workbook
    # If file doesn't exist, create new with the sheets
    if not XLSX_PATH.exists():
        with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
            for name, df in sheet_dfs.items():
                df.to_excel(writer, sheet_name=name, index=False)
        return
    # Load existing workbook
    wb = load_workbook(XLSX_PATH)
    # Remove sheets that we will overwrite if they exist
    for name in list(sheet_dfs.keys()):
        if name in wb.sheetnames:
            std = wb[name]
            wb.remove(std)
    # Save workbook temporarily then append sheets via pandas
    wb.save(XLSX_PATH)
    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl", mode="a") as writer:
        for name, df in sheet_dfs.items():
            df.to_excel(writer, sheet_name=name, index=False)

def read_excel_sheets(names):
    """Lê as sheets especificadas do Excel se existirem; retorna dict name->df"""
    out = {}
    if not XLSX_PATH.exists():
        return out
    try:
        xls = pd.ExcelFile(XLSX_PATH)
    except Exception as e:
        st.error(f"Erro ao abrir Excel: {e}")
        return out
    for name in names:
        if name in xls.sheet_names:
            out[name] = pd.read_excel(XLSX_PATH, sheet_name=name)
    return out

def init_db_from_excel(mapping):
    """Carrega dados do excel para o DB (apenas se tabelas vazias)"""
    conn = sqlite3.connect(DB_PATH)
    for sheet, table in mapping.items():
        try:
            df = pd.read_excel(XLSX_PATH, sheet_name=sheet)
        except Exception:
            continue
        # write/replace table
        df.to_sql(table, conn, if_exists="replace", index=False)
    conn.close()

def read_table(table):
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql_query(f"SELECT * FROM [{table}]", conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df

def write_table(table, df):
    conn = sqlite3.connect(DB_PATH)
    df.to_sql(table, conn, if_exists="replace", index=False)
    conn.close()

# Auto-detect sheets
detected = []
if XLSX_PATH.exists():
    try:
        x = pd.ExcelFile(XLSX_PATH)
        detected = x.sheet_names
    except Exception:
        detected = []

# Suggested mapping (customize if your sheet names differ)
suggested = {
    "Estoque MP": "estoque_mp",
    "Estoque Injetados": "estoque_injetados",
    "Produção - injeção+ Zamac": "producao"
}

# Allow user to adjust mapping
st.sidebar.header("Mapeamento de sheets")
user_mapping = {}
for sx, tbl in suggested.items():
    if sx in detected:
        newname = st.sidebar.text_input(f"Sheet fonte para '{tbl}'", value=sx, key=sx)
        user_mapping[newname] = tbl

st.sidebar.markdown("---")
st.sidebar.header("Operações")
uploaded = st.sidebar.file_uploader("Substituir arquivo Excel (opcional)", type=["xlsx","xls"])
if uploaded is not None:
    # overwrite local copy
    with open(XLSX_PATH, "wb") as f:
        f.write(uploaded.getbuffer())
    st.sidebar.success("Arquivo Excel substituído. Recarregue a página se necessário.")

if st.sidebar.button("Forçar sincronização inicial (Excel -> DB)"):
    init_db_from_excel({k:v for k,v in user_mapping.items()})
    st.sidebar.success("Sincronização inicial realizada.")

# On start: ensure DB exists and load initial data
if not DB_PATH.exists() and XLSX_PATH.exists():
    init_db_from_excel({k:v for k,v in user_mapping.items()})

# Load current tables into dataframes
dfs = {}
for sheet_name, table_name in user_mapping.items():
    dfs[table_name] = read_table(table_name)

# Show tabs: Estoque MP, Estoque Injetados, Produção
tabs = st.tabs(["Estoque MP", "Estoque Injetados", "Produção", "Admin"])
# Estoque MP
with tabs[0]:
    st.header("Estoque MP")
    df_mp = dfs.get("estoque_mp", pd.DataFrame(columns=["mp_id","mp_nome","quantidade","unidade","local"]))
    edited = st.experimental_data_editor(df_mp, num_rows="dynamic")
    if st.button("Salvar Estoque MP"):
        write_table("estoque_mp", edited)
        # also write to excel
        write_excel_sheets({"Estoque MP": edited})
        st.success("Estoque MP salvo no DB e Excel.")

# Estoque Injetados
with tabs[1]:
    st.header("Estoque Injetados")
    df_inj = dfs.get("estoque_injetados", pd.DataFrame(columns=["sku","nome","quantidade","unidade","local"]))
    edited2 = st.experimental_data_editor(df_inj, num_rows="dynamic")
    if st.button("Salvar Estoque Injetados"):
        write_table("estoque_injetados", edited2)
        write_excel_sheets({"Estoque Injetados": edited2})
        st.success("Estoque Injetados salvo no DB e Excel.")

# Produção
with tabs[2]:
    st.header("Produção / Ordens")
    df_prod = dfs.get("producao", pd.DataFrame(columns=["id","prod_code","qtd","data","bom_json"]))
    st.subheader("Log de Ordens")
    st.dataframe(df_prod)
    st.subheader("Criar nova ordem de produção (consome MP)")
    with st.form("new_prod"):
        prod_code = st.text_input("Código do produto")
        qtd = st.number_input("Quantidade", value=1, step=1)
        bom = st.text_area("BOM JSON (ex: [{'mp_id':'M001','qty_per_product':0.5}])", value="[]", height=120)
        submit = st.form_submit_button("Criar ordem")
        if submit:
            try:
                bom_list = json.loads(bom)
            except Exception as e:
                st.error("BOM inválido: " + str(e))
                bom_list = []
            # read current estoque_mp
            estoque_mp = read_table("estoque_mp")
            if estoque_mp.empty:
                st.error("Estoque MP vazio. Impossível consumir.")
            else:
                # verify availability
                insufficient = []
                for comp in bom_list:
                    mid = str(comp.get("mp_id"))
                    need = float(comp.get("qty_per_product",0)) * float(qtd)
                    mask = estoque_mp["mp_id"].astype(str) == mid if "mp_id" in estoque_mp.columns else pd.Series([False]*len(estoque_mp))
                    if not mask.any():
                        insufficient.append(f"{mid} (não encontrado)")
                    else:
                        avail = float(estoque_mp.loc[mask,"quantidade"].sum())
                        if avail < need - 1e-9:
                            insufficient.append(f"{mid} (falta {need-avail:.3f})")
                if insufficient:
                    st.error("MP insuficiente: " + "; ".join(insufficient))
                else:
                    # consume from estoque_mp (proporcional)
                    for comp in bom_list:
                        mid = str(comp.get("mp_id"))
                        need = float(comp.get("qty_per_product",0)) * float(qtd)
                        mask = estoque_mp["mp_id"].astype(str) == mid
                        idxs = estoque_mp.loc[mask].index
                        remaining = need
                        for ix in idxs:
                            take = min(remaining, float(estoque_mp.at[ix,"quantidade"]))
                            estoque_mp.at[ix,"quantidade"] = float(estoque_mp.at[ix,"quantidade"]) - take
                            remaining -= take
                            if remaining <= 1e-9:
                                break
                    # write back estoque_mp and record production
                    write_table("estoque_mp", estoque_mp)
                    # append to producao table
                    prod_table = read_table("producao")
                    import datetime
                    new_id = int(prod_table["id"].max())+1 if (not prod_table.empty and "id" in prod_table.columns) else 1
                    bom_json = json.dumps(bom_list, ensure_ascii=False)
                    new_row = {"id": new_id, "prod_code": prod_code, "qtd": int(qtd), "data": datetime.datetime.now().isoformat(), "bom_json": bom_json}
                    prod_table = pd.concat([prod_table, pd.DataFrame([new_row])], ignore_index=True)
                    write_table("producao", prod_table)
                    # persist both tables to Excel
                    write_excel_sheets({"Estoque MP": estoque_mp, "Produção - injeção+ Zamac": prod_table})
                    st.success("Ordem criada, MP consumida, DB e Excel atualizados.")

with tabs[3]:
    st.header("Admin")
    st.write("Banco:", DB_PATH)
    st.write("Excel:", XLSX_PATH)
    if st.button("Forçar reescrita de todas sheets do DB para Excel"):
        # Read all three tables and write to excel
        d1 = read_table("estoque_mp")
        d2 = read_table("estoque_injetados")
        d3 = read_table("producao")
        towrite = {}
        if not d1.empty:
            towrite["Estoque MP"] = d1
        if not d2.empty:
            towrite["Estoque Injetados"] = d2
        if not d3.empty:
            towrite["Produção - injeção+ Zamac"] = d3
        if towrite:
            write_excel_sheets(towrite)
            st.success("Sheets regravadas no Excel.")
        else:
            st.info("Nenhuma tabela para gravar.")
    st.markdown("### Backup")
    if st.button("Criar backup do Excel (copy)"):
        import shutil, datetime
        dst = XLSX_PATH.parent / f"backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2(XLSX_PATH, dst)
        st.success(f"Backup criado: {dst.name}")

st.sidebar.caption("""1) Ajuste nomes das sheets no mapeamento se necessário.\n2) Este app regrava as sheets indicadas no arquivo Excel — mantenha backup.\n3) Em ambientes concorrentes, evite editar o Excel manualmente enquanto o app salva.""")
