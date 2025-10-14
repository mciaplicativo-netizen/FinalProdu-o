"""
Streamlit app - Produção & Estoque (Dashboard)
Features:
- Sidebar navigation: Produção / Estoque MP / Estoque Injetados
- Production dashboard with KPIs, selectable chart (Produção diária / Eficiência média)
- Filters: período, máquina, produto, turno
- SQLite <-> Excel synchronization (reads from DB if exists; can import from Excel; writes update both DB and Excel)
Notes:
- Keep backups of the Excel file before using write features.
"""

import streamlit as st
import pandas as pd
import sqlite3
from pathlib import Path
import time, json, datetime
from io import BytesIO

APP_DIR = Path(".")
DB_PATH = APP_DIR / "database.db"
XLSX_PATH = APP_DIR / "Indicadores_CPP1.xlsx"
LOCK_PATH = APP_DIR / ".write_lock"

st.set_page_config(page_title="Produção & Estoque - Dashboard", layout="wide")

# -------------------- Helpers: locking + excel/db sync --------------------
def with_lock(fn):
    def wrapper(*args, **kwargs):
        # simple file lock
        while LOCK_PATH.exists():
            time.sleep(0.05)
        try:
            LOCK_PATH.write_text("lock")
            return fn(*args, **kwargs)
        finally:
            if LOCK_PATH.exists():
                LOCK_PATH.unlink()
    return wrapper

@with_lock
def write_excel_sheets(sheet_dfs: dict):
    # writes/overwrites the given sheets into the Excel file, preserving others
    import pandas as pd
    from openpyxl import load_workbook
    if not XLSX_PATH.exists():
        with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
            for name, df in sheet_dfs.items():
                df.to_excel(writer, sheet_name=name, index=False)
        return
    wb = load_workbook(XLSX_PATH)
    # remove sheets to overwrite
    for name in list(sheet_dfs.keys()):
        if name in wb.sheetnames:
            std = wb[name]
            wb.remove(std)
    wb.save(XLSX_PATH)
    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl", mode="a") as writer:
        for name, df in sheet_dfs.items():
            df.to_excel(writer, sheet_name=name, index=False)

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

def init_db_from_excel(mapping):
    # mapping: sheet_name -> table_name
    conn = sqlite3.connect(DB_PATH)
    for sheet, table in mapping.items():
        try:
            df = pd.read_excel(XLSX_PATH, sheet_name=sheet)
        except Exception:
            continue
        df.to_sql(table, conn, if_exists="replace", index=False)
    conn.close()

# -------------------- Detect sheets and mapping --------------------
detected = []
if XLSX_PATH.exists():
    try:
        x = pd.ExcelFile(XLSX_PATH)
        detected = x.sheet_names
    except Exception:
        detected = []

default_mapping = {
    "Estoque MP": "estoque_mp",
    "Estoque Injetados": "estoque_injetados",
    "Produção - injeção+ Zamac": "producao"
}

# Sidebar: menu and mapping adjustments
st.sidebar.title("Navegação")
menu = st.sidebar.radio("Ir para", ["Produção", "Estoque MP", "Estoque Injetados", "Admin"])

st.sidebar.markdown("---")
st.sidebar.header("Mapeamento de sheets")
user_mapping = {}
for sx, tbl in default_mapping.items():
    val = sx if sx in detected else sx
    newname = st.sidebar.text_input(f"Sheet fonte para '{tbl}'", value=val, key="map_"+tbl)
    user_mapping[newname] = tbl

st.sidebar.markdown("---")
st.sidebar.header("Import / Export")
uploaded = st.sidebar.file_uploader("Substituir arquivo Excel (opcional)", type=["xlsx","xls"])
if uploaded is not None:
    with open(XLSX_PATH, "wb") as f:
        f.write(uploaded.getbuffer())
    st.sidebar.success("Arquivo Excel substituído. Recarregue para aplicar.")

if st.sidebar.button("Sincronizar Excel → DB (forçar)"):
    init_db_from_excel({k:v for k,v in user_mapping.items()})
    st.sidebar.success("Sincronização inicial (Excel→DB) feita. Recarregue a página.")

# Ensure DB exists (import if necessary)
if not DB_PATH.exists() and XLSX_PATH.exists():
    init_db_from_excel({k:v for k,v in user_mapping.items()})

# Load tables
tables = {}
for sheet_name, table_name in user_mapping.items():
    tables[table_name] = read_table(table_name)

# Ensure expected columns for production table
prod_df = tables.get("producao", pd.DataFrame())
# attempt to coerce date column
if "Data" in prod_df.columns:
    try:
        prod_df["Data"] = pd.to_datetime(prod_df["Data"], errors="coerce").dt.date
    except Exception:
        pass

# -------------------- Produção page --------------------
if menu == "Produção":
    st.title("Produção — Dashboard")
    st.markdown("Resumo rápido e gráficos interactivos extraídos da sheet de produção.")

    # Filters column and main column
    sidebar_filters, main = st.columns([1,3], gap="large")

    with sidebar_filters:
        st.subheader("Filtros")
        min_date = None
        max_date = None
        if "Data" in prod_df.columns:
            min_date = pd.to_datetime(prod_df["Data"]).min().date() if not prod_df["Data"].isna().all() else None
            max_date = pd.to_datetime(prod_df["Data"]).max().date() if not prod_df["Data"].isna().all() else None
            date_range = st.date_input("Período", value=(min_date, max_date) if min_date and max_date else None)
        else:
            date_range = None

        machines = sorted(prod_df["Máquina"].dropna().unique().tolist()) if "Máquina" in prod_df.columns else []
        sel_machine = st.multiselect("Máquina", options=machines, default=machines)

        products = sorted(prod_df["Produto"].dropna().unique().tolist()) if "Produto" in prod_df.columns else []
        sel_product = st.multiselect("Produto", options=products, default=products)

        turns = sorted(prod_df["Turno"].dropna().unique().tolist()) if "Turno" in prod_df.columns else []
        sel_turn = st.multiselect("Turno", options=turns, default=turns)

        chart_type = st.selectbox("Tipo de métrica", ["Produção diária (Realizado)", "Eficiência média (por dia)"])

    with main:
        # Apply filters
        df = prod_df.copy()
        if df.empty:
            st.info("Nenhum dado de produção disponível. Importe do Excel no sidebar ou sincronize.")
        else:
            if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
                start, end = date_range
                df = df[(pd.to_datetime(df["Data"]) >= pd.to_datetime(start)) & (pd.to_datetime(df["Data"]) <= pd.to_datetime(end))]
            if "Máquina" in df.columns and sel_machine:
                df = df[df["Máquina"].isin(sel_machine)]
            if "Produto" in df.columns and sel_product:
                df = df[df["Produto"].isin(sel_product)]
            if "Turno" in df.columns and sel_turn:
                df = df[df["Turno"].isin(sel_turn)]

            # KPIs
            col1, col2, col3, col4 = st.columns(4)
            total_prod = int(df["Realizado"].sum()) if "Realizado" in df.columns else 0
            eficiencia_mean = float(df["Eficiência"].mean()) if "Eficiência" in df.columns and not df["Eficiência"].isna().all() else None
            total_ciclos = int(df["Ciclos"].sum()) if "Ciclos" in df.columns else None
            aparas = float(df["Kg Aparas"].sum()) if "Kg Aparas" in df.columns else None
            with col1:
                st.metric("Total produzido", f"{total_prod:,d}")
            with col2:
                st.metric("Eficiência média", f"{eficiencia_mean:.3f}" if eficiencia_mean is not None else "—")
            with col3:
                st.metric("Total de ciclos", f"{total_ciclos:,d}" if total_ciclos is not None else "—")
            with col4:
                st.metric("Kg Aparas", f"{aparas:.2f}" if aparas is not None else "—")

            # Time series chart
            # aggregate by date
            if "Data" in df.columns:
                agg = None
                if chart_type.startswith("Produção"):
                    agg = df.groupby("Data", as_index=False)["Realizado"].sum().sort_values("Data")
                    st.subheader("Produção por dia")
                    st.line_chart(agg.set_index("Data")["Realizado"])
                else:
                    agg = df.groupby("Data", as_index=False)["Eficiência"].mean().sort_values("Data")
                    st.subheader("Eficiência média por dia")
                    st.line_chart(agg.set_index("Data")["Eficiência"])

            # Show top machines/products
            st.subheader("Detalhes rápidas")
            if "Máquina" in df.columns and "Realizado" in df.columns:
                top_mach = df.groupby("Máquina", as_index=False)["Realizado"].sum().sort_values("Realizado", ascending=False).head(5)
                st.table(top_mach)
            if "Produto" in df.columns and "Realizado" in df.columns:
                top_prod = df.groupby("Produto", as_index=False)["Realizado"].sum().sort_values("Realizado", ascending=False).head(8)
                st.table(top_prod)

            # Editable/filterable table for inspection
            st.subheader("Tabela filtrada (editar se necessário)")
            edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
            if st.button("Salvar alterações na produção"):
                # save to DB and Excel (preserve original column order)
                write_table("producao", edited)
                write_excel_sheets({list(user_mapping.keys())[list(user_mapping.values()).index("producao")]: edited})
                st.success("Produção salva no DB e Excel.")

# -------------------- Estoque MP page --------------------
elif menu == "Estoque MP":
    st.title("Estoque MP")
    df_mp = tables.get("estoque_mp", pd.DataFrame(columns=["mp_id","mp_nome","quantidade","unidade","local"]))
    st.write("Edite o estoque MP abaixo:")
    edited_mp = st.data_editor(df_mp, num_rows="dynamic", use_container_width=True)
    if st.button("Salvar Estoque MP"):
        write_table("estoque_mp", edited_mp)
        write_excel_sheets({list(user_mapping.keys())[list(user_mapping.values()).index("estoque_mp")]: edited_mp})
        st.success("Estoque MP salvo no DB e Excel.")

# -------------------- Estoque Injetados page --------------------
elif menu == "Estoque Injetados":
    st.title("Estoque Injetados")
    df_inj = tables.get("estoque_injetados", pd.DataFrame(columns=["sku","nome","quantidade","unidade","local"]))
    st.write("Edite o estoque de peças injetadas abaixo:")
    edited_inj = st.data_editor(df_inj, num_rows="dynamic", use_container_width=True)
    if st.button("Salvar Estoque Injetados"):
        write_table("estoque_injetados", edited_inj)
        write_excel_sheets({list(user_mapping.keys())[list(user_mapping.values()).index("estoque_injetados")]: edited_inj})
        st.success("Estoque Injetados salvo no DB e Excel.")

# -------------------- Admin --------------------
else:
    st.title("Admin")
    st.write("DB:", DB_PATH)
    st.write("Excel:", XLSX_PATH)
    if st.button("Forçar reescrita DB → Excel (todas)"):
        d1 = read_table("estoque_mp")
        d2 = read_table("estoque_injetados")
        d3 = read_table("producao")
        towrite = {}
        if not d1.empty:
            towrite[list(user_mapping.keys())[list(user_mapping.values()).index("estoque_mp")]] = d1
        if not d2.empty:
            towrite[list(user_mapping.keys())[list(user_mapping.values()).index("estoque_injetados")]] = d2
        if not d3.empty:
            towrite[list(user_mapping.keys())[list(user_mapping.values()).index("producao")]] = d3
        if towrite:
            write_excel_sheets(towrite)
            st.success("Todas sheets regravadas no Excel.")
    if st.button("Criar backup do Excel"):
        dst = XLSX_PATH.parent / f"backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2(XLSX_PATH, dst)
        st.success(f"Backup criado: {dst.name}")

st.sidebar.caption("""Use este app com cuidado: o recurso de gravação reescreve sheets no seu arquivo Excel. Faça backup antes de usar.""")
