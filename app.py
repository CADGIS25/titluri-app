import streamlit as st
import pandas as pd
import pyodbc
import io

st.set_page_config(page_title="Titluri și Proprietari", layout="wide")
st.title("📄 Generator Titluri - Proprietari - Parcelele")

def read_excel(file):
    xls = pd.ExcelFile(file)
    data = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    return data

def read_mdb(file):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=" + file.name + ";"
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    tables = [t.table_name for t in cursor.tables(tableType='TABLE')]
    data = {table: pd.read_sql(f"SELECT * FROM [{table}]", conn) for table in tables}
    conn.close()
    return data

def process_data(data):
    try:
        owner = data["Owner"]
        titles = data["Titluri_L18"]
        townership = data["townership"]
        parcel = data["Parcel"]
    except KeyError:
        st.error("Fișierul nu conține toate tabelele necesare.")
        return None, None

    # Pregătim chei comune
    parcel["parcel_dno_str"] = parcel["parcel_dno"].astype(str)
    titles["nr_titlu_str"] = titles["nr_titlu"].astype(str)

    titles_with_parcels = titles.merge(parcel, left_on="nr_titlu_str", right_on="parcel_dno_str", how="left")
    town_with_owner = townership.merge(owner, left_on="townership_owner_id", right_on="owner_id", how="left")
    full_table = town_with_owner.merge(titles_with_parcels, left_on="townership_nr_titlu", right_on="nr_titlu", how="left")

    compact_cols = [
        "no_titlu", "data_titlu", "pdf_titlu",
        "owner_lastname", "owner_firstname",
        "aria_reconst", "aria_constit",
        "parcel_larea", "parcel_tno", "parcel_pno"
    ]
    compact_table = full_table[compact_cols].drop_duplicates()

    return full_table, compact_table

# Upload
uploaded_file = st.file_uploader("Încarcă fișierul Excel (.xlsx) sau Access (.mdb)", type=["xlsx", "mdb", "accdb"])

if uploaded_file:
    filetype = uploaded_file.name.split(".")[-1].lower()
    
    if filetype in ["mdb", "accdb"]:
        st.error("⚠️ Fișierele .mdb nu sunt suportate pe Streamlit Cloud. Te rugăm să folosești .xlsx.")
        st.stop()
        
    if filetype == "xlsx":
        data = read_excel(uploaded_file)
    elif filetype in ["mdb", "accdb"]:
        with open(f"/tmp/{uploaded_file.name}", "wb") as f:
            f.write(uploaded_file.read())
        data = read_mdb(f"/tmp/{uploaded_file.name}")
    else:
        st.warning("Format de fișier neacceptat.")
        data = None

    if data:
        full_table, compact_table = process_data(data)
        if full_table is not None:
            st.success("Datele au fost procesate cu succes.")
            st.subheader("🔍 Tabel Complet")
            st.dataframe(full_table.head(50))
            st.subheader("📄 Tabel Compact")
            st.dataframe(compact_table.head(50))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                full_table.to_excel(writer, sheet_name="Tabel_Complet", index=False)
                compact_table.to_excel(writer, sheet_name="Format_Compact", index=False)
            st.download_button("📥 Descarcă Excel", data=output.getvalue(), file_name="Tabel_Titluri_Export.xlsx")



