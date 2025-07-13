import streamlit as st
import pandas as pd
import io
from pathlib import Path

# Try both locations (with fallback)
EXCEL_PATH = Path("data/BIM_BOKS_V1.xlsx")
if not EXCEL_PATH.exists():
    EXCEL_PATH = Path("BIM_BOKS_V1.xlsx")

@st.cache_data
def load_data(path):
    xls = pd.ExcelFile(path)
    df_boks = pd.read_excel(xls, sheet_name="BIM_boks")
    df_cursussen = pd.read_excel(xls, sheet_name="CursusCode")
    df_categorieen = pd.read_excel(xls, sheet_name="CategorieDefinities")
    df_onderwerpen = pd.read_excel(xls, sheet_name="OnderwerpDefinities")

    return {
        "boks": df_boks,
        "cursussen": df_cursussen.rename(columns=lambda c: c.strip()),
        "categorieen": df_categorieen.rename(columns=lambda c: c.strip()),
        "onderwerpen": df_onderwerpen.rename(columns=lambda c: c.strip()),
    }

st.title("📘 BIM BOKS Beheer App")

try:
    data = load_data(EXCEL_PATH)
except Exception as e:
    st.error(f"Kan Excel-bestand niet laden: {e}")
    st.stop()

if "df_bim_boks" not in st.session_state:
    st.session_state.df_bim_boks = data["boks"].copy()

with st.expander("🔍 Filter"):
    col1, col2, col3 = st.columns(3)
    f1 = col1.selectbox("Cursus", ["Alle"] + data["cursussen"]["CursusCode"].dropna().unique().tolist())
    f2 = col2.selectbox("Categorie", ["Alle"] + data["categorieen"]["BOKS categorie"].dropna().unique().tolist())
    f3 = col3.selectbox("Onderwerp", ["Alle"] + data["onderwerpen"]["BOKS onderwerp"].dropna().unique().tolist())

df_view = st.session_state.df_bim_boks.copy()
if f1 != "Alle":
    df_view = df_view[df_view["CursusCode"] == f1]
if f2 != "Alle":
    df_view = df_view[df_view["BOKS categorie"] == f2]
if f3 != "Alle":
    df_view = df_view[df_view["BOKS onderwerp"] == f3]

st.subheader("📄 Overzicht")
st.dataframe(df_view, use_container_width=True)

st.subheader("➕ Nieuwe regel")

with st.form("invoer"):
    c_code = st.selectbox("Cursuscode", data["cursussen"]["CursusCode"].dropna().unique())
    cat = st.selectbox("Categorie", data["categorieen"]["BOKS categorie"].dropna().unique())
    onderwerpen = data["onderwerpen"]
    opties = onderwerpen[onderwerpen["BOKS categorie"] == cat]["BOKS onderwerp"].dropna().unique()
    onder = st.selectbox("Onderwerp", opties)
    ok = st.form_submit_button("Toevoegen")

    if ok:
        combi_ok = ((onderwerpen["BOKS categorie"] == cat) & (onderwerpen["BOKS onderwerp"] == onder)).any()
        if not combi_ok:
            st.error("❌ Combinatie ongeldig")
        else:
            st.session_state.df_bim_boks = pd.concat([
                st.session_state.df_bim_boks,
                pd.DataFrame([{
                    "CursusCode": c_code,
                    "BOKS categorie": cat,
                    "BOKS onderwerp": onder
                }])
            ], ignore_index=True)
            st.success("✅ Toegevoegd!")

st.subheader("⬇️ Download")

def make_download(df, fmt):
    buffer = io.BytesIO()
    if fmt == "csv":
        return df.to_csv(index=False).encode("utf-8")
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="BIM_boks", index=False)
    return buffer.getvalue()

col1, col2 = st.columns(2)
col1.download_button("Download CSV", make_download(st.session_state.df_bim_boks, "csv"),
                     file_name="BIM_boks.csv", mime="text/csv")
col2.download_button("Download Excel", make_download(st.session_state.df_bim_boks, "xlsx"),
                     file_name="BIM_boks.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
