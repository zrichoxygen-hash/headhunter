import os
import io
from typing import Optional

import pandas as pd
import streamlit as st

from serpapi_scraper import process_rows

st.set_page_config(page_title="Scraper UI", layout="wide")
st.title("Scraper d'exploration de profils")
st.caption("Saisissez les recherches (company, role, ville), lancez la collecte, puis visualisez/exportez les résultats.")

if "results_df" not in st.session_state:
    st.session_state.results_df = pd.DataFrame()

DEFAULT_INPUT = pd.DataFrame(
    [
        {"company": "Consort", "role": "Data Engineer", "skills": "Python; SQL; Azure", "experience": "Data pipelines; ETL; Databricks", "sites_cibles": "linkedin.com/in; ictjob.com", "ville": "Bruxelles", "pays": "Belgique"},
        {"company": "Consort", "role": "Data Scientist", "skills": "Python; Machine Learning; SQL", "experience": "Model training; analytics; forecasting", "sites_cibles": "linkedin.com/in; ictjob.com", "ville": "Bruxelles", "pays": "Belgique"},
        {"company": "Consort", "role": "DevOps Engineer", "skills": "AWS; Kubernetes; Docker", "experience": "CI/CD; cloud infrastructure", "sites_cibles": "linkedin.com/in; ictjob.com", "ville": "Bruxelles", "pays": "Belgique"},
    ]
)

with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input("Clé SerpApi", value=os.getenv("SERPAPI_API_KEY", ""), type="password")
    max_results = st.slider("Résultats par requête (0 à 1000)", min_value=0, max_value=1000, value=3)
    delay = st.slider("Pause entre requêtes (s)", min_value=0.0, max_value=2.0, value=0.5, step=0.1)

    uploaded_file = st.file_uploader("Importer un CSV de recherches", type=["csv"])
    if uploaded_file is not None:
        try:
            uploaded_df = pd.read_csv(uploaded_file)
            expected_cols = ["company", "role", "skills", "experience", "sites_cibles", "ville", "pays"]
            normalized = {}
            for col in expected_cols:
                if col in uploaded_df.columns:
                    normalized[col] = uploaded_df[col]
                else:
                    normalized[col] = ""
            input_df = pd.DataFrame(normalized)
            st.session_state["input_df"] = input_df[expected_cols].copy()
            st.success("CSV importé")
        except Exception as exc:
            st.error(f"Impossible de lire le CSV: {exc}")

    run_button = st.button("Lancer le scraping", use_container_width=True)
    clear_button = st.button("Effacer les résultats", use_container_width=True)
    if clear_button:
        st.session_state.results_df = pd.DataFrame()

if "input_df" not in st.session_state:
    st.session_state.input_df = DEFAULT_INPUT

expected_cols = ["company", "role", "skills", "experience", "sites_cibles", "ville", "pays"]
st.session_state.input_df = st.session_state.input_df.reindex(columns=expected_cols, fill_value="")

st.subheader("Recherches à exécuter")
input_df = st.data_editor(
    st.session_state.input_df,
    use_container_width=True,
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "company": st.column_config.TextColumn("Company"),
        "role": st.column_config.TextColumn("Role"),
        "skills": st.column_config.TextColumn("Skills"),
        "experience": st.column_config.TextColumn("Experience"),
        "sites_cibles": st.column_config.TextColumn("Sites cibles"),
        "ville": st.column_config.TextColumn("Ville"),
        "pays": st.column_config.TextColumn("Pays"),
    },
)
st.session_state.input_df = input_df.reindex(columns=expected_cols, fill_value="")

if run_button:
    if not api_key:
        st.error("Une clé SerpApi est requise.")
    elif max_results < 1:
        st.error("Le nombre de résultats doit être supérieur à 0.")
    elif input_df.empty:
        st.warning("Ajoutez au moins une ligne de recherche.")
    else:
        with st.spinner("Exécution du scraping en cours…"):
            results_df = process_rows(
                input_df.to_dict(orient="records"),
                api_key,
                output_xlsx=None,
                max_results=max_results,
                delay=delay,
            )
        st.session_state.results_df = results_df

st.subheader("Résultats")
if st.session_state.results_df.empty:
    st.info("Aucun résultat pour le moment. Lancez un scraping pour voir les données ici.")
else:
    results_df = st.session_state.results_df.copy()
    if "status" not in results_df.columns:
        results_df["status"] = "new"
    results_df["status"] = results_df["status"].fillna("new")

    status_options = ["new", "to_review", "validated", "blocked"]
    status_options = sorted(set(status_options + list(results_df["status"].astype(str).unique())))

    cols = st.columns(4)
    with cols[0]:
        company_filter = st.multiselect("Company", options=sorted({str(x) for x in results_df.get("company", pd.Series([""])).fillna("").astype(str).unique()}), default=[])
    with cols[1]:
        role_filter = st.multiselect("Role", options=sorted({str(x) for x in results_df.get("role", pd.Series([""])).fillna("").astype(str).unique()}), default=[])
    with cols[2]:
        city_filter = st.multiselect("Ville", options=sorted({str(x) for x in results_df.get("city", pd.Series([""])).fillna("").astype(str).unique()}), default=[])
    with cols[3]:
        status_filter = st.multiselect("Status", options=status_options, default=status_options)

    filtered_df = results_df.copy()
    if company_filter:
        filtered_df = filtered_df[filtered_df.get("company", "").fillna("").astype(str).isin(company_filter)]
    if role_filter:
        filtered_df = filtered_df[filtered_df.get("role", "").fillna("").astype(str).isin(role_filter)]
    if city_filter:
        filtered_df = filtered_df[filtered_df.get("city", "").fillna("").astype(str).isin(city_filter)]
    if status_filter:
        filtered_df = filtered_df[filtered_df.get("status", "new").fillna("new").astype(str).isin(status_filter)]

    st.dataframe(filtered_df, use_container_width=True, height=500)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="results")
    buffer.seek(0)
    st.download_button(
        label="Exporter vers Excel",
        data=buffer,
        file_name="results.xlsx",
        mime="application/vnd.openxmlformats-officedocument/spreadsheetml.sheet",
        use_container_width=True,
    )

    csv_download = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Exporter vers CSV",
        data=csv_download,
        file_name="results.csv",
        mime="text/csv",
        use_container_width=True,
    )
