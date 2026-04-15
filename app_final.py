import streamlit as st
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
import unicodedata

st.set_page_config(page_title="Sistema de Faltas", layout="wide")

# ======================
# NORMALIZAR
# ======================

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8")
    texto = " ".join(texto.split())
    return texto

# ======================
# HEADER
# ======================

st.title("📊 SISTEMA DE APURAÇÃO DE FALTAS")
st.markdown("---")

# ======================
# UPLOAD
# ======================

arquivo = st.file_uploader("📂 Envie a planilha Excel", type=["xlsx"])

if arquivo:

    df = pd.read_excel(arquivo)
    df.columns = df.columns.str.strip().str.upper()

    # ======================
    # TRATAMENTO
    # ======================

    df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce")

    df["MES"] = df["DATA"].dt.strftime("%b/%Y").str.upper()

    df = df[df["EVENTO_DECRICAO"].astype(str).str.contains("FALTA", na=False)]

    df["EQUIPE"] = df["EQUIPE"].apply(normalizar_texto)

    if df.empty:
        st.error("❌ Nenhuma falta encontrada")
        st.stop()

    # ======================
    # KPIs
    # ======================

    col1, col2, col3 = st.columns(3)

    col1.metric("Total de Faltas", len(df))
    col2.metric("Gestores", df["EQUIPE"].nunique())
    col3.metric("Setores", df["SECAO_NOME"].nunique())

    st.markdown("---")

    # ======================
    # TOP 5 GESTORES
    # ======================

    st.subheader("🏆 Top 5 Equipe de Gestores com Mais Faltas")

    ranking = df.groupby(["EQUIPE", "MES"]).size().reset_index(name="TOTAL")

    pivot = ranking.pivot(index="EQUIPE", columns="MES", values="TOTAL").fillna(0)

    pivot["TOTAL"] = pivot.sum(axis=1)

    top5 = pivot.sort_values("TOTAL", ascending=False).head(5)

    st.dataframe(top5)

    st.bar_chart(top5["TOTAL"])

    st.markdown("---")

    # ======================
    # TOP 5 SETORES
    # ======================

    st.subheader("🏭 Top 5 Setores com Mais Faltas")

    setores = df.groupby("SECAO_NOME").size().reset_index(name="TOTAL")
    setores = setores.sort_values(by="TOTAL", ascending=False).head(5)

    st.dataframe(setores)

    st.bar_chart(setores.set_index("SECAO_NOME"))

    st.markdown("---")

    # ======================
    # EVOLUÇÃO MENSAL
    # ======================

    st.subheader("📈 Evolução de Faltas por Mês")

    evolucao = df.groupby("MES").size().reset_index(name="TOTAL")

    st.line_chart(evolucao.set_index("MES"))

    st.markdown("---")

    # ======================
    # ENVIO DE EMAIL (WEB)
    # ======================

    if st.button("📧 Enviar e-mails"):

        st.warning("⚠️ Envio de e-mail desativado na versão web por segurança.")
