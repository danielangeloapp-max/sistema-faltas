import streamlit as st
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
import unicodedata
import plotly.express as px

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

st.title("SISTEMA DE APURAÇÃO DE FALTAS")
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

    df["MES_NUM"] = df["DATA"].dt.month
    df["ANO"] = df["DATA"].dt.year

    df["MES"] = df["DATA"].dt.strftime("%b/%Y").str.upper()

    df = df[df["EVENTO_DECRICAO"].str.contains("FALTA", na=False)]

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
    # TOP 5 GESTORES (COM MESES)
    # ======================

    st.subheader("🏆 Top 5 Equipe de Gestores com Mais Faltas")

    ranking = df.groupby(["EQUIPE", "MES"]).size().reset_index(name="TOTAL")

    pivot = ranking.pivot(index="EQUIPE", columns="MES", values="TOTAL").fillna(0)

    pivot["TOTAL"] = pivot.sum(axis=1)

    # 🔥 ORDENAR MESES CORRETAMENTE
    meses_ordenados = sorted(
        df["DATA"].dropna().dt.to_period("M").astype(str).unique()
    )

    meses_formatados = [pd.to_datetime(m).strftime("%b/%Y").upper() for m in meses_ordenados]

    colunas_finais = meses_formatados + ["TOTAL"]

    pivot = pivot.reindex(columns=colunas_finais, fill_value=0)

    top5 = pivot.sort_values("TOTAL", ascending=False).head(5)

    st.dataframe(top5)

    # gráfico
    fig = px.bar(
        top5.reset_index(),
        x="EQUIPE",
        y="TOTAL",
        text="TOTAL",
        color_discrete_sequence=["#0B3D91"]
    )

    fig.update_traces(textfont_size=16)

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ======================
    # TOP 5 SETORES
    # ======================

    st.subheader("🏭 Top 5 Setores com Mais Faltas")

    setores = df.groupby("SECAO_NOME").size().reset_index(name="TOTAL")
    setores = setores.sort_values(by="TOTAL", ascending=False).head(5)

    st.dataframe(setores)

    fig2 = px.bar(
        setores,
        x="SECAO_NOME",
        y="TOTAL",
        text="TOTAL",
        color_discrete_sequence=["#0B3D91"]
    )

    fig2.update_traces(textfont_size=16)

    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    # ======================
    # EVOLUÇÃO MENSAL
    # ======================

    st.subheader("📈 Evolução de Faltas por Mês")

    evolucao = df.groupby("MES").size().reset_index(name="TOTAL")

    fig3 = px.line(
        evolucao,
        x="MES",
        y="TOTAL",
        markers=True,
        color_discrete_sequence=["#0B3D91"]
    )

    fig3.update_traces(textfont_size=16)

    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")

    # ======================
    # ENVIO DE EMAIL (MANTIDO)
    # ======================

    if st.button("📧 Enviar e-mails"):

        st.info("🚀 Enviando...")

        arquivo_emails = r"C:\Users\BENEL\Desktop\diretorio python\LISTA DE E-MAIL DE GESTORES\E-mail Gestores.xlsx"
        df_emails = pd.read_excel(arquivo_emails)

        df_emails.columns = df_emails.columns.str.strip().str.upper()
        df_emails.rename(columns={"E-MAIL": "EMAIL"}, inplace=True)

        df_emails["EQUIPE"] = df_emails["EQUIPE"].apply(normalizar_texto)
        df_emails["EMAIL"] = df_emails["EMAIL"].astype(str).str.strip()

        emails_gestores = dict(zip(df_emails["EQUIPE"], df_emails["EMAIL"]))

        EMAIL_REMETENTE = "daniel.angelo.app@gmail.com"
        SENHA_APP = "umidwbvhbagtiwxh"

        pasta_saida = "temp_envio"
        os.makedirs(pasta_saida, exist_ok=True)

        grupos = df.groupby("EQUIPE")

        for gestor, dados in grupos:

            gestor_norm = normalizar_texto(gestor)
            email_destino = emails_gestores.get(gestor_norm)

            if not email_destino:
                st.warning(f"⚠️ Sem e-mail: {gestor}")
                continue

            caminho = os.path.join(pasta_saida, f"{gestor}.xlsx")
            dados.to_excel(caminho, index=False)

            msg = EmailMessage()
            msg["Subject"] = f"Relatório de Faltas - {gestor}"
            msg["From"] = EMAIL_REMETENTE
            msg["To"] = email_destino

            msg.set_content(f"""Olá {gestor},

Segue o relatório de faltas da sua equipe.

Att,
Departamento Pessoal
""")

            with open(caminho, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=f"{gestor}.xlsx")

            try:
                with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
                    smtp.starttls()
                    smtp.login(EMAIL_REMETENTE, SENHA_APP)
                    smtp.send_message(msg)

                st.success(f"📧 Enviado: {gestor}")

            except Exception as e:
                st.error(f"❌ Erro {gestor}: {e}")

        st.success("🚀 Envio finalizado!")