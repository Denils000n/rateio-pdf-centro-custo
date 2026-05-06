import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Rateio PDF - Centro de Custo", layout="wide")

st.title("📄 Rateio por Centro de Custo - PDF")
st.write("Envie a fatura em PDF para calcular quantidade de usuários/equipamentos por centro de custo e valor a pagar.")


def converter_brl_para_float(valor):
    if valor is None:
        return 0.0

    valor = str(valor).strip()
    valor = valor.replace("R$", "").replace(" ", "")
    valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except:
        return 0.0


def formatar_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def limpar_centro_custo(cc):
    if not cc:
        return ""

    cc = str(cc).strip().upper()
    cc = cc.replace(" ", "")
    return cc


def extrair_linhas_pdf(arquivo_pdf):
    registros = []

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()

            if not texto:
                continue

            linhas = texto.split("\n")

            for linha in linhas:
                linha = linha.strip()

                # Captura linhas de item que começam com quantidade
                # Exemplo:
                # 1 Notebook Dell ... EC-706 202617-6304 005723 J487KH3 01/04/2026 30/04/2026 R$ 198,00
                padrao = re.search(
                    r"^(\d+)\s+(.+?)\s+(EC-\d{3})\s+(\S+)\s+(\S+)\s+(\S+)\s+"
                    r"(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+R?\$?\s*([\d\.,]+)$",
                    linha,
                    flags=re.IGNORECASE
                )

                if padrao:
                    qtd = int(padrao.group(1))
                    descricao = padrao.group(2).strip()
                    centro_custo = limpar_centro_custo(padrao.group(3))
                    contrato = padrao.group(4)
                    patrimonio = padrao.group(5)
                    serial = padrao.group(6)
                    data_de = padrao.group(7)
                    data_ate = padrao.group(8)
                    valor = converter_brl_para_float(padrao.group(9))

                    registros.append({
                        "Qtd": qtd,
                        "Descrição": descricao,
                        "Centro de Custo": centro_custo,
                        "Contrato": contrato,
                        "Patrimônio": patrimonio,
                        "Serial": serial,
                        "De": data_de,
                        "Até": data_ate,
                        "Valor": valor
                    })

    return pd.DataFrame(registros)


arquivo_pdf = st.file_uploader(
    "📁 Envie a fatura em PDF",
    type=["pdf"]
)

if arquivo_pdf:
    st.success("PDF carregado com sucesso!")

    df = extrair_linhas_pdf(arquivo_pdf)

    if df.empty:
        st.error("Não consegui identificar os itens da fatura no PDF.")
        st.warning("Verifique se o PDF segue o mesmo modelo da fatura base.")
        st.stop()

    st.subheader("📋 Base extraída do PDF")
    st.dataframe(df, use_container_width=True)

    resumo = df.groupby("Centro de Custo").agg(
        qtd_usuarios=("Qtd", "sum"),
        valor_a_pagar=("Valor", "sum")
    ).reset_index()

    resumo = resumo.sort_values("Centro de Custo")

    resumo["Valor Formatado"] = resumo["valor_a_pagar"].apply(formatar_brl)

    total_usuarios = int(resumo["qtd_usuarios"].sum())
    total_valor = resumo["valor_a_pagar"].sum()
    total_centros = resumo["Centro de Custo"].nunique()

    st.subheader("📊 Resultado do Rateio")

    col1, col2, col3 = st.columns(3)
    col1.metric("Centros de custo", total_centros)
    col2.metric("Total usuários/equipamentos", total_usuarios)
    col3.metric("Valor total", formatar_brl(total_valor))

    st.subheader("📌 Resumo por centro de custo")

    for _, row in resumo.iterrows():
        st.success(
            f"• {row['Centro de Custo']} – "
            f"{int(row['qtd_usuarios'])} usuário(s) | "
            f"Valor a pagar: {row['Valor Formatado']}"
        )

    tabela_final = resumo[[
        "Centro de Custo",
        "qtd_usuarios",
        "valor_a_pagar"
    ]].copy()

    tabela_final.columns = [
        "Centro de Custo",
        "Qtd Usuários",
        "Valor a Pagar"
    ]

    tabela_final["Valor a Pagar"] = tabela_final["Valor a Pagar"].round(2)

    st.subheader("📑 Tabela final")
    st.dataframe(tabela_final, use_container_width=True)

    st.subheader("📧 Texto para copiar e colar no e-mail")

    linhas_email = []

    for _, row in resumo.iterrows():
        linhas_email.append(
            f"• {row['Centro de Custo']} – "
            f"{int(row['qtd_usuarios'])} usuário(s) | "
            f"Valor a pagar: {row['Valor Formatado']}"
        )

    texto_email = (
        "Prezados,\n\n"
        "Segue abaixo o rateio atualizado por centro de custo:\n\n"
        + "\n".join(linhas_email)
        + "\n\n"
        "Fico à disposição para qualquer dúvida."
    )

    st.text_area(
        "Copie o texto abaixo:",
        value=texto_email,
        height=350
    )

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tabela_final.to_excel(writer, index=False, sheet_name="Rateio")
        df.to_excel(writer, index=False, sheet_name="Base_Extraida")

    st.download_button(
        label="📥 Baixar Excel",
        data=buffer.getvalue(),
        file_name="rateio_centro_custo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Envie o PDF mensal para iniciar o cálculo.")