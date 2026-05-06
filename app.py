import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Rateio PDF - Centro de Custo", layout="wide")

st.title("📄 Rateio por Centro de Custo - PDF")
st.write(
    "Envie a fatura em PDF para calcular quantidade de usuários/equipamentos "
    "por centro de custo e valor a pagar."
)


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


def extrair_valor_total_fatura(texto_total):
    padroes = [
        r"Total\s*Fatura[:\s]*R\$\s*([\d\.,]+)",
        r"Valor\s*R\$\s*([\d\.,]+)",
        r"R\$\s*([\d\.,]+)\s*\d{2}/\d{2}/\d{4}"
    ]

    valores = []

    for padrao in padroes:
        encontrados = re.findall(padrao, texto_total, flags=re.IGNORECASE)
        for valor in encontrados:
            valores.append(converter_brl_para_float(valor))

    if valores:
        return max(valores)

    return 0.0


def extrair_dados_pdf(arquivo_pdf):
    registros = []
    texto_total = ""

    with pdfplumber.open(arquivo_pdf) as pdf:
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text()

            if not texto:
                continue

            texto_total += texto + "\n"
            linhas = texto.split("\n")

            for linha in linhas:
                linha = linha.strip()

                if "Notebook" not in linha and "Notebool" not in linha:
                    continue

                cc_match = re.search(r"(EC-\d{3})", linha, flags=re.IGNORECASE)
                valor_match = re.search(r"R\$\s*([\d\.,]+)", linha)

                if cc_match and valor_match:
                    centro_custo = cc_match.group(1).upper()
                    valor = converter_brl_para_float(valor_match.group(1))

                    registros.append({
                        "Página": numero_pagina,
                        "Centro de Custo": centro_custo,
                        "Qtd": 1,
                        "Valor": valor,
                        "Linha Extraída": linha
                    })

    valor_total_fatura = extrair_valor_total_fatura(texto_total)

    return pd.DataFrame(registros), valor_total_fatura


arquivo_pdf = st.file_uploader(
    "📁 Envie a fatura em PDF",
    type=["pdf"]
)

if arquivo_pdf:
    st.success("PDF carregado com sucesso!")

    df, valor_total_fatura = extrair_dados_pdf(arquivo_pdf)

    if df.empty:
        st.error("Não consegui identificar os itens da fatura no PDF.")
        st.warning("Verifique se o PDF segue o mesmo modelo da fatura base.")
        st.stop()

    resumo = df.groupby("Centro de Custo").agg(
        qtd_usuarios=("Qtd", "sum"),
        valor_a_pagar=("Valor", "sum")
    ).reset_index()

    resumo = resumo.sort_values("Centro de Custo")
    resumo["Valor Formatado"] = resumo["valor_a_pagar"].apply(formatar_brl)

    total_usuarios = int(resumo["qtd_usuarios"].sum())
    total_valor_itens = resumo["valor_a_pagar"].sum()
    total_centros = resumo["Centro de Custo"].nunique()
    divergencia = round(valor_total_fatura - total_valor_itens, 2)

    st.subheader("📊 Validação da fatura")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Centros de custo", total_centros)
    col2.metric("Total usuários/equipamentos", total_usuarios)
    col3.metric("Soma dos itens", formatar_brl(total_valor_itens))
    col4.metric("Total da fatura", formatar_brl(valor_total_fatura))

    if valor_total_fatura == 0:
        st.warning("Não consegui identificar automaticamente o valor total da fatura no PDF.")
    elif abs(divergencia) <= 0.05:
        st.success("✅ Validação OK: a soma dos itens bate com o total da fatura.")
    else:
        st.error(
            f"⚠️ Divergência encontrada: diferença de {formatar_brl(divergencia)} "
            f"entre a soma dos itens e o total da fatura."
        )

    st.subheader("📋 Base extraída do PDF")
    st.dataframe(df, use_container_width=True)

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

    validacao = pd.DataFrame([{
        "Soma dos Itens": total_valor_itens,
        "Total da Fatura": valor_total_fatura,
        "Divergência": divergencia,
        "Status": "OK" if abs(divergencia) <= 0.05 else "Divergente"
    }])

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tabela_final.to_excel(writer, index=False, sheet_name="Rateio")
        df.to_excel(writer, index=False, sheet_name="Base_Extraida")
        validacao.to_excel(writer, index=False, sheet_name="Validacao")

    st.download_button(
        label="📥 Baixar Excel",
        data=buffer.getvalue(),
        file_name="rateio_centro_custo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Envie o PDF mensal para iniciar o cálculo.")
