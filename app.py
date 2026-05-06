import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Rateio PDF - Centro de Custo", layout="wide")

st.title("📄 Rateio por Centro de Custo - PDF")
st.write("Envie a fatura em PDF para calcular quantidade por centro de custo e valor a pagar.")


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


def extrair_total_fatura(texto_total):
    padrao = r"Total\s*Fatura[:\s]*R\$\s*([\d\.,]+)"
    encontrado = re.search(padrao, texto_total, flags=re.IGNORECASE)

    if encontrado:
        return converter_brl_para_float(encontrado.group(1))

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

            for linha in texto.split("\n"):
                linha = linha.strip()

                # Considera apenas linhas reais de item
                if not linha.startswith("1 "):
                    continue

                # Centro de custo no padrão EC-XXX
                centro_match = re.search(r"EC-(\d{3})", linha, flags=re.IGNORECASE)

                # Pega o último valor em R$ da linha
                valores = re.findall(r"R\$\s*([\d\.,]+)", linha)

                if centro_match and valores:
                    centro_custo = centro_match.group(1)
                    valor = converter_brl_para_float(valores[-1])

                    registros.append({
                        "Página": numero_pagina,
                        "Centro de Custo": centro_custo,
                        "Qtd": 1,
                        "Valor": valor,
                        "Linha Extraída": linha
                    })

    total_fatura = extrair_total_fatura(texto_total)

    return pd.DataFrame(registros), total_fatura


arquivo_pdf = st.file_uploader(
    "📁 Envie a fatura em PDF",
    type=["pdf"]
)

if arquivo_pdf:
    st.success("PDF carregado com sucesso!")

    df, total_fatura = extrair_dados_pdf(arquivo_pdf)

    if df.empty:
        st.error("Não consegui identificar os itens da fatura no PDF.")
        st.stop()

    resumo = df.groupby("Centro de Custo").agg(
        qtd_usuarios=("Qtd", "sum"),
        valor_a_pagar=("Valor", "sum")
    ).reset_index()

    resumo["Centro de Custo"] = resumo["Centro de Custo"].astype(int)
    resumo = resumo.sort_values("Centro de Custo")
    resumo["Centro de Custo"] = resumo["Centro de Custo"].astype(str)

    resumo["Valor Formatado"] = resumo["valor_a_pagar"].apply(formatar_brl)

    soma_itens = resumo["valor_a_pagar"].sum()
    divergencia = round(total_fatura - soma_itens, 2)

    st.subheader("📊 Validação da fatura")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Centros de custo", resumo["Centro de Custo"].nunique())
    col2.metric("Total usuários", int(resumo["qtd_usuarios"].sum()))
    col3.metric("Soma dos itens", formatar_brl(soma_itens))
    col4.metric("Total da fatura", formatar_brl(total_fatura))

    if total_fatura > 0:
        if abs(divergencia) <= 0.05:
            st.success("✅ A soma dos itens bate com o total da fatura.")
        else:
            st.error(f"⚠️ Divergência encontrada: {formatar_brl(divergencia)}")
    else:
        st.warning("Não consegui identificar automaticamente o total da fatura.")

    st.subheader("📌 Resumo por centro de custo")

    for _, row in resumo.iterrows():
        texto_usuario = "usuário" if int(row["qtd_usuarios"]) == 1 else "usuários"

        st.success(
            f"• {row['Centro de Custo']} – "
            f"{int(row['qtd_usuarios'])} {texto_usuario} | "
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

    st.subheader("📋 Base extraída do PDF")
    st.dataframe(df, use_container_width=True)

    st.subheader("📧 Texto para copiar e colar no e-mail")

    linhas_email = []

    for _, row in resumo.iterrows():
        texto_usuario = "usuário" if int(row["qtd_usuarios"]) == 1 else "usuários"

        linhas_email.append(
            f"• {row['Centro de Custo']} – "
            f"{int(row['qtd_usuarios'])} {texto_usuario} | "
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
        height=400
    )

    validacao = pd.DataFrame([{
        "Soma dos Itens": soma_itens,
        "Total da Fatura": total_fatura,
        "Divergência": divergencia,
        "Status": "OK" if abs(divergencia) <= 0.05 else "Divergente"
    }])

    buffer = BytesIO()

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
    st.info("Envie o PDF mensal para iniciar.")
