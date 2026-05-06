import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

st.set_page_config(page_title="Rateio PDF - Centro de Custo", layout="wide")

st.title("📄 Rateio por Centro de Custo - PDF")


# ----------------------------
# FUNÇÕES
# ----------------------------

def converter_brl_para_float(valor):
    if valor is None:
        return 0.0

    valor = str(valor).strip()
    valor = valor.replace("R$", "")
    valor = valor.replace(" ", "")  # remove espaço quebrado do PDF
    valor = valor.replace(".", "")
    valor = valor.replace(",", ".")

    try:
        return float(valor)
    except:
        return 0.0


def formatar_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def extrair_total_fatura(texto_total):
    match = re.search(r"Total\s*Fatura[:\s]*R\$\s*([\d\.\s]+,\d{2})", texto_total, re.IGNORECASE)

    if match:
        return converter_brl_para_float(match.group(1))

    return 0.0


def extrair_dados_pdf(arquivo_pdf):
    registros = []
    texto_total = ""

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()

            if not texto:
                continue

            texto_total += texto + "\n"

            for linha in texto.split("\n"):
                linha = linha.strip()

                # Considera apenas linhas reais de item (sempre começam com número)
                if not linha.startswith("1 "):
                    continue

                # Centro de custo
                cc_match = re.search(r"EC-(\d{3})", linha, re.IGNORECASE)

                # Valor (corrigido para pegar "R$ 1 98,00")
                valor_match = re.findall(r"R\$\s*([\d\.\s]+,\d{2})", linha)

                if cc_match and valor_match:
                    centro = cc_match.group(1)

                    # pega o último valor da linha (mais seguro)
                    valor = converter_brl_para_float(valor_match[-1])

                    registros.append({
                        "Centro de Custo": centro,
                        "Qtd": 1,
                        "Valor": valor,
                        "Linha": linha
                    })

    total_fatura = extrair_total_fatura(texto_total)

    return pd.DataFrame(registros), total_fatura


# ----------------------------
# APP
# ----------------------------

arquivo_pdf = st.file_uploader("📁 Envie o PDF da fatura", type=["pdf"])

if arquivo_pdf:

    df, total_fatura = extrair_dados_pdf(arquivo_pdf)

    if df.empty:
        st.error("Não consegui ler o PDF corretamente.")
        st.stop()

    resumo = df.groupby("Centro de Custo").agg(
        qtd_usuarios=("Qtd", "sum"),
        valor=("Valor", "sum")
    ).reset_index()

    resumo["Centro de Custo"] = resumo["Centro de Custo"].astype(int)
    resumo = resumo.sort_values("Centro de Custo")
    resumo["Centro de Custo"] = resumo["Centro de Custo"].astype(str)

    resumo["Valor Formatado"] = resumo["valor"].apply(formatar_brl)

    soma_itens = resumo["valor"].sum()
    divergencia = round(total_fatura - soma_itens, 2)

    # ----------------------------
    # DASHBOARD
    # ----------------------------

    st.subheader("📊 Validação")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Centros", resumo.shape[0])
    col2.metric("Usuários", int(resumo["qtd_usuarios"].sum()))
    col3.metric("Soma Itens", formatar_brl(soma_itens))
    col4.metric("Fatura", formatar_brl(total_fatura))

    if total_fatura > 0:
        if abs(divergencia) <= 0.05:
            st.success("✅ Valores batem com a fatura")
        else:
            st.error(f"⚠️ Divergência: {formatar_brl(divergencia)}")

    # ----------------------------
    # RESULTADO
    # ----------------------------

    st.subheader("📌 Rateio por Centro de Custo")

    linhas_email = []

    for _, row in resumo.iterrows():
        usuarios = int(row["qtd_usuarios"])
        texto_usuario = "usuário" if usuarios == 1 else "usuários"

        linha = (
            f"• {row['Centro de Custo']} – "
            f"{usuarios} {texto_usuario} | "
            f"Valor a pagar: {row['Valor Formatado']}"
        )

        st.success(linha)
        linhas_email.append(linha)

    # ----------------------------
    # EMAIL
    # ----------------------------

    st.subheader("📧 Texto para e-mail")

    texto_email = (
        "Prezados,\n\n"
        "Segue abaixo o rateio atualizado por centro de custo:\n\n"
        + "\n".join(linhas_email)
        + "\n\nFico à disposição para qualquer dúvida."
    )

    st.text_area("Copiar:", texto_email, height=400)

    # ----------------------------
    # TABELA
    # ----------------------------

    tabela = resumo[[
        "Centro de Custo",
        "qtd_usuarios",
        "valor"
    ]].copy()

    tabela.columns = [
        "Centro de Custo",
        "Qtd Usuários",
        "Valor a Pagar"
    ]

    tabela["Valor a Pagar"] = tabela["Valor a Pagar"].round(2)

    st.subheader("📑 Tabela")
    st.dataframe(tabela, use_container_width=True)

    # ----------------------------
    # DOWNLOAD
    # ----------------------------

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tabela.to_excel(writer, index=False, sheet_name="Rateio")
        df.to_excel(writer, index=False, sheet_name="Base")

    st.download_button(
        "📥 Baixar Excel",
        buffer.getvalue(),
        "rateio.xlsx"
    )

else:
    st.info("Envie o PDF para iniciar")
