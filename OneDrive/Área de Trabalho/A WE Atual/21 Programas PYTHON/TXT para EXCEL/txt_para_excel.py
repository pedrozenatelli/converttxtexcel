import streamlit as st
import openpyxl
import re
import io
from openpyxl import Workbook

st.set_page_config(page_title="Conversor de Questões TXT para Excel", layout="centered")

st.title("📄➡️📊 Conversor de Questões (.txt) para Excel (.xlsx)")

uploaded_file = st.file_uploader("📤 Envie o arquivo .txt com as questões:", type=["txt"])

def extrair_entre(texto, inicio, fim):
    padrao = re.escape(inicio) + r"(.*?)" + re.escape(fim)
    resultado = re.search(padrao, texto, re.DOTALL)
    return resultado.group(1).strip() if resultado else ""

if uploaded_file is not None:
    nome_base = uploaded_file.name.rsplit(".", 1)[0]
    conteudo = uploaded_file.read().decode("utf-8")

    blocos = conteudo.split("###ENUNCIADO###")[1:]

    # Cria planilha na memória
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Questões"

    # Cabeçalhos
    sheet["B1"] = "Enunciado"
    sheet["C1"] = "Comentário"
    sheet["P1"] = "Alternativa A"
    sheet["Q1"] = "Alternativa B"
    sheet["R1"] = "Alternativa C"
    sheet["S1"] = "Alternativa D"
    sheet["U1"] = "Gabarito"

    linha = 2

    for bloco in blocos:
        try:
            enunciado = extrair_entre(bloco, "", "###ALTERNATIVA_A###")
            alt_a = extrair_entre(bloco, "###ALTERNATIVA_A###", "###ALTERNATIVA_B###")
            alt_b = extrair_entre(bloco, "###ALTERNATIVA_B###", "###ALTERNATIVA_C###")
            alt_c = extrair_entre(bloco, "###ALTERNATIVA_C###", "###ALTERNATIVA_D###")
            alt_d = extrair_entre(bloco, "###ALTERNATIVA_D###", "###GABARITO###")

            gabarito_match = re.search(r"<b>GABARITO:</b>\s*([A-D])", bloco)
            gabarito = gabarito_match.group(1).strip() if gabarito_match else ""

            comentario = extrair_entre(bloco, "###GABARITO###", "###ENUNCIADO###")
            if comentario == "":
                comentario = bloco.split("###GABARITO###")[1].strip()

            # Substituições e formatação
            comentario = comentario.replace("<b>GABARITO:</b>", "<b>Gabarito:</b>")
            comentario = re.sub(r"\n*\s*<b>EXPLICAÇÃO:</b>", r"\n\n<b>Comentário:</b>", comentario)
            comentario = re.sub(r"\n*\s*<span style=", r"\n\n<b>Texto original:</b>&nbsp;<span style=", comentario)

            # Preencher a planilha
            sheet[f"B{linha}"] = enunciado
            sheet[f"C{linha}"] = comentario
            sheet[f"P{linha}"] = alt_a
            sheet[f"Q{linha}"] = alt_b
            sheet[f"R{linha}"] = alt_c
            sheet[f"S{linha}"] = alt_d
            sheet[f"U{linha}"] = gabarito

            linha += 1

        except Exception as e:
            st.error(f"Erro ao processar uma questão: {e}")

    # Salva a planilha em memória
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    st.success("✅ Conversão concluída!")

    st.download_button(
        label="📥 Baixar arquivo Excel",
        data=output,
        file_name=f"{nome_base}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
