import streamlit as st
import ezdxf
import pandas as pd
import tempfile
import os
from openpyxl import Workbook


def extract_text_entities(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    texts = []
    for entity in msp:
        if entity.dxftype() in ["TEXT", "MTEXT"]:
            try:
                text = entity.plain_text() if entity.dxftype() == "MTEXT" else entity.text
                x, y = round(entity.dxf.insert.x, 2), round(entity.dxf.insert.y, 2)
                texts.append((x, y, text.strip()))
            except Exception:
                pass

    return texts


def group_by_rows(texts, y_threshold=2.5):
    rows = []
    for x, y, text in sorted(texts, key=lambda t: -t[1]):  # sort by y desc
        added = False
        for row in rows:
            if abs(row[0][1] - y) <= y_threshold:
                row.append((x, y, text))
                added = True
                break
        if not added:
            rows.append([(x, y, text)])
    return rows


def build_tables_from_text_rows(text_rows, x_threshold=10):
    tables = []
    current_table = []
    empty_row_count = 0

    for row in text_rows:
        if len(row) < 2:
            empty_row_count += 1
            if empty_row_count >= 2 and current_table:
                tables.append(current_table)
                current_table = []
            continue

        empty_row_count = 0
        sorted_row = sorted(row, key=lambda t: t[0])
        line = []
        last_x = None
        for x, y, text in sorted_row:
            if last_x is not None and abs(x - last_x) > x_threshold:
                line.append("")
            line.append(text)
            last_x = x
        current_table.append(line)

    if current_table:
        tables.append(current_table)

    return tables


# Streamlit App
st.set_page_config(page_title="Extrair Tabela do DXF", layout="centered")
st.title("📐 Extrator de Tabelas DXF para Excel")

uploaded_file = st.file_uploader("Faça upload do arquivo DXF", type=["dxf"])

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        file_path = os.path.join(tmpdir, uploaded_file.name)

        with open(file_path, "wb") as f:
            f.write(uploaded_file.read())

        st.info("⏳ Processando o arquivo...")

        texts = extract_text_entities(file_path)
        text_rows = group_by_rows(texts)
        tables = build_tables_from_text_rows(text_rows)

        if tables:
            output_path = os.path.join(tmpdir, "tabelas_extraidas.xlsx")
            writer = pd.ExcelWriter(output_path, engine="openpyxl")

            for i, table in enumerate(tables):
                df = pd.DataFrame(table)
                df.to_excel(writer, index=False, header=False, sheet_name=f"Tabela_{i+1}")

            writer.close()

            st.success(f"✅ {len(tables)} tabela(s) extraída(s) com sucesso!")

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Baixar Excel com todas as tabelas",
                    data=f,
                    file_name="tabelas_extraidas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Nenhuma tabela reconhecida com os textos disponíveis.")
