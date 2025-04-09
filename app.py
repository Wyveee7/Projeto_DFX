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


def extract_lines(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    h_lines = []  # horizontal
    v_lines = []  # vertical

    for entity in msp:
        if entity.dxftype() == "LINE":
            x1, y1 = round(entity.dxf.start.x, 2), round(entity.dxf.start.y, 2)
            x2, y2 = round(entity.dxf.end.x, 2), round(entity.dxf.end.y, 2)

            if abs(y1 - y2) < 0.5:  # horizontal
                h_lines.append(round((y1 + y2) / 2, 2))
            elif abs(x1 - x2) < 0.5:  # vertical
                v_lines.append(round((x1 + x2) / 2, 2))

    return sorted(set(h_lines), reverse=True), sorted(set(v_lines))


def cluster_lines(lines, threshold=5):
    clusters = []
    for line in lines:
        found = False
        for cluster in clusters:
            if abs(cluster[-1] - line) < threshold:
                cluster.append(line)
                found = True
                break
        if not found:
            clusters.append([line])
    return [sorted(set(c), reverse=(lines == sorted(lines, reverse=True))) for c in clusters if len(c) > 1]


def build_table_from_grid(texts, h_lines, v_lines):
    table = [["" for _ in range(len(v_lines) - 1)] for _ in range(len(h_lines) - 1)]

    for x, y, text in texts:
        row, col = None, None
        for i in range(len(h_lines) - 1):
            if h_lines[i] >= y > h_lines[i + 1]:
                row = i
                break
        for j in range(len(v_lines) - 1):
            if v_lines[j] <= x < v_lines[j + 1]:
                col = j
                break
        if row is not None and col is not None:
            if table[row][col] == "":
                table[row][col] = text
            else:
                table[row][col] += " " + text

    return table


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
        h_lines, v_lines = extract_lines(file_path)

        h_clusters = cluster_lines(h_lines)
        v_clusters = cluster_lines(v_lines)

        workbook = pd.ExcelWriter(os.path.join(tmpdir, "tabelas_extraidas.xlsx"), engine="openpyxl")
        table_count = 0

        for i, h_group in enumerate(h_clusters):
            for j, v_group in enumerate(v_clusters):
                table = build_table_from_grid(texts, h_group, v_group)
                if any(any(cell for cell in row) for row in table):
                    df = pd.DataFrame(table)
                    sheet_name = f"Tabela_{table_count+1}"
                    df.to_excel(workbook, index=False, header=False, sheet_name=sheet_name)
                    table_count += 1

        if table_count > 0:
            workbook.close()
            st.success(f"✅ {table_count} tabela(s) extraída(s) com sucesso!")

            with open(workbook.path, "rb") as f:
                st.download_button(
                    label="📥 Baixar Excel com todas as tabelas",
                    data=f,
                    file_name="tabelas_extraidas.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Nenhuma tabela reconhecida com linhas e textos.")
