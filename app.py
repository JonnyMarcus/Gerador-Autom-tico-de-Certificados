"""Interface web do gerador, preparada para Streamlit Community Cloud."""

from __future__ import annotations

from io import BytesIO
from hashlib import sha256
from pathlib import Path
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from gerador_certificados.exportador import gerar_certificados
from gerador_certificados.planilha import ler_participantes
from gerador_certificados.validacao import validar_participantes


st.set_page_config(page_title="Gerador de certificados", page_icon="🎓", layout="centered")


def criar_zip(arquivos: list[Path]) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as pacote:
        for arquivo in arquivos:
            pacote.write(arquivo, arcname=arquivo.name)
    return buffer.getvalue()


st.title("🎓 Gerador de certificados")
st.write("Envie a planilha dos participantes e um modelo PowerPoint para gerar certificados personalizados.")

with st.expander("Como preparar os arquivos", expanded=True):
    st.markdown(
        """
        - A planilha deve ser `.xlsx`, com os cabeçalhos na primeira linha e uma coluna `nome`.
        - O modelo deve ser `.pptx` e usar marcadores como `{nome}`, `{curso}` e `{carga_horaria}`.
        - O nome de cada marcador deve corresponder ao cabeçalho da planilha.
        - Cada marcador deve estar inteiro no mesmo parágrafo do PowerPoint.
        """
    )

coluna_nome = st.text_input("Coluna que contém o nome", value="nome").strip().lower()
arquivo_planilha = st.file_uploader("Planilha de participantes (.xlsx)", type=["xlsx"])
arquivo_modelo = st.file_uploader("Modelo de certificado (.pptx)", type=["pptx"])

fingerprint = (
    coluna_nome,
    sha256(arquivo_planilha.getvalue()).hexdigest() if arquivo_planilha else None,
    sha256(arquivo_modelo.getvalue()).hexdigest() if arquivo_modelo else None,
)
if st.session_state.get("fingerprint") != fingerprint:
    st.session_state.pop("zip_certificados", None)
    st.session_state["fingerprint"] = fingerprint

if arquivo_planilha:
    try:
        with tempfile.TemporaryDirectory(prefix="certificados_preview_") as diretorio:
            caminho_planilha = Path(diretorio) / "participantes.xlsx"
            caminho_planilha.write_bytes(arquivo_planilha.getvalue())
            amostra = ler_participantes(caminho_planilha)
        st.caption(f"{len(amostra)} participante(s) encontrado(s).")
        st.dataframe(amostra[:10], use_container_width=True, hide_index=True)
    except (ValueError, OSError, KeyError) as error:
        st.error(f"Não foi possível ler a planilha: {error}")

if st.button("Gerar certificados", type="primary", use_container_width=True):
    if not coluna_nome:
        st.error("Informe o cabeçalho da coluna com os nomes.")
    elif not arquivo_planilha or not arquivo_modelo:
        st.error("Envie a planilha e o modelo antes de gerar os certificados.")
    else:
        try:
            with tempfile.TemporaryDirectory(prefix="gerador_certificados_") as diretorio:
                pasta = Path(diretorio)
                caminho_planilha = pasta / "participantes.xlsx"
                caminho_modelo = pasta / "modelo.pptx"
                caminho_planilha.write_bytes(arquivo_planilha.getvalue())
                caminho_modelo.write_bytes(arquivo_modelo.getvalue())

                participantes = ler_participantes(caminho_planilha)
                validar_participantes(participantes, coluna_nome, campos=[])
                arquivos = gerar_certificados(
                    caminho_modelo,
                    pasta / "certificados",
                    participantes,
                    coluna_nome,
                )
                st.session_state["zip_certificados"] = criar_zip(arquivos)

            st.success(f"{len(arquivos)} certificado(s) gerado(s).")
        except (FileNotFoundError, ValueError, PermissionError, OSError) as error:
            st.error(f"Não foi possível gerar os certificados: {error}")
        except Exception as error:
            st.error(f"Ocorreu um erro ao processar os arquivos: {error}")

if st.session_state.get("zip_certificados"):
    st.download_button(
        "Baixar certificados (.zip)",
        data=st.session_state["zip_certificados"],
        file_name="certificados.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )
