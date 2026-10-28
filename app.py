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


st.set_page_config(page_title="Certifica | Gerador de certificados", page_icon="🎓", layout="wide")


def criar_zip(arquivos: list[Path]) -> bytes:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as pacote:
        for arquivo in arquivos:
            pacote.write(arquivo, arcname=arquivo.name)
    return buffer.getvalue()


st.markdown(
    """
    <style>
      :root { --ink:#17233c; --muted:#66748b; --brand:#5148e5; --mint:#22b8a5; }
      [data-testid="stAppViewContainer"] { background:linear-gradient(180deg,#f3f6fc 0%,#fbfcff 340px,#fbfcff 100%); }
      [data-testid="stHeader"] { background:transparent; }
      .block-container { max-width:1120px; padding-top:2rem; padding-bottom:4rem; }
      .brand-line { color:#5148e5; font-size:.78rem; font-weight:800; letter-spacing:.13em; text-transform:uppercase; margin-bottom:.7rem; }
      .hero-card { background:linear-gradient(120deg,#182647 0%,#293a68 58%,#5148e5 100%); border-radius:24px; padding:2.2rem 2.5rem; color:white; box-shadow:0 18px 45px rgba(38,53,94,.18); margin-bottom:1.3rem; }
      .hero-card h1 { color:white; font-size:2.25rem; line-height:1.12; margin:0 0 .65rem 0; letter-spacing:-.04em; }
      .hero-card p { color:#dce5ff; font-size:1.02rem; margin:0; max-width:700px; }
      .hero-chip { display:inline-block; background:rgba(255,255,255,.13); border:1px solid rgba(255,255,255,.2); border-radius:99px; padding:.3rem .7rem; color:#e9edff; font-size:.75rem; font-weight:700; margin-bottom:1rem; }
      .step-card { background:#fff; border:1px solid #e5eaf2; border-radius:14px; padding:.85rem 1rem; color:#66748b; font-size:.88rem; margin:.4rem 0 1.2rem; }
      .step-card strong { color:#17233c; }
      .step-card.active { border-color:#b9b5ff; background:#f4f3ff; color:#5148e5; }
      .section-title { color:#17233c; font-size:1.25rem; font-weight:750; margin:.7rem 0 .2rem; }
      .section-help { color:#718096; font-size:.9rem; margin-bottom:1rem; }
      [data-testid="stFileUploader"] section { background:#fff; border:1px dashed #bac6d8; border-radius:15px; }
      [data-testid="stFileUploader"] section:hover { border-color:#5148e5; }
      [data-testid="stExpander"] { background:#fff; border:1px solid #e6eaf1; border-radius:14px; }
      div.stButton > button, div.stDownloadButton > button { min-height:3rem; border-radius:12px; font-weight:700; }
      [data-testid="stDataFrame"] { border:1px solid #e5eaf2; border-radius:12px; overflow:hidden; }
      .privacy-note { color:#7b8798; font-size:.78rem; text-align:center; margin-top:2rem; }
      @media (max-width: 700px) { .block-container { padding:1.1rem 1rem 3rem; } .hero-card { padding:1.5rem; border-radius:18px; } .hero-card h1 { font-size:1.8rem; } }
    </style>
    <div class="brand-line">Certifica · emissão simplificada</div>
    <div class="hero-card">
      <span class="hero-chip">CERTIFICADOS EM LOTE</span>
      <h1>Transforme sua lista<br>em certificados prontos.</h1>
      <p>Envie os participantes e seu modelo. A gente personaliza cada certificado e entrega tudo em um único arquivo.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Antes de começar · confira o formato dos arquivos", expanded=False):
    st.markdown(
        """
        **Planilha:** arquivo `.xlsx`, cabeçalhos na primeira linha e uma coluna com o nome de cada participante.

        **Modelo:** arquivo `.pptx` com marcadores entre chaves, como `{nome}`, `{curso}` e `{carga_horaria}`.

        O nome de cada marcador precisa corresponder ao cabeçalho da planilha. Mantenha cada marcador inteiro no mesmo parágrafo do PowerPoint.
        """
    )

uploaded_count = 0

left_step, right_step = st.columns(2)
with left_step:
    st.markdown('<div class="section-title">1. Lista de participantes</div><div class="section-help">Selecione a planilha Excel (.xlsx)</div>', unsafe_allow_html=True)
    arquivo_planilha = st.file_uploader("Escolher planilha", type=["xlsx"], label_visibility="collapsed", key="planilha_upload")
    if arquivo_planilha:
        uploaded_count += 1
        st.caption(f"✓ {arquivo_planilha.name}")
with right_step:
    st.markdown('<div class="section-title">2. Modelo visual</div><div class="section-help">Selecione o certificado PowerPoint (.pptx)</div>', unsafe_allow_html=True)
    arquivo_modelo = st.file_uploader("Escolher modelo", type=["pptx"], label_visibility="collapsed", key="modelo_upload")
    if arquivo_modelo:
        uploaded_count += 1
        st.caption(f"✓ {arquivo_modelo.name}")

step_class = "active" if uploaded_count == 2 else ""
st.markdown(
    f'<div class="step-card {step_class}"><strong>3. Gerar e baixar</strong> · Os arquivos ficam disponíveis em um único ZIP</div>',
    unsafe_allow_html=True,
)

name_col, info_col = st.columns([1, 2])
with name_col:
    coluna_nome = st.text_input("Nome da coluna de participantes", value="nome", help="Deve corresponder ao cabeçalho na primeira linha da planilha.").strip().lower()
with info_col:
    st.markdown('<div style="padding-top:2rem;color:#718096;font-size:.9rem">Os demais campos da planilha são usados para preencher marcadores equivalentes no modelo.</div>', unsafe_allow_html=True)

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
        st.markdown(f'<div class="section-title">Prévia da planilha <span style="color:#5148e5;font-size:.9rem">· {len(amostra)} participantes</span></div>', unsafe_allow_html=True)
        st.dataframe(amostra[:10], use_container_width=True, hide_index=True)
    except (ValueError, OSError, KeyError) as error:
        st.error(f"Não foi possível ler a planilha: {error}")

st.markdown("<div style='height:.5rem'></div>", unsafe_allow_html=True)
if st.button("✦  Gerar certificados", type="primary", use_container_width=True):
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
    st.success("Seus certificados estão prontos. Baixe o pacote abaixo.")
    st.download_button(
        "↓  Baixar certificados (.zip)",
        data=st.session_state["zip_certificados"],
        file_name="certificados.zip",
        mime="application/zip",
        type="primary",
        use_container_width=True,
    )

st.markdown(
    '<div class="privacy-note">Seus arquivos são processados para esta geração e não devem conter dados confidenciais.</div>',
    unsafe_allow_html=True,
)
