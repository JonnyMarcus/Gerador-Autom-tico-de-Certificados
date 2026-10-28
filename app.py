"""Interface web do gerador, preparada para Streamlit Community Cloud."""

from __future__ import annotations

from io import BytesIO
from hashlib import sha256
from pathlib import Path
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st
import streamlit.components.v1 as components

from gerador_certificados.planilha import ler_participantes
from gerador_certificados.pdf_certificado import gerar_certificados_pdf
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
      @media (max-width: 700px) { .block-container { padding:1.1rem 1rem 3rem; } }
    </style>
    """,
    unsafe_allow_html=True,
)

animar_entrada = not st.session_state.get("entrada_animada", False)
st.session_state["entrada_animada"] = True

components.html(
    """
    <!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
    <script src="https://cdn.jsdelivr.net/npm/gsap@3.13.0/dist/gsap.min.js"></script>
    <style>
      *{box-sizing:border-box} html,body{margin:0;padding:0;background:transparent;font-family:Inter,ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif}
      .hero{height:326px;position:relative;overflow:hidden;border-radius:26px;padding:36px 44px;color:#fff;background:radial-gradient(ellipse at 83% 12%,rgba(123,115,255,.48),transparent 35%),linear-gradient(118deg,#17233c 0%,#253865 57%,#5148e5 100%);box-shadow:0 18px 42px rgba(38,53,94,.16)}
      .hero:before,.hero:after{content:"";position:absolute;border:1px solid rgba(255,255,255,.1);border-radius:50%;pointer-events:none}
      .hero:before{width:370px;height:370px;right:45px;top:-155px}.hero:after{width:500px;height:500px;right:-20px;top:-215px}
      .hero-copy{position:relative;z-index:2;width:58%;padding-top:7px}.eyebrow{display:inline-flex;align-items:center;gap:8px;padding:7px 11px;border:1px solid rgba(255,255,255,.2);border-radius:30px;background:rgba(255,255,255,.1);color:#e7eaff;font-size:11px;font-weight:750;letter-spacing:.12em;text-transform:uppercase}
      .eyebrow i{width:7px;height:7px;border-radius:50%;background:#50e0c2;box-shadow:0 0 12px #50e0c2}.title{margin:17px 0 12px;font-size:clamp(30px,4vw,43px);line-height:1.08;letter-spacing:-.045em;font-weight:760}.title-line{display:block;overflow:hidden}.title-line span{display:block}.copy{max-width:450px;color:#d8e2ff;font-size:14px;line-height:1.65;margin:0}.hero-foot{display:flex;align-items:center;gap:9px;margin-top:20px;color:#becbf0;font-size:11px}.hero-foot b{color:#fff;font-weight:650}.spark{color:#58e0c5;font-size:16px}
      .art{position:absolute;z-index:1;right:7.5%;top:47px;width:250px;height:230px;display:grid;place-items:center;perspective:800px}.orbit{position:absolute;inset:0;border:1px solid rgba(255,255,255,.16);border-radius:50%;transform:rotate(-20deg) scaleY(.54)}.orbit:after{content:"";position:absolute;left:18%;top:7%;width:8px;height:8px;border-radius:50%;background:#64e3ce;box-shadow:0 0 18px #64e3ce}
      .certificate{position:relative;width:208px;height:148px;padding:16px 17px;border:1px solid rgba(255,255,255,.8);border-radius:12px;background:linear-gradient(145deg,#fff,#edf2ff);box-shadow:0 22px 48px rgba(8,16,42,.32);transform:rotate(5deg)}.cert-top{display:flex;justify-content:space-between;align-items:center;color:#697799;font-size:7px;letter-spacing:.1em;text-transform:uppercase}.seal{width:21px;height:21px;border:1px solid #776ee7;border-radius:50%;display:grid;place-items:center;color:#5148e5;font-size:10px}.cert-label{text-align:center;color:#5148e5;font-size:7px;font-weight:800;letter-spacing:.16em;margin-top:13px}.cert-name{text-align:center;color:#182647;font:italic 19px Georgia,serif;margin-top:5px}.cert-rule{height:1px;width:75%;margin:6px auto;background:#cbd3e5}.cert-copy{text-align:center;color:#7d88a0;font-size:6px;line-height:1.5}.cert-bottom{display:flex;justify-content:space-between;margin-top:11px;color:#7d88a0;font-size:6px}.cert-bottom span{border-top:1px solid #c8d0e0;padding-top:3px;width:61px;text-align:center}
      .float-chip{position:absolute;padding:8px 11px;border:1px solid rgba(255,255,255,.18);border-radius:10px;background:rgba(255,255,255,.12);backdrop-filter:blur(8px);color:#fff;font-size:9px;font-weight:650;box-shadow:0 9px 22px rgba(8,16,42,.15)}.chip-one{right:0;top:32px}.chip-two{left:0;bottom:26px}.check{color:#55e0c2;margin-right:5px}
      @media(max-width:700px){.hero{height:326px;padding:26px 23px;border-radius:20px}.hero-copy{width:100%;padding-top:1px}.title{font-size:32px;max-width:360px}.copy{font-size:13px;max-width:350px}.art{width:150px;height:125px;right:7%;top:auto;bottom:-8px;opacity:.82}.certificate{transform:scale(.77) rotate(5deg)}.orbit{inset:0}.chip-one{right:0;top:9px}.chip-two{left:-5px;bottom:2px}}
      @media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important}}
    </style></head><body>
    <section class="hero">
      <div class="hero-copy">
        <div class="eyebrow"><i></i> Certifica · emissão simplificada</div>
        <h1 class="title"><span class="title-line"><span>Transforme sua lista</span></span><span class="title-line"><span>em certificados prontos.</span></span></h1>
        <p class="copy">Envie a lista de participantes e seu certificado em PDF. Cada nome vira um certificado pronto para compartilhar.</p>
        <div class="hero-foot"><span class="spark">✦</span><span><b>Simples de preparar.</b> Fácil de compartilhar.</span></div>
      </div>
      <div class="art"><div class="orbit"></div><div class="certificate"><div class="cert-top"><span>Certifica · 2026</span><span class="seal">✦</span></div><div class="cert-label">CERTIFICADO DE CONCLUSÃO</div><div class="cert-name">Seu participante</div><div class="cert-rule"></div><div class="cert-copy">concluiu com êxito o programa<br>de formação e desenvolvimento</div><div class="cert-bottom"><span>Coordenação</span><span>Data de emissão</span></div></div><div class="float-chip chip-one"><span class="check">✓</span>Personalizado</div><div class="float-chip chip-two">PDF pronto para compartilhar</div></div>
    </section>
    <script>
      const reduceMotion=window.matchMedia('(prefers-reduced-motion: reduce)').matches;
      const playIntro=__PLAY_INTRO__;
      if(window.gsap&&!reduceMotion&&playIntro){
        const intro=gsap.timeline({defaults:{ease:'power3.out'}});
        intro.from('.hero',{autoAlpha:0,y:20,duration:.65})
          .from('.eyebrow',{autoAlpha:0,y:12,duration:.38},'-=.28')
          .from('.title-line span',{yPercent:115,duration:.72,stagger:.12},'-=.18')
          .from('.copy,.hero-foot',{autoAlpha:0,y:12,duration:.45,stagger:.1},'-=.35')
          .from('.certificate',{autoAlpha:0,y:22,rotateY:-18,rotateZ:-3,duration:.8},'-=.72')
          .from('.float-chip',{autoAlpha:0,y:10,scale:.92,duration:.35,stagger:.1},'-=.45');
        gsap.to('.certificate',{y:-7,rotateZ:3,duration:2.8,ease:'sine.inOut',repeat:-1,yoyo:true,delay:1.4});
        gsap.to('.orbit',{rotate:0,duration:22,ease:'none',repeat:-1});
        gsap.to('.eyebrow i',{scale:1.35,opacity:.65,duration:.9,repeat:-1,yoyo:true,ease:'sine.inOut'});
      }
    </script></body></html>
    """.replace("__PLAY_INTRO__", str(animar_entrada).lower()),
    height=350,
    scrolling=False,
)

with st.expander("Antes de começar · confira o formato dos arquivos", expanded=False):
    st.markdown(
        """
        **Planilha:** arquivo `.xlsx`, com os cabeçalhos na primeira linha e uma coluna chamada `nome`.

        **Certificado:** arquivo `.pdf`. O nome será aplicado ao centro da primeira página.

        Deixe uma área livre no centro do PDF para o nome do participante.
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
    st.markdown('<div class="section-title">2. Certificado base</div><div class="section-help">Envie o arquivo PDF do certificado</div>', unsafe_allow_html=True)
    arquivo_pdf = st.file_uploader("Escolher certificado PDF", type=["pdf"], label_visibility="collapsed", key="pdf_upload")
    if arquivo_pdf:
        uploaded_count += 1
        st.caption(f"✓ {arquivo_pdf.name}")

step_class = "active" if uploaded_count == 2 else ""
st.markdown(
    f'<div class="step-card {step_class}"><strong>3. Gerar e baixar</strong> · Os arquivos ficam disponíveis em um único ZIP</div>',
    unsafe_allow_html=True,
)

coluna_nome = "nome"

fingerprint = (
    coluna_nome,
    sha256(arquivo_planilha.getvalue()).hexdigest() if arquivo_planilha else None,
    sha256(arquivo_pdf.getvalue()).hexdigest() if arquivo_pdf else None,
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
    if not arquivo_planilha or not arquivo_pdf:
        st.error("Envie a planilha e o PDF do certificado antes de continuar.")
    else:
        try:
            with tempfile.TemporaryDirectory(prefix="gerador_certificados_") as diretorio:
                pasta = Path(diretorio)
                caminho_planilha = pasta / "participantes.xlsx"
                caminho_pdf = pasta / "certificado.pdf"
                caminho_planilha.write_bytes(arquivo_planilha.getvalue())
                caminho_pdf.write_bytes(arquivo_pdf.getvalue())

                participantes = ler_participantes(caminho_planilha)
                validar_participantes(participantes, coluna_nome, campos=[])
                arquivos = gerar_certificados_pdf(
                    caminho_pdf,
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
