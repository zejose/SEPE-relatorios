import io
import zipfile
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

import sepe_core as core

st.set_page_config(page_title="Gerador de Relatórios SEPE", page_icon="🏗️", layout="wide")
ss = st.session_state


def segredo(chave: str, padrao: str = "") -> str:
    """Valor de [odk] no secrets.toml (ou nos Secrets do Streamlit Cloud)."""
    try:
        return str(st.secrets["odk"][chave])
    except Exception:
        return padrao


def carregar(planilha: core.Planilha, fonte: str, *, odk: dict | None = None,
             midia: dict[str, bytes] | None = None) -> None:
    ss.planilha = planilha
    ss.fonte = fonte
    ss.odk = odk
    ss.midia = midia or {}
    # fotos ja baixadas e ZIPs prontos sao da carga anterior
    ss.fotos = {}
    ss.pop("zip_relatorios", None)
    ss.pop("zip_fotos", None)


def cliente_odk() -> core.OdkCentral:
    return core.OdkCentral(**ss.odk)


def faixa(texto: str, total: int) -> set[int]:
    """'1, 3, 5-10' -> posicoes (0-based) validas."""
    posicoes = set()
    for parte in texto.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if "-" in parte:
            a, b = (int(x) for x in parte.split("-", 1))
            posicoes.update(range(a - 1, b))
        else:
            posicoes.add(int(parte) - 1)
    return {p for p in posicoes if 0 <= p < total}


def garantir_fotos(posicoes: list[int]) -> tuple[int, int]:
    """Baixa (ou tira do ZIP) e reduz as fotos que os relatorios vao usar.

    So as dos registros escolhidos, e so uma vez por carga: gerar de novo
    reaproveita o que ja esta em memoria. Devolve (achadas, faltando)."""
    pedidos = core.pedidos_de_fotos(ss.planilha, posicoes)
    novos = [p for p in pedidos if p not in ss.fotos]
    if novos:
        barra = st.progress(0.0, text="Preparando fotos...")
        if ss.fonte == "odk":
            origem = cliente_odk().anexos(novos, reduzir=True)
        else:
            origem = (
                (p, core.reduzir_foto(ss.midia[p[1]]) if p[1] in ss.midia else None)
                for p in novos
            )
        for i, (pedido, foto) in enumerate(origem, 1):
            ss.fotos[pedido] = foto
            barra.progress(i / len(novos), text=f"Fotos: {i}/{len(novos)}")
        barra.empty()
    achadas = sum(1 for p in pedidos if ss.fotos.get(p))
    return achadas, len(pedidos) - achadas


def zip_fotos_originais(posicoes: list[int]) -> bytes:
    planilha: core.Planilha = ss.planilha
    nomes_zip = {}
    for p in posicoes:
        linha = planilha.linhas[p]
        instancia = core.campo(linha, planilha.col_instancia)
        id_proj = core.campo(linha, core.COL_ID_PROJ)
        for c in core.COLS_IMAGENS:
            nome = core.campo(linha, c)
            if nome:
                nomes_zip[(instancia, nome)] = f"foto_{id_proj}_{nome}" if id_proj else f"foto_{nome}"

    buf = io.BytesIO()
    barra = st.progress(0.0, text="Juntando fotos...")
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
        if ss.fonte == "odk":
            origem = cliente_odk().anexos(nomes_zip, reduzir=False)
        else:
            origem = ((p, ss.midia.get(p[1])) for p in nomes_zip)
        for i, (pedido, bruto) in enumerate(origem, 1):
            if bruto:
                zf.writestr(nomes_zip[pedido], bruto)
            barra.progress(i / len(nomes_zip), text=f"Fotos: {i}/{len(nomes_zip)}")
    barra.empty()
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════
st.title("🏗️ Gerador de Relatórios de Vistoria")

aba_odk, aba_arquivo = st.tabs(["📡 ODK Central", "📁 Arquivo exportado"])

with aba_odk:
    senha_nos_secrets = bool(segredo("senha"))
    with st.form("form_odk"):
        c1, c2 = st.columns(2)
        url = c1.text_input("URL do ODK Central",
                            segredo("url", "https://levantamentos.dflegal.df.gov.br"))
        projeto = c1.text_input("ID do projeto", segredo("projeto", "4"))
        formulario = c1.text_input("ID do formulário",
                                   segredo("formulario", "aWX8oXGiD9zmcpDX7KtAer"))
        email = c2.text_input("E-mail", segredo("email"))
        senha = c2.text_input(
            "Senha", type="password",
            placeholder="em branco = senha configurada no servidor" if senha_nos_secrets else "",
        )
        buscar = st.form_submit_button("🔄 Buscar dados", type="primary",
                                       use_container_width=True)

    if buscar:
        senha = senha or segredo("senha")
        if not email or not senha:
            st.error("Preencha e-mail e senha.")
        else:
            credenciais = dict(url=url, projeto=projeto, formulario=formulario,
                               email=email, senha=senha)
            try:
                with st.spinner("Buscando registros no ODK Central..."):
                    texto = core.OdkCentral(**credenciais).csv()
                carregar(core.ler_csv(texto), "odk", odk=credenciais)
                st.rerun()
            except requests.HTTPError as e:
                codigo = e.response.status_code if e.response is not None else None
                if codigo in (401, 403):
                    st.error("O ODK recusou e-mail/senha (ou o usuário não tem acesso a este projeto).")
                elif codigo == 404:
                    st.error("Projeto ou formulário não encontrado -- confira os IDs.")
                else:
                    st.error(f"O ODK respondeu com erro: {e}")
            except requests.RequestException as e:
                st.error(f"Não foi possível falar com o ODK Central: {e}")

with aba_arquivo:
    st.caption("Use o ZIP exportado pelo ODK Central (CSV + pasta media) para gerar "
               "os relatórios com as fotos, sem precisar de senha. Só o CSV também "
               "funciona, mas aí os relatórios saem sem foto.")
    arquivo = st.file_uploader("CSV ou ZIP exportado do ODK", type=["csv", "zip"])
    if arquivo is not None and ss.get("arquivo_carregado") != (arquivo.name, arquivo.size):
        try:
            if arquivo.name.lower().endswith(".zip"):
                texto, midia = core.ler_zip_odk(arquivo.getvalue())
            else:
                texto, midia = arquivo.getvalue().decode("utf-8-sig"), {}
            carregar(core.ler_csv(texto), "arquivo", midia=midia)
            ss.arquivo_carregado = (arquivo.name, arquivo.size)
            st.rerun()
        except Exception as e:
            st.error(f"Não consegui ler o arquivo: {e}")

planilha: core.Planilha | None = ss.get("planilha")
if planilha is None:
    st.info("Busque os dados no ODK Central ou envie um arquivo exportado para começar.")
    st.stop()

if not planilha.linhas:
    st.warning("O CSV não tem nenhum registro.")
    st.stop()

# ── Dados carregados ───────────────────────────────────────────────────────
st.divider()
fonte = "ODK Central" if ss.fonte == "odk" else "arquivo enviado"
c1, c2 = st.columns([3, 1])
c1.success(f"**{len(planilha.linhas)}** registros carregados de **{fonte}**"
           + (f" ({len(ss.midia)} arquivos de mídia no ZIP)" if ss.midia else ""))
c2.download_button("⬇️ Baixar CSV", planilha.texto.encode("utf-8"), "dados_odk.csv",
                   "text/csv", use_container_width=True)
if ss.fonte == "odk" and planilha.col_instancia is None:
    st.warning("O CSV não tem a coluna KEY/meta-instanceID -- não dá para buscar as fotos.")

# ── Escolha dos registros ──────────────────────────────────────────────────
st.subheader("📋 Registros")
colunas = [core.COL_ID_PROJ, core.COL_TIPO_REL, core.COL_DATA, core.COL_CIDADE,
           core.COL_PROCESSO_SEI, core.COL_TIPO_PROJ]
nomes = []
for c in colunas:
    nome = planilha.cabecalho[c] if c < len(planilha.cabecalho) else f"coluna {c + 1}"
    nomes.append(nome if nome not in nomes else f"{nome} ({c + 1})")
tabela = pd.DataFrame(
    [[i + 1] + [core.campo(l, c) for c in colunas] for i, l in enumerate(planilha.linhas)],
    columns=["#"] + nomes,
)
tabela[nomes[2]] = tabela[nomes[2]].map(core.formatar_data)

modo = st.radio("Quais relatórios gerar?", ["Todos", "Escolher"], horizontal=True)
evento = st.dataframe(
    tabela, hide_index=True, use_container_width=True, height=380,
    on_select="rerun" if modo == "Escolher" else "ignore",
    selection_mode="multi-row", key="tabela_registros",
)

if modo == "Todos":
    posicoes = list(range(len(planilha.linhas)))
else:
    st.caption("Marque as linhas na tabela acima, ou digite os números da coluna #.")
    digitado = st.text_input("Números (ex: 1, 3, 5-10)", key="faixa")
    try:
        posicoes = set(evento.selection.rows) | faixa(digitado, len(planilha.linhas))
    except ValueError:
        st.error("Formato inválido -- use números separados por vírgula e faixas como 5-10.")
        posicoes = set(evento.selection.rows)
    posicoes = sorted(posicoes)
st.caption(f"**{len(posicoes)}** relatório(s) selecionado(s).")

# ── Modelo e geração ───────────────────────────────────────────────────────
st.subheader("📄 Modelo do relatório")
modelo = st.file_uploader("Modelo DOCX (formulario.docx)", type=["docx"])

c1, c2 = st.columns(2)
gerar = c1.button(f"🚀 Gerar {len(posicoes)} relatório(s)", type="primary",
                  use_container_width=True, disabled=not (modelo and posicoes))
juntar_fotos = c2.button("📦 Preparar ZIP com as fotos originais", use_container_width=True,
                         disabled=not posicoes)
if not modelo:
    c1.caption("Envie o modelo DOCX para habilitar.")

if gerar:
    try:
        achadas, faltando = garantir_fotos(posicoes)
        modelo_bytes = modelo.getvalue()
        buf = io.BytesIO()
        barra = st.progress(0.0, text="Gerando relatórios...")
        # DOCX ja e ZIP comprimido; comprimir de novo so gasta CPU
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
            for i, p in enumerate(posicoes, 1):
                zf.writestr(f"{core.numero_relatorio(p)}.docx",
                            core.gerar_docx(modelo_bytes, planilha, p, ss.fotos))
                barra.progress(i / len(posicoes), text=f"Relatórios: {i}/{len(posicoes)}")
        barra.empty()
        ss.zip_relatorios = buf.getvalue()
        ss.zip_relatorios_info = (len(posicoes), achadas, faltando,
                                  datetime.now().strftime("%H:%M:%S"))
    except Exception as e:
        st.error(f"Erro ao gerar os relatórios: {e}")
        st.exception(e)

if juntar_fotos:
    try:
        ss.zip_fotos = zip_fotos_originais(posicoes)
    except Exception as e:
        st.error(f"Erro ao juntar as fotos: {e}")

if ss.get("zip_relatorios"):
    qtd, achadas, faltando, hora = ss.zip_relatorios_info
    aviso = f"✅ {qtd} relatório(s) gerado(s) às {hora} — {achadas} foto(s)"
    if faltando:
        aviso += f", {faltando} não encontrada(s) (saíram com 'sem imagem')"
    st.success(aviso)
    st.download_button("📥 Baixar relatórios (ZIP)", ss.zip_relatorios,
                       "relatorios_vistoria.zip", "application/zip",
                       type="primary", use_container_width=True)

if ss.get("zip_fotos"):
    st.download_button("📥 Baixar fotos originais (ZIP)", ss.zip_fotos,
                       "imagens_odk.zip", "application/zip", use_container_width=True)

st.divider()
st.caption("SEPE — Gerador de Relatórios de Vistoria · versão 3.0")
