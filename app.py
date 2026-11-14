import io
import zipfile
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

import sepe_core as core

st.set_page_config(page_title="Gerador de Relatórios SEPE", page_icon="🏗️", layout="wide")
ss = st.session_state
POR_PAGINA = 100


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
    # fotos ja baixadas, linhas marcadas e ZIPs prontos sao da carga anterior
    ss.fotos = {}
    ss.marcadas = {}
    ss.pop("pagina", None)
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
        instancia = planilha.instancia(p)
        id_proj = planilha.valor(p, core.CAMPOS["id_proj"])
        for coluna in core.IMAGENS:
            nome = planilha.valor(p, coluna)
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
faltando = planilha.colunas_faltando()
if faltando:
    st.warning("Este CSV não tem as colunas " + ", ".join(faltando) + " -- esses campos "
               "saem em branco no relatório. É o formulário 'Vistoria - Adote uma Praça'?")

# ── Escolha dos registros ──────────────────────────────────────────────────
st.subheader("📋 Registros")
colunas_tabela = {
    "ID": "id_proj", "Adotante": "adotante", "Tipo": "id_tipo_rel", "Data": "data",
    "Cidade": "cidade", "Processo SEI": "processo_sei", "Responsável": "responsavel",
}
tabela = pd.DataFrame(
    [
        [i + 1] + [planilha.valor(i, core.CAMPOS[c]) for c in colunas_tabela.values()]
        for i in range(len(planilha.linhas))
    ],
    columns=["#", *colunas_tabela],
)
tabela["Data"] = tabela["Data"].map(core.formatar_data)
tabela["Cidade"] = tabela["Cidade"].map(core.texto_de_escolha)

modo = st.radio("Quais relatórios gerar?", ["Todos", "Escolher"], horizontal=True)

total = len(planilha.linhas)
paginas = max(1, -(-total // POR_PAGINA))
rotulos = {p: f"{(p - 1) * POR_PAGINA + 1}–{min(p * POR_PAGINA, total)}"
           for p in range(1, paginas + 1)}
if paginas > 1:
    pagina = st.radio(f"Página ({POR_PAGINA} por página)", list(rotulos),
                      format_func=rotulos.get, horizontal=True, key="pagina")
else:
    pagina = 1
inicio = (pagina - 1) * POR_PAGINA

# A selecao de cada pagina fica guardada aqui: a tabela de uma pagina que nao
# esta na tela deixa de existir, e com ela as linhas marcadas nela.
marcadas: dict[int, set[int]] = ss.setdefault("marcadas", {})


def _guardar_marcadas(pag: int = pagina, ini: int = inicio) -> None:
    linhas = ss[f"tabela_{pag}"].selection.rows
    marcadas[pag] = {ini + r for r in linhas}


st.dataframe(
    tabela.iloc[inicio:inicio + POR_PAGINA], hide_index=True, use_container_width=True,
    height=380, selection_mode="multi-row", key=f"tabela_{pagina}",
    on_select=_guardar_marcadas if modo == "Escolher" else "ignore",
)

if modo == "Todos":
    posicoes = list(range(total))
else:
    st.caption("Marque as linhas na tabela (em qualquer página), ou digite os números da coluna #.")
    digitado = st.text_input("Números (ex: 1, 3, 5-10)", key="faixa")
    na_tabela = set().union(*marcadas.values())
    try:
        posicoes = na_tabela | faixa(digitado, total)
    except ValueError:
        st.error("Formato inválido -- use números separados por vírgula e faixas como 5-10.")
        posicoes = na_tabela
    posicoes = sorted(posicoes)
    if na_tabela:
        c1, c2 = st.columns([4, 1])
        c1.caption("Marcados na tabela: " + ", ".join(str(p + 1) for p in sorted(na_tabela)))
        if c2.button("Limpar marcados", use_container_width=True):
            marcadas.clear()
            for p in range(1, paginas + 1):
                ss.pop(f"tabela_{p}", None)
            st.rerun()
st.caption(f"**{len(posicoes)}** relatório(s) selecionado(s).")

# ── Modelo e geração ───────────────────────────────────────────────────────
st.subheader("📄 Modelo do relatório")
AUTOMATICO = "Automático — pelo tipo de vistoria de cada registro"
opcoes_modelo = {AUTOMATICO: None, **{f"{t.capitalize()} para todos": t for t in core.MODELOS}}
escolha_modelo = opcoes_modelo[st.radio("Qual modelo usar?", list(opcoes_modelo), key="escolha_modelo")]


def modelo_de(posicao: int) -> str:
    return escolha_modelo or core.tipo_de_vistoria(planilha, posicao)


st.caption("Os dois modelos já vêm no app; envie outro DOCX só se quiser trocar o layout.")
modelos: dict[str, bytes] = {}
for tipo, coluna in zip(core.MODELOS, st.columns(len(core.MODELOS))):
    with coluna:
        qtd = sum(1 for p in posicoes if modelo_de(p) == tipo)
        st.markdown(f"**{tipo.capitalize()}** — usado em {qtd} relatório(s)")
        enviado = st.file_uploader(f"Trocar modelo ({tipo})", type=["docx"], key=f"modelo_{tipo}",
                                   label_visibility="collapsed")
        if enviado:
            modelos[tipo] = enviado.getvalue()
            st.caption(f"Usando o enviado: **{enviado.name}**")
        else:
            modelos[tipo] = core.MODELOS[tipo].read_bytes()
            st.download_button("⬇️ Baixar modelo atual", modelos[tipo], core.MODELOS[tipo].name,
                               key=f"baixar_modelo_{tipo}")

c1, c2 = st.columns(2)
gerar = c1.button(f"🚀 Gerar {len(posicoes)} relatório(s)", type="primary",
                  use_container_width=True, disabled=not posicoes)
juntar_fotos = c2.button(
    "📦 Preparar ZIP com as fotos originais", use_container_width=True, disabled=not posicoes,
    help="Só as fotos dos registros selecionados, em tamanho original (como saíram do "
         "celular), num ZIP à parte -- para arquivar ou anexar no SEI. Nos relatórios as "
         "fotos vão reduzidas.")

if gerar:
    try:
        achadas, faltando = garantir_fotos(posicoes)
        buf = io.BytesIO()
        barra = st.progress(0.0, text="Gerando relatórios...")
        # DOCX ja e ZIP comprimido; comprimir de novo so gasta CPU
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
            for i, p in enumerate(posicoes, 1):
                modelo = modelos[modelo_de(p)]
                zf.writestr(core.nome_arquivo(planilha, p),
                            core.gerar_docx(modelo, planilha, p, ss.fotos))
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
