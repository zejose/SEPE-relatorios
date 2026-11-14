"""Tudo que nao e tela: ODK Central, leitura do CSV, fotos e montagem dos DOCX.

Separado do app.py para dar para testar sem abrir o Streamlit, e para o app
so cuidar de botao e mensagem.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Iterator

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage, Listing
from PIL import Image, ImageDraw, ImageFont, ImageOps

# um modelo por tipo de vistoria (campo "Tipo Relatorio" do formulario): o de
# acompanhamento tem placa e plano de trabalho, que a inicial nao tem
_MODELOS = Path(__file__).parent / "modelos"
MODELOS = {
    "inicial": _MODELOS / "relatorio_visita_inicial.docx",
    "acompanhamento": _MODELOS / "relatorio_visita_acompanhamento.docx",
}
TIPO_PADRAO = "inicial"

# Campo do modelo DOCX -> coluna do CSV do formulario "Vistoria - Adote uma
# Praca - SEPE" (projeto 4). Por nome, e nao por posicao: campo novo no
# formulario desloca as colunas, e o relatorio saia com dado trocado.
CAMPOS = {
    "data": "today",
    "id_proj": "details-N_mero_ID",
    "processo_sei": "details-Processo_sei",
    "adotante": "details-Nome_do_adotante",
    "id_tipo_rel": "details-Tipo_Relat_rio",
    "possui_placa": "details-placa",
    "plano_trabalho": "details-Cumprindo_plano_de_trabalho",
    "obs_plano_trabalho": "details-Observa_o_Plano_de_Trabalho",
    "cidade": "details-regiao_adm",
    "lat": "details-location-Latitude",
    "long": "details-location-Longitude",
    "tipo_proj": "details-Tipo_de_Projeto",
    "observacao": "Observa_es",
    "meta": "meta-instanceID",
    "responsavel": "SubmitterName",
}
# imagem_1 e o croqui (planta de localizacao), imagem_2 a imagem_5 as fotos
IMAGENS = (
    "details-Planta_de_Localizacao",
    "details-Foto_1",
    "details-Foto_2",
    "details-Foto_3",
    "details-Foto_4",
)
COLUNAS_INSTANCIA = ("KEY", "meta-instanceID", "instanceID")

# Foto de celular chega com 4000 px e 4 MB; sem reduzir cada relatorio
# passava de 20 MB e estourava a memoria do Streamlit Cloud (1 GB) em poucas
# dezenas de relatorios. 1280 px ainda e ~230 dpi no tamanho impresso.
LADO_MAX_FOTO = 1280
QUALIDADE_JPEG = 80
# o espaco de cada foto no modelo (o croqui la tem 14,5 x 9,5 cm); foto em pe
# e limitada pela altura -- com 10 cm o croqui ja nao cabia abaixo do titulo
# e pulava de pagina, deixando meia pagina em branco
FOTO_MAX_LARGURA_CM = 14.0
FOTO_MAX_ALTURA_CM = 9.0
LARGURA_SEM_FOTO_CM = 3

TIMEOUT = 60


@dataclass
class Planilha:
    texto: str
    cabecalho: list[str]
    linhas: list[list[str]]

    def __post_init__(self) -> None:
        self._indice = {nome: i for i, nome in reversed(list(enumerate(self.cabecalho)))}
        self.col_instancia = next(
            (self._indice[c] for c in COLUNAS_INSTANCIA if c in self._indice), None
        )

    def valor(self, posicao: int, coluna: str) -> str:
        i = self._indice.get(coluna)
        linha = self.linhas[posicao]
        return linha[i].strip() if i is not None and i < len(linha) else ""

    def instancia(self, posicao: int) -> str:
        i = self.col_instancia
        linha = self.linhas[posicao]
        return linha[i].strip() if i is not None and i < len(linha) else ""

    def colunas_faltando(self) -> list[str]:
        return [c for c in (*CAMPOS.values(), *IMAGENS) if c not in self._indice]


def ler_csv(texto: str) -> Planilha:
    # csv.reader sobre o texto inteiro, e nao linha a linha: observacao com
    # quebra de linha vem entre aspas e ocupa mais de uma linha do arquivo
    registros = list(csv.reader(io.StringIO(texto)))
    if not registros:
        return Planilha(texto, [], [])
    cabecalho = [c.strip() for c in registros[0]]
    linhas = [r for r in registros[1:] if any(c.strip() for c in r)]
    return Planilha(texto, cabecalho, linhas)


def ler_zip_odk(conteudo: bytes) -> tuple[str, dict[str, bytes]]:
    """CSV principal e fotos de um ZIP exportado pelo ODK Central.

    O principal e o de nome mais curto: os de grupos repetidos vem como
    "<formulario>-<grupo>.csv".
    """
    with zipfile.ZipFile(io.BytesIO(conteudo)) as zf:
        nomes_csv = sorted((n for n in zf.namelist() if n.lower().endswith(".csv")), key=len)
        if not nomes_csv:
            raise ValueError("O ZIP não tem nenhum arquivo CSV.")
        texto = zf.read(nomes_csv[0]).decode("utf-8-sig")
        midia = {
            n.rsplit("/", 1)[-1]: zf.read(n)
            for n in zf.namelist()
            if not n.endswith("/") and not n.lower().endswith(".csv")
        }
    return texto, midia


class OdkCentral:
    def __init__(self, url: str, projeto: str, formulario: str, email: str, senha: str):
        self.base = f"{url.rstrip('/')}/v1/projects/{projeto}/forms/{formulario}"
        self.sessao = requests.Session()
        self.sessao.auth = (email, senha)

    def csv(self) -> str:
        # attachments=false: sem isso o ZIP traz TODAS as fotos de todas as
        # submissoes, e o app baixava tudo so para ler o CSV -- e depois
        # baixava as fotos de novo, uma a uma
        r = self.sessao.get(
            f"{self.base}/submissions.csv.zip",
            params={"attachments": "false"},
            timeout=TIMEOUT * 3,
        )
        r.raise_for_status()
        texto, _ = ler_zip_odk(r.content)
        return texto

    def anexo(self, instancia: str, nome: str) -> bytes | None:
        r = self.sessao.get(
            f"{self.base}/submissions/{instancia}/attachments/{nome}", timeout=TIMEOUT
        )
        if r.status_code != 200:
            return None
        return r.content

    def anexos(
        self, pedidos: Iterable[tuple[str, str]], *, reduzir: bool, workers: int = 8
    ) -> Iterator[tuple[tuple[str, str], bytes | None]]:
        """Baixa em paralelo e devolve cada um assim que chega, para a tela
        poder mostrar o progresso (o Streamlit so desenha da thread principal)."""

        def _um(pedido: tuple[str, str]) -> bytes | None:
            bruto = self.anexo(*pedido)
            if bruto is None:
                return None
            return reduzir_foto(bruto) if reduzir else bruto

        with ThreadPoolExecutor(max_workers=workers) as ex:
            futuros = {ex.submit(_um, p): p for p in pedidos}
            for fut in as_completed(futuros):
                try:
                    yield futuros[fut], fut.result()
                except Exception:
                    yield futuros[fut], None


def reduzir_foto(bruto: bytes) -> bytes | None:
    """JPEG menor e ja na posicao certa -- ou None se nao for imagem.

    exif_transpose: o celular grava a foto deitada e so marca no EXIF que ela
    deve ser girada; o Word ignora essa marca, e a foto saia de lado.
    """
    try:
        with Image.open(io.BytesIO(bruto)) as img:
            img = ImageOps.exif_transpose(img).convert("RGB")
            img.thumbnail((LADO_MAX_FOTO, LADO_MAX_FOTO))
            saida = io.BytesIO()
            img.save(saida, "JPEG", quality=QUALIDADE_JPEG, optimize=True)
            return saida.getvalue()
    except Exception:
        return None


_SEM_FOTO: bytes | None = None


def foto_padrao() -> bytes:
    """Gerada aqui mesmo: antes vinha de um banco de imagens na internet, e o
    relatorio dependia de um site de terceiros estar no ar."""
    global _SEM_FOTO
    if _SEM_FOTO is None:
        img = Image.new("RGB", (400, 300), (235, 235, 235))
        desenho = ImageDraw.Draw(img)
        desenho.rectangle((8, 8, 391, 291), outline=(170, 170, 170), width=3)
        fonte = ImageFont.load_default(size=44)
        desenho.text((200, 150), "SEM IMAGEM", fill=(110, 110, 110), anchor="mm", font=fonte)
        saida = io.BytesIO()
        img.save(saida, "JPEG", quality=85)
        _SEM_FOTO = saida.getvalue()
    return _SEM_FOTO


def formatar_data(valor: str) -> str:
    if not valor:
        return valor
    try:
        if "T" in valor:
            dt = datetime.fromisoformat(valor.replace("Z", "+00:00"))
        else:
            dt = datetime.strptime(valor[:10], "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except ValueError:
        return valor


# perguntas de multipla escolha: o ODK grava o "name" da opcao, sem espaco e
# sem acento ("Plano_Piloto", "n_o" para "nao")
CAMPOS_DE_ESCOLHA = ("cidade", "id_tipo_rel", "tipo_proj", "possui_placa", "plano_trabalho")
ROTULOS_DE_ESCOLHA = {"n_o": "não"}


def texto_de_escolha(valor: str) -> str:
    return ROTULOS_DE_ESCOLHA.get(valor, valor.replace("_", " "))


def tipo_de_vistoria(planilha: Planilha, posicao: int) -> str:
    tipo = planilha.valor(posicao, CAMPOS["id_tipo_rel"]).lower()
    return tipo if tipo in MODELOS else TIPO_PADRAO


def observacao_completa(geral: str, plano: str) -> Listing:
    """Observacao geral e, abaixo, a do plano de trabalho (so o acompanhamento
    pergunta) -- no mesmo campo do modelo. Listing mantem as quebras de linha,
    que como texto simples viravam espaco no Word."""
    partes = [geral] if geral else []
    if plano:
        partes.append(f"Plano de trabalho: {plano}")
    return Listing("\n".join(partes))


def formatar_coordenada(valor: str) -> str:
    try:
        return f"{float(valor):.6f}"
    except ValueError:
        return valor


def numero_relatorio(posicao: int) -> int:
    """Numero que aparece no relatorio e no nome do arquivo -- o mesmo da
    coluna # da tabela na tela."""
    return posicao + 1


def nome_arquivo(planilha: Planilha, posicao: int) -> str:
    """'<#>_ID<id>.docx' -- o # mantem a ordem da tabela e o ID diz de qual
    projeto e sem precisar abrir."""
    id_proj = "".join(c for c in planilha.valor(posicao, CAMPOS["id_proj"]) if c.isalnum())
    sufixo = f"_ID{id_proj}" if id_proj else ""
    return f"{numero_relatorio(posicao)}{sufixo}.docx"


def pedidos_de_fotos(planilha: Planilha, posicoes: Iterable[int]) -> set[tuple[str, str]]:
    pedidos = set()
    for p in posicoes:
        instancia = planilha.instancia(p)
        for coluna in IMAGENS:
            nome = planilha.valor(p, coluna)
            if nome:
                pedidos.add((instancia, nome))
    return pedidos


_TWIPS_POR_CM = 567
_MARGEM_CELULA_CM = 0.4


@lru_cache(maxsize=8)
def larguras_no_modelo(modelo: bytes) -> dict[str, float]:
    """Largura (cm) da celula de tabela onde cada {{imagem_N}} esta no modelo.

    O modelo de acompanhamento poe as fotos numa tabela de 2 colunas (~7 cm
    cada); com a largura fixa de 14 cm, foto deitada estourava a celula.
    Imagem fora de tabela nao entra aqui e usa FOTO_MAX_LARGURA_CM."""
    larguras = {}
    for t in Document(io.BytesIO(modelo)).element.body.iter(qn("w:t")):
        achado = re.search(r"imagem_\d", t.text or "")
        if not achado:
            continue
        celula = next((a for a in t.iterancestors() if a.tag == qn("w:tc")), None)
        if celula is None:
            continue
        twips = _largura_celula(celula)
        if twips:
            larguras[achado.group(0)] = twips / _TWIPS_POR_CM - _MARGEM_CELULA_CM
    return larguras


def _largura_celula(celula) -> float | None:
    """tcW da celula; sem ele (Google Docs nao grava), a coluna da grade."""
    tc_w = celula.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
    if tc_w is not None and tc_w.get(qn("w:type")) == "dxa":
        return float(tc_w.get(qn("w:w")))
    linha = celula.getparent()
    tabela = linha.getparent()
    grade = [float(c.get(qn("w:w"))) for c in tabela.findall(f"{qn('w:tblGrid')}/{qn('w:gridCol')}")]
    coluna = 0
    for tc in linha.findall(qn("w:tc")):
        span = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
        n = int(span.get(qn("w:val"))) if span is not None else 1
        if tc is celula:
            return sum(grade[coluna:coluna + n]) or None
        coluna += n
    return None


def _tamanho_foto(foto: bytes, largura_max_cm: float) -> dict:
    """Largura OU altura (o que limitar primeiro) para caber no espaco do modelo."""
    largura_max_cm = min(largura_max_cm, FOTO_MAX_LARGURA_CM)
    with Image.open(io.BytesIO(foto)) as img:
        largura, altura = img.size
    if largura / altura >= largura_max_cm / FOTO_MAX_ALTURA_CM:
        return {"width": Cm(largura_max_cm)}
    return {"height": Cm(FOTO_MAX_ALTURA_CM)}


def gerar_docx(
    modelo: bytes, planilha: Planilha, posicao: int, fotos: dict[tuple[str, str], bytes | None]
) -> bytes:
    instancia = planilha.instancia(posicao)
    doc = DocxTemplate(io.BytesIO(modelo))
    larguras = larguras_no_modelo(modelo)

    def imagem(i: int, coluna: str) -> InlineImage:
        nome = planilha.valor(posicao, coluna)
        foto = fotos.get((instancia, nome)) if nome else None
        if foto:
            largura = larguras.get(f"imagem_{i}", FOTO_MAX_LARGURA_CM)
            return InlineImage(doc, io.BytesIO(foto), **_tamanho_foto(foto, largura))
        return InlineImage(doc, io.BytesIO(foto_padrao()), Cm(LARGURA_SEM_FOTO_CM))

    contexto = {campo: planilha.valor(posicao, coluna) for campo, coluna in CAMPOS.items()}
    for campo in CAMPOS_DE_ESCOLHA:
        contexto[campo] = texto_de_escolha(contexto[campo])
    contexto["relatorio"] = numero_relatorio(posicao)
    contexto["observacao"] = observacao_completa(
        contexto["observacao"], contexto["obs_plano_trabalho"]
    )
    contexto["data"] = formatar_data(contexto["data"])
    contexto["lat"] = formatar_coordenada(contexto["lat"])
    contexto["long"] = formatar_coordenada(contexto["long"])
    for i, coluna in enumerate(IMAGENS, 1):
        contexto[f"imagem_{i}"] = imagem(i, coluna)

    # autoescape: observacao digitada com "&" ou "<" ia crua para o XML do
    # DOCX e o Word recusava abrir o arquivo
    doc.render(contexto, autoescape=True)
    saida = io.BytesIO()
    doc.save(saida)
    return saida.getvalue()
