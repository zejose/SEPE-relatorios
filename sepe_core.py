"""Tudo que nao e tela: ODK Central, leitura do CSV, fotos e montagem dos DOCX.

Separado do app.py para dar para testar sem abrir o Streamlit, e para o app
so cuidar de botao e mensagem.
"""

from __future__ import annotations

import csv
import io
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Iterator

import requests
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from PIL import Image, ImageDraw, ImageOps

# Posicao de cada campo no CSV exportado pelo ODK (0 = primeira coluna). O
# modelo DOCX depende dessa ordem; se o formulario mudar, e so aqui.
COL_DATA = 1
COL_ID_PROJ = 4
COL_PROCESSO_SEI = 5
COL_TIPO_REL = 7
COL_POSSUI_PLACA = 8
COL_PLANO_TRABALHO = 9
COL_CIDADE = 11
COL_LAT = 12
COL_LONG = 13
COL_TIPO_PROJ = 16
COLS_IMAGENS = (17, 18, 19, 20, 21)
COL_OBSERVACAO = 22
COL_META = 24
COL_RESPONSAVEL = 26

COLUNAS_INSTANCIA = ("KEY", "meta-instanceID", "instanceID")

# 12 cm de largura no DOCX a ~300 dpi; foto de celular chega com 4000 px e
# 4 MB, e sem reduzir cada relatorio passava de 20 MB e estourava a memoria
# do Streamlit Cloud (1 GB) em poucas dezenas de relatorios.
LADO_MAX_FOTO = 1400
QUALIDADE_JPEG = 82
LARGURA_FOTO_CM = 12
LARGURA_SEM_FOTO_CM = 3

TIMEOUT = 60


@dataclass
class Planilha:
    texto: str
    cabecalho: list[str]
    linhas: list[list[str]]
    col_instancia: int | None


def ler_csv(texto: str) -> Planilha:
    # csv.reader sobre o texto inteiro, e nao linha a linha: observacao com
    # quebra de linha vem entre aspas e ocupa mais de uma linha do arquivo
    registros = list(csv.reader(io.StringIO(texto)))
    if not registros:
        return Planilha(texto, [], [], None)
    cabecalho = [c.strip() for c in registros[0]]
    linhas = [r for r in registros[1:] if any(c.strip() for c in r)]
    col_instancia = next(
        (cabecalho.index(c) for c in COLUNAS_INSTANCIA if c in cabecalho), None
    )
    return Planilha(texto, cabecalho, linhas, col_instancia)


def campo(linha: list[str], indice: int | None) -> str:
    if indice is None or indice >= len(linha):
        return ""
    return linha[indice].strip()


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
        desenho.text((200, 150), "SEM IMAGEM", fill=(120, 120, 120), anchor="mm")
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
        return dt.strftime("%d-%m-%Y")
    except ValueError:
        return valor


def numero_relatorio(posicao: int) -> int:
    """Numero que aparece no relatorio e no nome do arquivo.

    E a linha da planilha (cabecalho = linha 1), como a versao anterior
    fazia ao passar o CSV pelo Excel: o 1o registro e o relatorio 2.
    """
    return posicao + 2


def pedidos_de_fotos(planilha: Planilha, posicoes: Iterable[int]) -> set[tuple[str, str]]:
    pedidos = set()
    for p in posicoes:
        linha = planilha.linhas[p]
        instancia = campo(linha, planilha.col_instancia)
        for c in COLS_IMAGENS:
            nome = campo(linha, c)
            if nome:
                pedidos.add((instancia, nome))
    return pedidos


def gerar_docx(
    modelo: bytes, planilha: Planilha, posicao: int, fotos: dict[tuple[str, str], bytes | None]
) -> bytes:
    linha = planilha.linhas[posicao]
    instancia = campo(linha, planilha.col_instancia)
    doc = DocxTemplate(io.BytesIO(modelo))

    def imagem(coluna: int) -> InlineImage:
        nome = campo(linha, coluna)
        foto = fotos.get((instancia, nome)) if nome else None
        if foto:
            return InlineImage(doc, io.BytesIO(foto), Cm(LARGURA_FOTO_CM))
        return InlineImage(doc, io.BytesIO(foto_padrao()), Cm(LARGURA_SEM_FOTO_CM))

    contexto = {
        "relatorio": numero_relatorio(posicao),
        "data": formatar_data(campo(linha, COL_DATA)),
        "id_proj": campo(linha, COL_ID_PROJ),
        "processo_sei": campo(linha, COL_PROCESSO_SEI),
        "id_tipo_rel": campo(linha, COL_TIPO_REL),
        "possui_placa": campo(linha, COL_POSSUI_PLACA),
        "plano_trabalho": campo(linha, COL_PLANO_TRABALHO),
        "cidade": campo(linha, COL_CIDADE),
        "lat": campo(linha, COL_LAT),
        "long": campo(linha, COL_LONG),
        "tipo_proj": campo(linha, COL_TIPO_PROJ),
        "observacao": campo(linha, COL_OBSERVACAO),
        "meta": campo(linha, COL_META),
        "responsavel": campo(linha, COL_RESPONSAVEL),
    }
    for i, coluna in enumerate(COLS_IMAGENS, 1):
        contexto[f"imagem_{i}"] = imagem(coluna)

    doc.render(contexto)
    saida = io.BytesIO()
    doc.save(saida)
    return saida.getvalue()
