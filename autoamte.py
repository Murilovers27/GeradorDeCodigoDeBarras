"""
Gerador de etiquetas
--------------------
Preenche o modelo Word (modelo_etiqueta.docx), gera o código de barras
Code 128 aqui mesmo (sem internet e sem token de API) e permite
salvar, abrir ou imprimir o resultado.

Instalação (uma vez):
    pip install docxtpl docxcompose python-barcode pillow

Uso:
    python gerar_etiqueta.py
"""
import io
import os
import re
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import barcode
from docx import Document
from docx.shared import Cm
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, InlineImage
from PIL import Image, ImageDraw, ImageFont

if getattr(sys, "frozen", False):
    PASTA = Path(sys.executable).resolve().parent
    RECURSOS = Path(getattr(sys, "_MEIPASS", PASTA))
else:
    PASTA = Path(__file__).resolve().parent
    RECURSOS = PASTA
MODELO = PASTA / "modelo_etiqueta.docx"

SAIDA = PASTA / "saida"

# Valores sugeridos no terminal (Enter aceita o sugerido).
# Durante o uso, o programa passa a sugerir os últimos valores digitados.
PADRAO = {"titulo": "AUEN", "rotulo": "CENTRO", "numero": "1062"}

LARGURA_CODIGO_CM = 21.67  # mesma largura da imagem do modelo original
LARGURA_PX = 2560          # resolução da imagem (~300 dpi nessa largura)

# Fonte do número que aparece embaixo das barras (usa a primeira que existir)
FONTES = [
    "C:/Windows/Fonts/arial.ttf",
    "/Library/Fonts/Arial.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]


# ---------------------------------------------------------------- geração ---
def caminho_recurso(nome: str) -> Path:
    """Arquivo embutido no .exe (ou ao lado do script, quando roda como .py)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / nome


# Se existir um modelo ao lado do .exe, ele tem prioridade; senão usa o embutido
MODELO = PASTA / "modelo_etiqueta.docx"
if not MODELO.exists():
    MODELO = caminho_recurso("modelo_etiqueta.docx")
def _fonte(tamanho: int):
    for caminho in FONTES:
        if os.path.exists(caminho):
            return ImageFont.truetype(caminho, tamanho)
    return ImageFont.load_default(tamanho)  # Pillow >= 10.1


def criar_codigo_barras(texto: str) -> io.BytesIO:
    """Desenha um Code 128 (barras + texto embaixo) como PNG em memória."""
    # build() devolve o padrão de barras como texto: "1101001..." (1 = barra preta)
    modulos = barcode.get("code128", texto).build()[0]

    px = max(1, LARGURA_PX // len(modulos))  # pixels por módulo; inteiro = barras nítidas
    largura = px * len(modulos)
    alt_barras = int(largura * 0.36)
    tam_fonte = int(largura * 0.075 * 1.10)
    altura = alt_barras + int(tam_fonte * 1.35)

    img = Image.new("L", (largura, altura), 255)
    desenho = ImageDraw.Draw(img)
    for i, bit in enumerate(modulos):
        if bit == "1":
            desenho.rectangle([i * px, 0, (i + 1) * px - 1, alt_barras], fill=0)
    desenho.text(
        (largura / 2, alt_barras + tam_fonte * 0.1),
        texto, font=_fonte(tam_fonte), fill=0, anchor="ma",
    )

    buffer = io.BytesIO()
    img.save(buffer, format="PNG", dpi=(300, 300))
    buffer.seek(0)
    return buffer


def _renderizar_documento(titulo: str, codigo: str, rotulo: str, numero: str, destino: Path) -> None:
    doc = DocxTemplate(MODELO)
    imagem = InlineImage(doc, criar_codigo_barras(codigo), width=Cm(LARGURA_CODIGO_CM))
    doc.render(
        {"titulo": titulo, "codigo_barras": imagem, "rotulo": rotulo, "numero": numero},
        autoescape=True,
    )

    doc.save(destino)


def montar_documento(
    titulo: str, codigo: str, rotulo: str, numero: str, pasta_saida: Path | None = None
) -> Path:
    """Preenche o modelo e salva em ./saida. Devolve o caminho do arquivo."""
    pasta = Path(pasta_saida) if pasta_saida else SAIDA
    pasta.mkdir(parents=True, exist_ok=True)
    nome = re.sub(r"[^\w\-]+", "_", codigo)
    destino = pasta / f"etiqueta_{nome}.docx"
    try:
        _renderizar_documento(titulo, codigo, rotulo, numero, destino)
    except PermissionError:  # o arquivo antigo está aberto no Word
        destino = pasta / f"etiqueta_{nome}_{datetime.now():%H%M%S}.docx"
        _renderizar_documento(titulo, codigo, rotulo, numero, destino)
    return destino


def gerar_codigos_lote(
    prateleira: str, coluna_inicial: str, coluna_final: str
) -> list[str]:
    """Gera etiquetas da prateleira escolhida, por coluna e andar."""
    try:
        numero_prateleira = int(prateleira)
    except ValueError as erro:
        raise ValueError("a prateleira deve ser 1, 2 ou 3") from erro
    if numero_prateleira not in range(1, 4):
        raise ValueError("a prateleira deve ser 1, 2 ou 3")
    inicio = ord(coluna_inicial.upper()) - ord("A")
    fim = ord(coluna_final.upper()) - ord("A")
    if not 0 <= inicio <= fim <= ord("R") - ord("A"):
        raise ValueError("as colunas devem estar entre A e R, em ordem")
    return [
        f"{numero_prateleira:02d}-{andar}-{chr(ord('A') + coluna)}{andar}"
        for coluna in range(inicio, fim + 1)
        for andar in range(1, 6)
    ]


def montar_lote(
    titulo: str,
    codigos: list[str],
    rotulo: str,
    numero: str,
    pasta_saida: Path | None = None,
) -> Path:
    """Monta todas as etiquetas em um único documento Word."""
    pasta = Path(pasta_saida) if pasta_saida else SAIDA
    pasta.mkdir(parents=True, exist_ok=True)
    prateleira = codigos[0].split("-")[0]
    nome = (
        f"etiquetas_{prateleira}_"
        f"{codigos[0].split('-')[2][0]}_{codigos[-1].split('-')[2][0]}"
    )
    destino = pasta / f"{nome}.docx"

    with tempfile.TemporaryDirectory() as pasta_temporaria:
        primeiro = Path(pasta_temporaria) / "etiqueta_1.docx"
        _renderizar_documento(titulo, codigos[0], rotulo, numero, primeiro)
        documento = Composer(Document(primeiro))
        for indice, codigo in enumerate(codigos[1:], start=2):
            etiqueta = Path(pasta_temporaria) / f"etiqueta_{indice}.docx"
            _renderizar_documento(titulo, codigo, rotulo, numero, etiqueta)
            documento.append(Document(etiqueta))

        try:
            documento.save(destino)
        except PermissionError:
            destino = pasta / f"{nome}_{datetime.now():%H%M%S}.docx"
            documento.save(destino)
    return destino


# ------------------------------------------------------- abrir / imprimir ---
def abrir(caminho: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(caminho)
    elif sys.platform == "darwin":
        subprocess.run(["open", caminho], check=False)
    else:
        subprocess.run(["xdg-open", caminho], check=False)


def imprimir(caminho: Path) -> None:
    """Envia para a impressora padrão."""
    if sys.platform.startswith("win"):
        os.startfile(caminho, "print")  # usa o Word instalado
    else:
        # macOS/Linux: converte para PDF com o LibreOffice e envia com 'lp'
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(caminho.parent), str(caminho)],
            check=True,
        )
        subprocess.run(["lp", str(caminho.with_suffix(".pdf"))], check=True)


# --------------------------------------------------------------- terminal ---
def perguntar(texto: str, padrao: str = "") -> str:
    sufixo = f" [{padrao}]" if padrao else ""
    while True:
        resposta = input(f"{texto}{sufixo}: ").strip()
        if resposta:
            return resposta
        if padrao:
            return padrao
        print("  (campo obrigatório)")


def main() -> None:
    print("=== Gerador de etiquetas ===\n")
    while True:
        modo = input("[M]anual ou [L]ote? (Enter = manual): ").strip().lower()
        titulo = perguntar("Título", PADRAO["titulo"])
        rotulo = perguntar("Rótulo do rodapé", PADRAO["rotulo"])
        numero = perguntar("Número do rodapé", PADRAO["numero"])

        try:
            if modo == "l":
                prateleira = perguntar("Prateleira (1, 2 ou 3)", "1")
                coluna_inicial = perguntar("Coluna inicial (A até R)", "A").upper()
                coluna_final = perguntar("Coluna final (A até R)", coluna_inicial).upper()
                codigos = gerar_codigos_lote(prateleira, coluna_inicial, coluna_final)
                arquivo = montar_lote(titulo, codigos, rotulo, numero)
                descricao = f"{len(codigos)} etiquetas"
            else:
                codigo = perguntar("Código de barras (ex.: 02-2-F2)")
                arquivo = montar_documento(titulo, codigo, rotulo, numero)
                descricao = "etiqueta"
        except Exception as erro:  # ex.: caractere inválido no Code 128
            print(f"\n  Não consegui gerar: {erro}\n")
            continue

        PADRAO.update(titulo=titulo, rotulo=rotulo, numero=numero)
        print(f"\n  {descricao.capitalize()} salvas em: {arquivo}")

        acao = input("  [A]brir  [I]mprimir  [Enter] só salvar: ").strip().lower()
        try:
            if acao == "a":
                abrir(arquivo)
            elif acao == "i":
                imprimir(arquivo)
        except Exception as erro:
            print(f"  Não consegui executar a ação: {erro}")

        if input("\nGerar outra? (s/N): ").strip().lower() != "s":
            break
        print()
if __name__ == "__main__":
    main()