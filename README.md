# Gerador de Etiquetas de Localização

Aplicação desktop que preenche um modelo Word (`.docx`), gera o **código de barras Code 128** e monta as etiquetas de localização do almoxarifado, uma por página, prontas para **salvar, abrir ou imprimir**, individualmente ou **em lote**.

Antes, cada etiqueta exigia gerar o código de barras em um site, colar a imagem no Word e trocar os textos à mão. Agora basta escolher o intervalo e clicar em um botão.


## Funcionalidades

- **Geração em lote**: escolha a prateleira e o intervalo de colunas (A a R) e todas as etiquetas saem em um único arquivo, uma por página.
- **Modo manual**: digite um código (ex.: `02-2-F2`) e gere só aquela etiqueta.
- **Prévia** dos códigos que serão gerados antes de criar o documento.
- **Código de barras local**: o Code 128 é desenhado no próprio computador. Não usa internet nem token de API.
- **Layout fiel ao modelo original**: título, código de barras e rodapé nas mesmas posições do documento Word de referência.
- **Abrir e imprimir** direto da interface, com envio para a impressora padrão.
- **Geração em segundo plano**, sem travar a janela em lotes grandes.
- **Pasta de saída configurável**.
- **Interface** em tons de vermelho, cinza-escuro e azul da ABB.

## Padrão dos códigos

```
PRATELEIRA - ANDAR - COLUNA + ANDAR
   01      -   1   -   A    +   1     →  01-1-A1
```

Exemplos: `01-1-A1`, `01-1-B1`, `01-2-A2`, `03-5-R5`.

O padrão atual cobre **3 prateleiras**, **5 andares** e colunas de **A a R**.
Um lote completo tem 3 × 5 × 18 = **270 etiquetas**, geradas em cerca de 9 segundos.

## Como funciona

```
Dados informados ──► Código de barras (Pillow) ──► Modelo .docx (docxtpl) ──► .docx final
                                                                                    │
                                                          Lote: junta as páginas (docxcompose)
                                                                                    │
                                                                        Salvar / Abrir / Imprimir
```

## Requisitos

- **Python 3.10+**
- **Windows** com **Microsoft Word** instalado (usado para abrir e imprimir)
- Fonte **Aptos** (vem no Microsoft 365 recente). O alinhamento do rodapé do modelo depende dela. Em versões do Word sem a Aptos, o texto pode sair deslocado.
- Uma **impressora padrão** configurada no Windows

Dependências Python:

| Biblioteca | Uso |
|---|---|
| `docxtpl` | Preencher o modelo Word |
| `python-barcode` | Padrão de barras Code 128 |
| `pillow` | Desenhar a imagem do código de barras |
| `docxcompose` | Juntar as etiquetas de um lote em um só arquivo |

O `tkinter` da interface já vem com o Python no Windows.

## Instalação

```bash
git clone https://github.com/‹seu-usuario›/‹nome-do-repositorio›.git
cd ‹nome-do-repositorio›

python -m venv .venv
.venv\Scripts\activate

pip install docxtpl python-barcode pillow docxcompose
```

## Uso

### Interface gráfica

```bash
python ‹pasta-da-interface›/‹arquivo-da-interface›.py
```

1. Escolha **Lote** ou **Manual**.
2. No lote, selecione a **prateleira**, a **coluna inicial** e a **coluna final**. A prévia mostra os códigos que serão criados.
3. Confira **título**, **rótulo** e **número** do rodapé (valores padrão: `AUEN`, `CENTRO`, `1062`).
4. Se quiser, escolha a **pasta de saída**.
5. Clique em **Gerar documento** ou em **Gerar e imprimir**.
6. Use **Abrir último** e **Imprimir último** para repetir a ação sobre o último arquivo gerado.

Os arquivos são salvos na pasta `saida/` (ou na pasta escolhida).

### Como biblioteca

```python
from pathlib import Path
from autoamte import gerar_codigos_lote, montar_documento, montar_lote

pasta = Path("saida")

# Uma etiqueta
montar_documento("AUEN", "02-2-F2", "CENTRO", "1062", pasta)

# Lote: prateleira 1, colunas A até C
codigos = gerar_codigos_lote("1", "A", "C")
montar_lote("AUEN", codigos, "CENTRO", "1062", pasta)
```

## Estrutura do projeto

```
.
├── autoamte.py                 # Núcleo: códigos, código de barras, modelo, lote, abrir e imprimir
├── ‹pasta-da-interface›/
│   └── ‹arquivo-da-interface›.py   # Interface desktop (tkinter)
├── modelo_etiqueta.docx        # Modelo Word com os campos a preencher
├── docs/
│   └── screenshot.png
└── README.md
```

## Modelo Word

O `modelo_etiqueta.docx` é uma página **A4 paisagem** com campos que o programa substitui:

| Campo | Conteúdo |
|---|---|
| `{{ titulo }}` | Título grande no topo (centralizado, negrito, 100 pt) |
| `{{ codigo_barras }}` | Imagem do código de barras (21,67 × 9,57 cm) |
| `{{ rotulo }}` | Texto do rodapé, à esquerda (negrito, 48 pt) |
| `{{ numero }}` | Número do rodapé, à direita (negrito, 52 pt) |

Para mudar o layout, edite o modelo no Word e **mantenha os campos entre chaves duplas** exatamente como estão. Cada campo deve ficar em um único trecho de texto, sem mudar de formatação no meio.

## Gerando o executável (.exe)

Com o [PyInstaller](https://pyinstaller.org), o programa vira um único `.exe`, com o modelo embutido, que roda em computadores **sem Python instalado**:

```bash
pip install pyinstaller

pyinstaller --onefile --windowed --name GeradorEtiquetas ^
  --icon etiqueta.ico ^
  --add-data "modelo_etiqueta.docx;." ^
  --paths . ^
  --collect-all docxtpl --collect-all docxcompose --collect-all docx --collect-all barcode ^
  ‹pasta-da-interface›/‹arquivo-da-interface›.py
```

O resultado fica em `dist/GeradorEtiquetas.exe`. Alguns detalhes:

- `--windowed` esconde a janela preta do terminal, o que faz sentido porque o programa tem interface gráfica.
- No Windows o separador do `--add-data` é `;`.
- Se existir um `modelo_etiqueta.docx` ao lado do `.exe`, ele tem prioridade sobre o embutido. Assim é possível ajustar o layout sem gerar o executável de novo.
- Antivírus e SmartScreen às vezes bloqueiam executáveis criados com PyInstaller. Se acontecer, peça à TI para liberar o arquivo.

## Limitações conhecidas

- O **Code 128 aceita apenas caracteres ASCII**. Acentos e `ç` no código geram erro.
- A impressão usa a **impressora padrão** do Windows, via Word.
- O alinhamento do rodapé depende da fonte **Aptos** estar instalada.
- Lotes muito grandes geram arquivos com centenas de páginas, então confira o intervalo antes de imprimir.

## Segurança e privacidade

O código de barras é gerado localmente. Não há chamadas a serviços externos nem tokens de API no projeto ou no modelo.

## Licença


## Autor

- Murilo Pires Andrade Cruz
- Estudante de analise e desenvolvimento de sistemas 

‹Seu nome› — [GitHub](https://github.com/Muriloves27)
