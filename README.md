# PDF Manager

Ferramenta desktop (Tkinter) para manipulação de arquivos PDF, sem depender de serviços online.

## Funcionalidades

- **Juntar PDFs** — combina vários arquivos, com opção de intercalar páginas
- **Cortar PDF** — extrai um intervalo de páginas (ex: `1-5, 7, 9-12`)
- **Dividir PDF** — separa em múltiplos arquivos por intervalos definidos manualmente ou importados do sumário/bookmarks do próprio PDF
- **Reorganizar páginas** — reordena, inverte ou separa páginas pares/ímpares
- **Converter PDF** — exporta páginas como imagens, com resolução (DPI) configurável
- **Inserir páginas** — insere páginas de outro PDF ou páginas em branco em qualquer posição

## Requisitos

- Python 3.12+
- Na primeira execução, o próprio script verifica e instala as dependências que faltarem (`PyPDF2`, `Pillow`, `PyMuPDF`, `python-docx`)

## Instalação

```bash
git clone https://github.com/eulerchavesneto/pdf_manager.git
cd pdf_manager
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

```bash
python pdf_manager.py
```
