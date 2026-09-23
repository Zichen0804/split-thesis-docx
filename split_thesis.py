#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
split_thesis.py
================
Separa uma tese em Word (.docx) em vários ficheiros .txt, um por secção,
usando os estilos de título (Heading 1, Heading 2, ...) e a numeração
automática real do Word para nomear cada ficheiro (1, 1.1, 1.1.1, ...).

Instalação (uma vez só):
    pip install python-docx

Utilização:
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida"
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida" --level 1
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida" --list

--level 1  -> separa por capítulo        (1, 2, 3, 4...)
--level 2  -> separa por secção X.Y      (1.1, 1.2, 2.1...)   [predefinição]
--level 3  -> separa por subsecção X.Y.Z (1.1.1, 3.1.1...)
--level 4  -> separa por X.Y.Z.W

--list     -> não escreve ficheiros, só mostra os títulos detetados e a
              numeração calculada, para confirmar antes de gerar tudo.

Notas sobre o que é preservado:
- O texto de cada secção mantém-se completo, incluindo títulos de nível
  mais profundo (aparecem marcados entre parênteses retos dentro do
  ficheiro, ex.: "[2.1.1. Linguistic tests]").
- Tabelas são convertidas para texto simples (colunas separadas por " | ").
- Notas de rodapé/finais NÃO ficam embutidas no meio do texto (o Word
  guarda o texto delas à parte); em vez disso, cada referência aparece
  como "[nota 12]" no sítio exato, e todo o conteúdo das notas é escrito
  num ficheiro à parte: "_notas_de_rodape.txt" (e "_notas_finais.txt" se
  existirem notas finais).
- Imagens/gráficos não podem ir para um .txt; aparecem como
  "[FIGURA]" no lugar onde estavam.
"""

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---------------------------------------------------------------------------
# Notas de rodapé / notas finais (lidas diretamente do XML interno do .docx)
# ---------------------------------------------------------------------------

def extract_notes(docx_path, part_name, note_tag):
    """Lê word/footnotes.xml ou word/endnotes.xml e devolve {id: texto}."""
    notes = {}
    with zipfile.ZipFile(docx_path) as z:
        if part_name not in z.namelist():
            return notes
        data = z.read(part_name)
    root = ET.fromstring(data)
    for note in root.findall(f"{W_NS}{note_tag}"):
        note_type = note.get(f"{W_NS}type")
        if note_type in ("separator", "continuationSeparator"):
            continue
        note_id = note.get(f"{W_NS}id")
        text = "".join(t.text or "" for t in note.iter(f"{W_NS}t"))
        notes[note_id] = text.strip()
    return notes


def write_notes_file(notes, output_dir, filename, label):
    if not notes:
        return
    path = os.path.join(output_dir, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"{label}\n{'=' * len(label)}\n\n")
        for note_id in sorted(notes, key=lambda x: int(x) if x.isdigit() else 0):
            f.write(f"[{note_id}] {notes[note_id]}\n\n")
    print(f"Escrito: {filename}  ({len(notes)} notas)")


# ---------------------------------------------------------------------------
# Leitura do corpo do documento (parágrafos e tabelas, na ordem em que
# aparecem, incluindo os que estão dentro de caixas de texto normais)
# ---------------------------------------------------------------------------

def iter_block_items(document):
    parent_elm = document.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def get_outline_level(paragraph):
    """Devolve 1..4 se o parágrafo usa um estilo 'Heading N', senão None."""
    style_name = (paragraph.style.name or "").strip().lower()
    m = re.match(r"heading (\d+)$", style_name)
    if m:
        return int(m.group(1))
    return None


def _numid_from_ppr(pPr):
    if pPr is None:
        return None
    numpr = pPr.find(qn("w:numPr"))
    if numpr is None:
        return None
    numid_el = numpr.find(qn("w:numId"))
    return numid_el.get(qn("w:val")) if numid_el is not None else None


def _style_numid(style):
    """Segue a cadeia de estilos (basedOn) até encontrar um numId definido
    no próprio estilo (é onde o Word costuma pôr a numeração automática
    dos títulos Heading 2/3/4, em vez de a repetir em cada parágrafo)."""
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        numid = _numid_from_ppr(style.element.find(qn("w:pPr")))
        if numid is not None:
            return numid
        style = style.base_style
    return None


def paragraph_is_numbered(paragraph):
    """True se o título tiver mesmo um número visível no Word.
    numId='0' é a forma como o Word desliga a numeração herdada para um
    título específico (ex.: 'Introduction', 'Bibliography'), por isso
    conta como NÃO numerado mesmo tendo numPr."""
    numid = _numid_from_ppr(paragraph._p.pPr)
    if numid is None:
        numid = _style_numid(paragraph.style)
    return numid is not None and numid != "0"


def paragraph_text_with_markers(paragraph):
    """Como paragraph.text, mas insere [nota N] onde há uma referência
    de nota de rodapé/final, e [FIGURA] onde há uma imagem."""
    parts = []
    for run in paragraph.runs:
        for ref in run._element.findall(qn("w:footnoteReference")):
            parts.append(f"[nota {ref.get(qn('w:id'))}]")
        for ref in run._element.findall(qn("w:endnoteReference")):
            parts.append(f"[nota final {ref.get(qn('w:id'))}]")
        if run._element.findall(qn("w:drawing")) or run._element.findall(qn("w:pict")):
            parts.append("[FIGURA]")
        parts.append(run.text)
    return "".join(parts)


def table_to_text(table):
    lines = []
    for row in table.rows:
        cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
        lines.append(" | ".join(cells))
    return "\n".join(lines)


def safe_filename(text, max_len=60):
    text = re.sub(r'[\\/:*?"<>|]', "", text).strip()
    text = re.sub(r"\s+", "_", text)
    return text[:max_len] if text else "sem_titulo"


# ---------------------------------------------------------------------------
# Lógica principal de separação
# ---------------------------------------------------------------------------

def split_thesis(docx_path, output_dir, split_level, list_only=False):
    document = Document(docx_path)

    counters = [0, 0, 0, 0, 0]  # posições 1..4 usadas
    current_number = None
    current_title = None
    current_lines = []
    seen_first_heading = False
    file_index = 0
    detected = []  # para o modo --list

    def flush():
        nonlocal current_lines, current_number, current_title, file_index
        if current_title is None:
            return
        text = "\n".join(current_lines).strip()
        file_index += 1
        label = f"{current_number}. {current_title}" if current_number else current_title
        if list_only:
            detected.append((label, len(text)))
            return
        if not text:
            return
        prefix = current_number if current_number else f"{file_index:02d}"
        fname = f"{prefix}_{safe_filename(current_title)}.txt"
        path = os.path.join(output_dir, fname)
        with open(path, "w", encoding="utf-8") as f:
            f.write(label + "\n")
            f.write("=" * len(label) + "\n\n")
            f.write(text)
        print(f"Escrito: {fname}  ({len(text)} caracteres)")

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            level = get_outline_level(block)

            if level is not None and 1 <= level <= 4:
                if paragraph_is_numbered(block):
                    counters[level] += 1
                    for l in range(level + 1, 5):
                        counters[l] = 0
                    number = ".".join(str(counters[i]) for i in range(1, level + 1))
                else:
                    number = None

                if level <= split_level:
                    flush()
                    current_lines = []
                    current_title = block.text.strip()
                    current_number = number
                    seen_first_heading = True
                else:
                    prefix = f"{number}. " if number else ""
                    current_lines.append("")
                    current_lines.append(f"[{prefix}{block.text.strip()}]")
                continue

            if seen_first_heading:
                text = paragraph_text_with_markers(block)
                if text.strip():
                    current_lines.append(text)
        else:
            if seen_first_heading:
                current_lines.append("")
                current_lines.append("[TABELA]")
                current_lines.append(table_to_text(block))
                current_lines.append("[FIM DA TABELA]")

    flush()

    if list_only:
        print(f"\n{'Nº/Título':<70} caracteres")
        print("-" * 85)
        for label, length in detected:
            print(f"{label[:68]:<70} {length}")
        print(f"\nTotal de secções detetadas ao nível {split_level}: {len(detected)}")
        return

    footnotes = extract_notes(docx_path, "word/footnotes.xml", "footnote")
    endnotes = extract_notes(docx_path, "word/endnotes.xml", "endnote")
    write_notes_file(footnotes, output_dir, "_notas_de_rodape.txt", "Notas de rodapé")
    write_notes_file(endnotes, output_dir, "_notas_finais.txt", "Notas finais")

    print(f"\nConcluído: {file_index} ficheiro(s) processado(s) em: {output_dir}")


def main():
    parser = argparse.ArgumentParser(
        description="Separa uma tese .docx em ficheiros .txt por secção, preservando a numeração real do Word."
    )
    parser.add_argument("docx_path", help="Caminho para o ficheiro .docx da tese")
    parser.add_argument("output_dir", help="Pasta onde gravar os ficheiros separados")
    parser.add_argument(
        "--level", type=int, default=2, choices=[1, 2, 3, 4],
        help="Nível de título pelo qual separar: 1=capítulo, 2=secção X.Y (predefinição), 3=X.Y.Z, 4=X.Y.Z.W",
    )
    parser.add_argument(
        "--list", action="store_true",
        help="Só mostra os títulos e a numeração detetados, sem escrever ficheiros",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.docx_path):
        print(f"Erro: não encontrei o ficheiro '{args.docx_path}'")
        sys.exit(1)

    if not args.list:
        os.makedirs(args.output_dir, exist_ok=True)

    split_thesis(args.docx_path, args.output_dir, args.level, list_only=args.list)


if __name__ == "__main__":
    main()
