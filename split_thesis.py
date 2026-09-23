#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
split_thesis.py — v0.2.0
=========================
Separa uma tese ou outro documento Word (.docx) estruturado em vários
ficheiros .txt, um por secção.

Principais melhorias da v0.2:
- Lê as definições de numeração do próprio ficheiro .docx (numbering.xml),
  incluindo valores iniciais, formatos comuns, startOverride e regras de
  reinício (lvlRestart), em vez de reconstruir simplesmente 1, 1.1, 1.2...
- Deteta níveis de título através de outlineLvl, inclusive em estilos
  personalizados, além dos estilos normais Heading 1, Heading 2, etc.
- Pode guardar o conteúdo anterior ao primeiro título com --include-preamble.
- Extrai texto encontrado em caixas de texto ancoradas em parágrafos e marca-o
  explicitamente; a posição/layout gráfico da caixa de texto não é preservado.

Instalação (uma vez só):
    pip install python-docx

Utilização:
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida"
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida" --level 1
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida" --list
    python split_thesis.py "caminho\\para\\tese.docx" "caminho\\para\\pasta_saida" --include-preamble

--level 1  -> separa por títulos de nível 1
--level 2  -> separa por títulos até ao nível 2 [predefinição]
--level 3  -> separa por títulos até ao nível 3
--level 4  -> separa por títulos até ao nível 4

--list              -> não escreve ficheiros; mostra os títulos detetados.
--include-preamble  -> grava o conteúdo anterior ao primeiro título em
                       "_preambulo.txt".

Notas e limitações:
- O texto das secções mantém títulos de nível mais profundo entre parênteses
  retos, por exemplo: "[2.1.1. Linguistic tests]".
- Tabelas são convertidas para texto simples, com colunas separadas por " | ".
- Referências a notas de rodapé/finais aparecem no texto como "[nota N]" e o
  conteúdo das notas é exportado para ficheiros separados.
- Imagens/gráficos aparecem como "[FIGURA]"; os ficheiros de imagem não são
  extraídos.
- Texto de caixas de texto ancoradas em parágrafos é extraído quando está
  presente em w:txbxContent, mas o layout/posição da caixa não é preservado.
- A numeração Word é complexa. Esta versão cobre numeração hierárquica normal,
  start/startOverride, lvlRestart e formatos decimal, letras e romanos. Formatos
  Word raros ou campos dinâmicos podem ser simplificados.
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

__version__ = "0.2.0"

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---------------------------------------------------------------------------
# Utilitários XML
# ---------------------------------------------------------------------------

def _attr(element, local_name, default=None):
    if element is None:
        return default
    return element.get(f"{W_NS}{local_name}", default)


def _child_val(parent, local_name, default=None):
    if parent is None:
        return default
    child = parent.find(f"{W_NS}{local_name}")
    return _attr(child, "val", default)


def _int_or_none(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Notas de rodapé / notas finais
# ---------------------------------------------------------------------------

def extract_notes(docx_path, part_name, note_tag):
    """Lê footnotes.xml/endnotes.xml e devolve {id: texto}."""
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
# Corpo do documento
# ---------------------------------------------------------------------------

def iter_block_items(document):
    """Percorre parágrafos e tabelas de nível superior pela ordem do corpo."""
    parent_elm = document.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def _outline_level_from_ppr(ppr):
    if ppr is None:
        return None
    outline = ppr.find(qn("w:outlineLvl"))
    value = _int_or_none(_attr(outline, "val"))
    if value is None:
        return None
    return value + 1  # XML é 0-based; interface é 1-based.


def _style_outline_level(style):
    """Procura outlineLvl no estilo e na cadeia basedOn."""
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        level = _outline_level_from_ppr(style.element.find(qn("w:pPr")))
        if level is not None:
            return level
        style = style.base_style
    return None


def get_outline_level(paragraph):
    """Devolve o nível de título (1..9) ou None.

    Prioridade:
      1) outlineLvl diretamente no parágrafo;
      2) outlineLvl no estilo (incluindo estilos personalizados/baseados);
      3) fallback para nomes de estilo Heading N.
    """
    level = _outline_level_from_ppr(paragraph._p.pPr)
    if level is not None:
        return level

    level = _style_outline_level(paragraph.style)
    if level is not None:
        return level

    style_name = (paragraph.style.name or "").strip().lower()
    match = re.fullmatch(r"heading\s+(\d+)", style_name)
    if match:
        return int(match.group(1))
    return None


def _textbox_texts(element):
    """Extrai texto de w:txbxContent dentro de um desenho/shape."""
    texts = []
    for txbx in element.iter(qn("w:txbxContent")):
        paragraphs = []
        for p in txbx.iter(qn("w:p")):
            pieces = []
            for node in p.iter():
                if node.tag == qn("w:t"):
                    pieces.append(node.text or "")
                elif node.tag == qn("w:tab"):
                    pieces.append("\t")
                elif node.tag in (qn("w:br"), qn("w:cr")):
                    pieces.append("\n")
                elif node.tag == qn("w:footnoteReference"):
                    pieces.append(f"[nota {_attr(node, 'id')}]" )
                elif node.tag == qn("w:endnoteReference"):
                    pieces.append(f"[nota final {_attr(node, 'id')}]" )
            text = "".join(pieces).strip()
            if text:
                paragraphs.append(text)
        if paragraphs:
            texts.append("\n".join(paragraphs))
    return texts


def _render_inline_xml(node):
    """Renderiza texto inline preservando marcadores relevantes."""
    tag = node.tag

    if tag == qn("w:t"):
        return node.text or ""
    if tag == qn("w:tab"):
        return "\t"
    if tag in (qn("w:br"), qn("w:cr")):
        return "\n"
    if tag == qn("w:footnoteReference"):
        return f"[nota {_attr(node, 'id')}]"
    if tag == qn("w:endnoteReference"):
        return f"[nota final {_attr(node, 'id')}]"

    if tag in (qn("w:drawing"), qn("w:pict")):
        boxes = _textbox_texts(node)
        if boxes:
            rendered = []
            for text in boxes:
                rendered.append(f"\n[CAIXA DE TEXTO]\n{text}\n[FIM DA CAIXA DE TEXTO]\n")
            return "".join(rendered)
        return "[FIGURA]"

    # Evita que propriedades XML sejam tratadas como conteúdo.
    if tag in (qn("w:pPr"), qn("w:rPr")):
        return ""

    return "".join(_render_inline_xml(child) for child in list(node))


def paragraph_text_with_markers(paragraph):
    """Texto do parágrafo com notas, figuras e caixas de texto marcadas."""
    return "".join(_render_inline_xml(child) for child in list(paragraph._p))


def table_to_text(table):
    lines = []
    for row in table.rows:
        cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
        lines.append(" | ".join(cells))
    return "\n".join(lines)


def append_table(lines, table):
    lines.append("")
    lines.append("[TABELA]")
    lines.append(table_to_text(table))
    lines.append("[FIM DA TABELA]")


def safe_filename(text, max_len=80):
    text = re.sub(r'[\\/:*?"<>|]', "", text).strip()
    text = re.sub(r"\s+", "_", text)
    text = text.rstrip(". ")
    return text[:max_len] if text else "sem_titulo"


def unique_path(directory, filename):
    """Evita sobrescrever ficheiros com o mesmo nome."""
    path = os.path.join(directory, filename)
    if not os.path.exists(path):
        return path

    stem, ext = os.path.splitext(filename)
    i = 2
    while True:
        candidate = os.path.join(directory, f"{stem}_{i}{ext}")
        if not os.path.exists(candidate):
            return candidate
        i += 1


# ---------------------------------------------------------------------------
# Numeração Word (numbering.xml)
# ---------------------------------------------------------------------------

def _roman(number):
    if number <= 0:
        return str(number)
    values = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    result = []
    remaining = number
    for value, symbol in values:
        while remaining >= value:
            result.append(symbol)
            remaining -= value
    return "".join(result)


def _letters(number):
    if number <= 0:
        return str(number)
    result = []
    n = number
    while n:
        n -= 1
        result.append(chr(ord("A") + (n % 26)))
        n //= 26
    return "".join(reversed(result))


def _format_number(number, fmt):
    fmt = fmt or "decimal"
    if fmt == "upperRoman":
        return _roman(number)
    if fmt == "lowerRoman":
        return _roman(number).lower()
    if fmt == "upperLetter":
        return _letters(number)
    if fmt == "lowerLetter":
        return _letters(number).lower()
    if fmt == "decimalZero":
        return f"{number:02d}"
    # Para formatos textuais/raros que exigiriam regras linguísticas específicas,
    # usa o valor decimal em vez de inventar uma representação incorreta.
    return str(number)


def _parse_level(level_el):
    ilvl = _int_or_none(_attr(level_el, "ilvl"))
    if ilvl is None:
        return None, None

    start = _int_or_none(_child_val(level_el, "start"))
    if start is None:
        start = 0

    restart_el = level_el.find(f"{W_NS}lvlRestart")
    restart = _int_or_none(_attr(restart_el, "val")) if restart_el is not None else None

    level = {
        "start": start,
        "numFmt": _child_val(level_el, "numFmt", "decimal"),
        "lvlText": _child_val(level_el, "lvlText", f"%{ilvl + 1}"),
        "restart": restart,
        "isLgl": level_el.find(f"{W_NS}isLgl") is not None,
        "pStyle": _child_val(level_el, "pStyle"),
    }
    return ilvl, level


class WordNumberingResolver:
    """Resolve a numeração visível de parágrafos Word em casos normais.

    Mantém estado por numId, processa os parágrafos na ordem do documento e
    respeita start/startOverride e lvlRestart. Não pretende implementar todos
    os formatos históricos/esotéricos do Word.
    """

    def __init__(self, docx_path):
        self.abstract_levels = {}
        self.num_to_abstract = {}
        self.overrides = {}
        self.state = {}
        self.style_candidates = {}
        self._load(docx_path)

    def _load(self, docx_path):
        with zipfile.ZipFile(docx_path) as z:
            if "word/numbering.xml" not in z.namelist():
                return
            root = ET.fromstring(z.read("word/numbering.xml"))

        for abstract in root.findall(f"{W_NS}abstractNum"):
            abstract_id = _attr(abstract, "abstractNumId")
            levels = {}
            for level_el in abstract.findall(f"{W_NS}lvl"):
                ilvl, level = _parse_level(level_el)
                if ilvl is not None:
                    levels[ilvl] = level
            self.abstract_levels[abstract_id] = levels

        for num in root.findall(f"{W_NS}num"):
            num_id = _attr(num, "numId")
            abstract_id = _child_val(num, "abstractNumId")
            if num_id is None or abstract_id is None:
                continue
            self.num_to_abstract[num_id] = abstract_id

            num_overrides = {}
            for override in num.findall(f"{W_NS}lvlOverride"):
                ilvl = _int_or_none(_attr(override, "ilvl"))
                if ilvl is None:
                    continue
                entry = {}
                start_override = _int_or_none(_child_val(override, "startOverride"))
                if start_override is not None:
                    entry["startOverride"] = start_override

                level_el = override.find(f"{W_NS}lvl")
                if level_el is not None:
                    _, parsed = _parse_level(level_el)
                    if parsed is not None:
                        # Word ignora lvlRestart dentro de lvlOverride; não o
                        # usamos para alterar a regra do nível abstrato.
                        parsed.pop("restart", None)
                        entry["level"] = parsed
                if entry:
                    num_overrides[ilvl] = entry
            self.overrides[num_id] = num_overrides

        # Fallback conservador para estilos ligados a um nível via w:pStyle.
        # Só é usado quando um estilo aponta inequivocamente para um único
        # (numId, ilvl).
        candidates = {}
        for num_id, abstract_id in self.num_to_abstract.items():
            for ilvl, level in self.abstract_levels.get(abstract_id, {}).items():
                style_id = level.get("pStyle")
                if style_id:
                    candidates.setdefault(style_id, set()).add((num_id, ilvl))
        self.style_candidates = candidates

    def _effective_level(self, num_id, ilvl):
        abstract_id = self.num_to_abstract.get(num_id)
        base = dict(self.abstract_levels.get(abstract_id, {}).get(ilvl, {}))
        override = self.overrides.get(num_id, {}).get(ilvl, {})
        if "level" in override:
            base.update(override["level"])
        if "startOverride" in override:
            base["start"] = override["startOverride"]
        return base

    @staticmethod
    def _numpr_values(ppr):
        if ppr is None:
            return None, None, False
        numpr = ppr.find(qn("w:numPr"))
        if numpr is None:
            return None, None, False
        num_id = _child_val(numpr, "numId")
        ilvl = _int_or_none(_child_val(numpr, "ilvl"))
        return num_id, ilvl, True

    def _paragraph_num_info(self, paragraph, outline_level=None):
        # Propriedades diretas têm prioridade; valores em falta podem ser
        # herdados da cadeia de estilos.
        num_id, ilvl, direct_has_numpr = self._numpr_values(paragraph._p.pPr)

        if num_id == "0":
            return None

        style = paragraph.style
        seen = set()
        while style is not None and style.style_id not in seen and (num_id is None or ilvl is None):
            seen.add(style.style_id)
            s_num_id, s_ilvl, has_numpr = self._numpr_values(style.element.find(qn("w:pPr")))
            if has_numpr:
                if num_id is None:
                    num_id = s_num_id
                if ilvl is None:
                    ilvl = s_ilvl
                if num_id == "0":
                    return None
            style = style.base_style

        if num_id is None:
            # Alguns modelos ligam o estilo ao nível através de pStyle.
            style_id = paragraph.style.style_id if paragraph.style is not None else None
            choices = self.style_candidates.get(style_id, set())
            if len(choices) == 1:
                num_id, linked_ilvl = next(iter(choices))
                if ilvl is None:
                    ilvl = linked_ilvl

        if num_id is None or num_id not in self.num_to_abstract:
            return None

        if ilvl is None and outline_level is not None:
            ilvl = outline_level - 1
        if ilvl is None:
            ilvl = 0

        return num_id, ilvl

    def _start(self, num_id, ilvl):
        return self._effective_level(num_id, ilvl).get("start", 0)

    def _reset_deeper_levels(self, num_id, used_ilvl):
        state = self.state.setdefault(num_id, {})
        for deeper_ilvl in list(state):
            if deeper_ilvl <= used_ilvl:
                continue

            level = self._effective_level(num_id, deeper_ilvl)
            restart = level.get("restart")

            if restart == 0:
                continue  # nunca reinicia

            if restart is None:
                # O Word reinicia quando o nível anterior ou qualquer nível
                # ainda mais alto é usado.
                should_reset = used_ilvl <= deeper_ilvl - 1
            else:
                # lvlRestart é 1-based: valor 2 => níveis 1 e 2 (ilvl 0/1)
                # fazem reiniciar este nível.
                should_reset = used_ilvl <= restart - 1

            if should_reset:
                state[deeper_ilvl] = None

    def _display_text(self, num_id, ilvl):
        current_level = self._effective_level(num_id, ilvl)
        fmt = current_level.get("numFmt", "decimal")
        if fmt in ("bullet", "none"):
            return None

        template = current_level.get("lvlText") or f"%{ilvl + 1}"
        legal = current_level.get("isLgl", False)
        state = self.state.setdefault(num_id, {})

        def replace(match):
            referenced_ilvl = int(match.group(1)) - 1
            value = state.get(referenced_ilvl)
            if value is None:
                value = self._start(num_id, referenced_ilvl)
            referenced_level = self._effective_level(num_id, referenced_ilvl)
            referenced_fmt = "decimal" if legal else referenced_level.get("numFmt", "decimal")
            return _format_number(value, referenced_fmt)

        return re.sub(r"%([1-9])", replace, template).strip()

    def number_for(self, paragraph, outline_level=None):
        """Atualiza o estado de numeração e devolve o símbolo visível."""
        info = self._paragraph_num_info(paragraph, outline_level=outline_level)
        if info is None:
            return None

        num_id, ilvl = info
        state = self.state.setdefault(num_id, {})
        current = state.get(ilvl)
        if current is None:
            state[ilvl] = self._start(num_id, ilvl)
        else:
            state[ilvl] = current + 1

        display = self._display_text(num_id, ilvl)
        self._reset_deeper_levels(num_id, ilvl)
        return display


# ---------------------------------------------------------------------------
# Escrita / separação
# ---------------------------------------------------------------------------

def heading_label(number, title):
    if not number:
        return title
    if re.search(r"[.\):;\-–—]$", number):
        return f"{number} {title}"
    return f"{number}. {title}"


def write_text_file(path, label, text):
    with open(path, "w", encoding="utf-8") as f:
        if label:
            f.write(label + "\n")
            f.write("=" * len(label) + "\n\n")
        f.write(text)


def split_thesis(docx_path, output_dir, split_level, list_only=False, include_preamble=False):
    document = Document(docx_path)
    numbering = WordNumberingResolver(docx_path)

    current_number = None
    current_title = None
    current_lines = []
    preamble_lines = []
    seen_first_heading = False
    file_index = 0
    written_count = 0
    detected = []

    def flush():
        nonlocal current_lines, current_number, current_title, file_index, written_count
        if current_title is None:
            return

        text = "\n".join(current_lines).strip()
        file_index += 1
        label = heading_label(current_number, current_title)

        if list_only:
            detected.append((label, len(text)))
            return
        if not text:
            return

        if current_number:
            prefix = safe_filename(current_number, max_len=35)
        else:
            prefix = f"{file_index:02d}"
        fname = f"{prefix}_{safe_filename(current_title)}.txt"
        path = unique_path(output_dir, fname)
        write_text_file(path, label, text)
        written_count += 1
        print(f"Escrito: {os.path.basename(path)}  ({len(text)} caracteres)")

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            level = get_outline_level(block)
            # Importante: processa a numeração de TODOS os parágrafos numerados,
            # não apenas dos títulos, para manter o estado Word coerente.
            number = numbering.number_for(block, outline_level=level)

            if level is not None and 1 <= level <= 9:
                title = block.text.strip() or paragraph_text_with_markers(block).strip()

                if level <= split_level:
                    flush()
                    current_lines = []
                    current_title = title
                    current_number = number
                    seen_first_heading = True
                else:
                    # Um título mais profundo fica dentro da secção atual.
                    if current_title is not None:
                        label = heading_label(number, title)
                        current_lines.append("")
                        current_lines.append(f"[{label}]")
                    elif include_preamble:
                        label = heading_label(number, title)
                        preamble_lines.append(f"[{label}]")
                    seen_first_heading = True
                continue

            text = paragraph_text_with_markers(block)
            if seen_first_heading and current_title is not None:
                if text.strip():
                    current_lines.append(text)
            elif not seen_first_heading and include_preamble:
                if text.strip():
                    preamble_lines.append(text)

        else:  # Table
            if seen_first_heading and current_title is not None:
                append_table(current_lines, block)
            elif not seen_first_heading and include_preamble:
                append_table(preamble_lines, block)

    flush()

    preamble_text = "\n".join(preamble_lines).strip()

    if list_only:
        if include_preamble and preamble_text:
            print(f"\n{'PREÂMBULO':<70} {len(preamble_text)} caracteres")
        print(f"\n{'Nº/Título':<70} caracteres")
        print("-" * 85)
        for label, length in detected:
            print(f"{label[:68]:<70} {length}")
        print(f"\nTotal de secções detetadas ao nível {split_level}: {len(detected)}")
        return

    if include_preamble and preamble_text:
        preamble_path = unique_path(output_dir, "_preambulo.txt")
        write_text_file(preamble_path, "Preâmbulo", preamble_text)
        print(f"Escrito: {os.path.basename(preamble_path)}  ({len(preamble_text)} caracteres)")

    footnotes = extract_notes(docx_path, "word/footnotes.xml", "footnote")
    endnotes = extract_notes(docx_path, "word/endnotes.xml", "endnote")
    write_notes_file(footnotes, output_dir, "_notas_de_rodape.txt", "Notas de rodapé")
    write_notes_file(endnotes, output_dir, "_notas_finais.txt", "Notas finais")

    print(f"\nConcluído: {written_count} ficheiro(s) de secção escritos em: {output_dir}")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Separa um .docx estruturado em ficheiros .txt por nível de título, "
            "lendo a numeração definida no próprio Word."
        )
    )
    parser.add_argument("docx_path", help="Caminho para o ficheiro .docx")
    parser.add_argument("output_dir", help="Pasta onde gravar os ficheiros separados")
    parser.add_argument(
        "--level",
        type=int,
        default=2,
        choices=[1, 2, 3, 4],
        help="Nível de título pelo qual separar: 1, 2 (predefinição), 3 ou 4",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Só mostra os títulos e a numeração detetados, sem escrever ficheiros",
    )
    parser.add_argument(
        "--include-preamble",
        action="store_true",
        help="Inclui o conteúdo anterior ao primeiro título em _preambulo.txt",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.docx_path):
        print(f"Erro: não encontrei o ficheiro '{args.docx_path}'")
        sys.exit(1)

    if not zipfile.is_zipfile(args.docx_path):
        print(f"Erro: '{args.docx_path}' não parece ser um ficheiro .docx válido")
        sys.exit(1)

    if not args.list:
        os.makedirs(args.output_dir, exist_ok=True)

    try:
        split_thesis(
            args.docx_path,
            args.output_dir,
            args.level,
            list_only=args.list,
            include_preamble=args.include_preamble,
        )
    except Exception as exc:
        print(f"Erro ao processar o documento: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
