# Split Thesis DOCX

A Python command-line utility for splitting a structured Microsoft Word (`.docx`) thesis or other long document into separate plain-text (`.txt`) files according to Word heading levels.

Current version: **0.2.0**

## Features

- Splits a `.docx` document by heading level.
- Reads numbering definitions from the Word document itself instead of simply reconstructing `1`, `1.1`, `1.2`, etc.
- Supports common Word numbering features, including:
  - starting values;
  - `startOverride`;
  - restart rules (`lvlRestart`);
  - decimal numbers;
  - Roman numerals;
  - alphabetical numbering.
- Detects headings through Word outline levels, including many custom heading styles, as well as standard `Heading 1`, `Heading 2`, etc.
- Can preview detected headings before creating any files.
- Can optionally preserve content before the first heading.
- Converts tables to plain text.
- Preserves footnote and endnote references and exports their contents to separate files.
- Marks figures and images in the text.
- Extracts text found in Word text boxes when it is stored in `w:txbxContent`.
- Avoids silently overwriting files when duplicate output filenames occur.

## Requirements

- Python 3
- [`python-docx`](https://python-docx.readthedocs.io/)

Install the dependency with:

```bash
pip install python-docx
```

or, if you cloned this repository:

```bash
pip install -r requirements.txt
```

## Usage

Basic usage:

```bash
python split_thesis.py "path/to/thesis.docx" "path/to/output_folder"
```

By default, the document is split at heading level 2.

### Split by chapter

```bash
python split_thesis.py "thesis.docx" "output" --level 1
```

### Split by section

```bash
python split_thesis.py "thesis.docx" "output" --level 2
```

### Split by subsection

```bash
python split_thesis.py "thesis.docx" "output" --level 3
```

### Preview headings without creating files

```bash
python split_thesis.py "thesis.docx" "output" --list
```

This is useful for checking how the program interprets the document structure before generating output.

### Include content before the first heading

```bash
python split_thesis.py "thesis.docx" "output" --include-preamble
```

The pre-heading content is written to:

```text
_preambulo.txt
```

### Show the program version

```bash
python split_thesis.py --version
```

## Example

Suppose a thesis contains:

```text
1. Introduction
1.1 Background
1.2 Research questions
2. Literature Review
2.1 Previous research
```

Running:

```bash
python split_thesis.py "thesis.docx" "output" --level 2
```

may produce files such as:

```text
1_Introduction.txt
1.1_Background.txt
1.2_Research_questions.txt
2_Literature_Review.txt
2.1_Previous_research.txt
```

The exact filenames depend on the numbering and heading titles stored in the Word document.

## How document content is represented

### Deeper headings

Headings below the selected split level remain inside the relevant output file and are marked in square brackets, for example:

```text
[2.1.1. Linguistic tests]
```

### Tables

Tables are converted to plain text:

```text
[TABELA]
Column 1 | Column 2 | Column 3
...
[FIM DA TABELA]
```

### Footnotes and endnotes

References remain in the main text as markers such as:

```text
[nota 12]
```

Their contents are exported separately to files such as:

```text
_notas_de_rodape.txt
_notas_finais.txt
```

### Figures and images

Images are represented as:

```text
[FIGURA]
```

The image files themselves are not extracted.

### Text boxes

Text found in supported Word text boxes is represented as:

```text
[CAIXA DE TEXTO]
Text from the text box
[FIM DA CAIXA DE TEXTO]
```

The graphical position and layout of the text box are not preserved.

## Limitations

Microsoft Word numbering is complex. Version 0.2 supports common hierarchical numbering, starting values, `startOverride`, `lvlRestart`, decimal numbering, letters, and Roman numerals. Rare Word numbering formats, dynamic fields, or unusual document templates may be simplified.

The program is designed primarily for structured `.docx` documents that use Word heading or outline levels. Documents whose visual headings are only manually formatted as bold or large text may not be recognised as headings.

Tables are converted to plain text, and their original formatting is not preserved. Images are marked but not extracted. Text-box content can be extracted in supported cases, but its visual placement is not retained.

## Repository files

```text
split-thesis-docx/
├── split_thesis.py
├── README.md
├── requirements.txt
└── .gitignore
```

`requirements.txt` should contain:

```text
python-docx
```

## Status

This is a research utility intended for processing structured academic Word documents. It has been developed and tested for common thesis structures, but users should use `--list` to inspect heading detection before processing important documents.
