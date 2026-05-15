# Shiori2 also adds bookmarks to PDF files that do not already contain bookmarks.
# Shiori2 covers more PDF files than Shiori.


## PDF Format

Shiori assumes that the target PDF file follows the structure below and searches for these heading patterns to create hierarchical bookmarks.

```text
H1 1. [Title text]
H2 1.1 [Title text]
H3 1.1.1 [Title text]
H4 1.1.1.1 [Title text]
H5 (1) [Title text]
H6 (a) [Title text]
```

- H1 headings must start at the first column of a line.
- H2 and lower-level headings may start at the first or any subsequent column.

## Installation

```bash
# Clone this repository.
git clone <repository_url>
cd <repository_name>

python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
% python3 shiori.py
```

Then follow the on-screen instructions.

## Environment

- Python >= 3.9
- FreeSimpleGUI >= 5.0.0
- pypdf >= 3.0.0
- pdfminer.six >= 20221105

## Remarks

This package was created collaboratively by Claude Code and a human with very limited programming experience.
