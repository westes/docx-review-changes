# docx-review-changes

Extract all tracked changes (insertions, deletions, substitutions) and comments
from a Microsoft Word `.docx` file and print them as readable, paragraph-numbered
text — useful for reviewing edits without opening Word.

## Install

Requires Python 3.9+ and [pipx](https://pipx.pypa.io/).

```sh
pipx install git+https://gitea.reground.org/will/docx-review-changes.git
```

To upgrade to the latest version later:

```sh
pipx install --force git+https://gitea.reground.org/will/docx-review-changes.git
```

No external dependencies — the tool only uses the Python standard library.

## Usage

```sh
docx-review-changes yourfile.docx
```

Each tracked change or comment is printed as a numbered entry with the
paragraph number, author, and surrounding context. Example output:

```
────────────────────────────────────────────────────────────
[1]  CHANGE  —  paragraph 3  —  Jane Doe

  Delete  : "quick"
  Insert  : "fast"
  Context : The …[^^^] brown fox jumps over the lazy dog…

────────────────────────────────────────────────────────────
[2]  COMMENT  —  paragraph 5  —  John Smith

  Highlighted : "lazy dog"
  Comment     : "Can we rename this?"
  Context     : …jumps over the [^^^]

────────────────────────────────────────────────────────────
Total: 2 item(s)
```

- **CHANGE** — a deletion immediately followed by an insertion from the same
  author is treated as one substitution.
- **DELETION** / **INSERTION** — standalone tracked changes.
- **COMMENT** — a comment anchored to highlighted text, with the comment body.
- `[^^^]` in the `Context` line marks where the change/comment sits relative
  to the surrounding paragraph text.

If the document has no tracked changes or comments, the tool prints
`No tracked changes or comments found.`

## How it works

`.docx` files are zip archives containing OOXML. This tool reads
`word/document.xml` (paragraphs, `w:ins`/`w:del` tracked-change markup, and
comment anchors) and `word/comments.xml` (comment bodies) directly via
`zipfile` and `xml.etree.ElementTree` — no third-party libraries required.
