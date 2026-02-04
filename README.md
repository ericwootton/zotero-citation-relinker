# Zotero Citation Relinker

Automatically relink orphaned Zotero citations in Word (`.docx`) documents using fuzzy matching against your local Zotero library.

## The Problem

When migrating from Mendeley to Zotero, collaborating across libraries, or syncing between machines, Zotero citations in Word documents can become "orphaned" -- the embedded URIs no longer point to valid items in your library. This breaks Zotero's ability to manage, update, or restyle those citations.

Manually relinking hundreds of citations in a mydocument or dissertation is painful. This script automates it.

## How It Works

1. **Extracts** all Zotero citation fields from the `.docx` XML
2. **Checks** each citation's URI against your local Zotero SQLite database
3. **Matches** orphaned citations to library items using a tiered strategy:
   - **DOI** (exact match, 100% confidence)
   - **ISBN** (exact match, 100% confidence)
   - **Fuzzy** title + author + year matching (configurable threshold)
   - **Title-only** fuzzy matching (90% threshold fallback)
4. **Replaces** orphaned URIs directly in the document XML with valid Zotero URIs
5. **Outputs** a new `.docx` with relinked citations and a detailed text report

The script handles Mendeley-format URIs, Zotero user/local/group URIs, and multi-URI citation items.

## Installation

### Requirements

- Python 3.8+
- [rapidfuzz](https://github.com/rapidfuzz/RapidFuzz)

### Install

```bash
pip install rapidfuzz
```

That's it -- no other dependencies beyond the Python standard library.

## Usage

### Analyze a document (report only)

```bash
python zotero_relinker.py mydocument.docx
```

This prints a report showing orphaned citations and potential matches without modifying anything.

### Generate a relinked document

```bash
python zotero_relinker.py mydocument.docx --output mydocument_fixed.docx
```

### Specify a custom Zotero data directory

```bash
python zotero_relinker.py mydocument.docx --zotero-path /path/to/Zotero
```

### Lower the match threshold for more matches

```bash
python zotero_relinker.py mydocument.docx --output mydocument_fixed.docx --threshold 70
```

### Generate a manual relinking guide

```bash
python zotero_relinker.py mydocument.docx --guide manual_guide.txt
```

### Full example

```bash
python zotero_relinker.py mydocument.docx \
  --output mydocument_fixed.docx \
  --zotero-path ~/Zotero \
  --threshold 75 \
  --guide manual_guide.txt
```

## Command-Line Options

| Option | Description |
|---|---|
| `docx_file` | Path to the Word document (required) |
| `--output, -o` | Output path for the relinked document |
| `--zotero-path, -z` | Path to Zotero data directory or `zotero.sqlite` |
| `--threshold, -t` | Fuzzy match threshold, 0-100 (default: 80) |
| `--guide, -g` | Generate a manual relinking guide to the given file |

## Output

The script always generates a text report (`<filename>_relink_report.txt`) summarizing:

- Total citations and items in the document
- Number of orphaned citations
- Match results for each orphaned item (method, score, matched library entry)

Example report snippet:

```
Total citations in document: 558
Total citation items: 919
Orphaned citations: 491
Orphaned items: 757

[1.1] ORPHANED CITATION:
    Title:   Non-exhaust traffic emissions: Sources, characterization...
    Authors: Piscitello Bianco Casasso Sethi
    Year:    2021
    DOI:     10.1016/J.SCITOTENV.2020.144440

    [MATCH] POTENTIAL MATCH FOUND (DOI, score: 100%):
      Library Key: ZC9XPU5Q
      Title:       Non-exhaust traffic emissions: Sources, characterization...
      Authors:     Piscitello Bianco Casasso Sethi
      Year:        2021
```

## How Zotero Citations Work in Word

Zotero stores citations as Word field codes containing JSON with CSL (Citation Style Language) data. Each citation item includes:

- **URIs** linking to the Zotero library item (e.g., `http://zotero.org/users/12345/items/ABC123`)
- **Metadata** (title, authors, year, DOI, etc.)

When URIs point to items that no longer exist in your library, Zotero marks them as orphaned. This script replaces those dead URIs with ones pointing to the correct items, identified through metadata matching.

## Zotero Database Location

The script auto-detects your Zotero database at standard locations:

| OS | Path |
|---|---|
| Windows | `%USERPROFILE%\Zotero\zotero.sqlite` |
| macOS | `~/Library/Application Support/Zotero/zotero.sqlite` |
| Linux | `~/.zotero/zotero/zotero.sqlite` or `~/Zotero/zotero.sqlite` |

Use `--zotero-path` to override if your database is elsewhere.

## Important Notes

- **Non-destructive**: The original document is never modified. A new file is created.
- **Close Zotero first**: The script copies the database to a temp file to avoid locking issues, but closing Zotero ensures a consistent snapshot.
- **Verify after relinking**: Open the output document in Word with Zotero running and click *Refresh* in the Zotero toolbar to verify the relinked citations.
- **Backup your document** before opening the relinked version in Word.

## License

MIT
