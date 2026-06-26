# Invoice Extractor

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py invoice.pdf
```

Or drag one or more PDFs onto the packaged executable.

The script:
- Runs Marker
- Reads JSON tables
- Detects the CS column automatically
- Keeps only rows where CS > 0
- Preserves all extracted columns
- Writes output/Extracted.xlsx
