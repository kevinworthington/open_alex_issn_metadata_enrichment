# OpenAlex ISSN Metadata

A command-line Python tool that looks up journal/source metadata for ISSNs using the OpenAlex API and writes it to an Excel workbook. It can also continue an existing workbook: all existing rows are retained, missing metadata is filled when possible, and incomplete saved records are re-queried even if their ISSNs are not in the current input file.

## Requirements

- Python 3.10+
- Internet access
- An OpenAlex API key is optional

## Install

```bash
python -m venv .venv
```

Activate the environment, then install dependencies:

**macOS/Linux**
```bash
source .venv/bin/activate
pip install -r requirements.txt
```

**Windows PowerShell**
```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Input

Use a `.txt` file with one ISSN per line, or a `.csv` file with an ISSN in the first column. CSV files may have a header named `issn`.

```text
0378-5955
0028-0836
```

A sample list is included as `sample_issns.txt`.

## Run

```bash
python openalex_issn.py sample_issns.txt --output metadata.xlsx
```

Supported terminal verbosity levels:

```bash
python openalex_issn.py issns.txt -o metadata.xlsx --verbosity quiet
python openalex_issn.py issns.txt -o metadata.xlsx --verbosity normal
python openalex_issn.py issns.txt -o metadata.xlsx --verbosity verbose
python openalex_issn.py issns.txt -o metadata.xlsx --verbosity debug
```

Shortcuts: `-v` enables verbose output; `-vv` enables debug output. Default is `normal`.

## Optional API key

Pass a key directly:

```bash
python openalex_issn.py issns.txt -o metadata.xlsx --api-key YOUR_KEY
```

Or set `OPENALEX_API_KEY` (preferred to avoid putting the key in shell history):

**macOS/Linux**
```bash
export OPENALEX_API_KEY="YOUR_KEY"
python openalex_issn.py issns.txt -o metadata.xlsx
```

**Windows PowerShell**
```powershell
$env:OPENALEX_API_KEY="YOUR_KEY"
python openalex_issn.py issns.txt -o metadata.xlsx
```

The key is only sent to OpenAlex when supplied. Without one, the public endpoint is used.

## Existing workbook behavior

If the output `.xlsx` already exists, its rows are loaded instead of being discarded. The script:
- retains existing records, including records not listed in the current input;
- looks up each input ISSN and merges matching records;
- fills blank fields in matching records without overwriting populated metadata;
- re-queries incomplete existing records based on their saved ISSNs, even if they are absent from the current input;
- preserves extra columns from the existing workbook.

Use a different output path if you want a separate workbook. Close the workbook in Excel before running the script.

## Output and status

The workbook has one row per OpenAlex source record, with input ISSNs, OpenAlex identifiers, source title, ISSN values, publisher, source type, country, homepage, counts, OA/DOAJ flags, dates, and summary fields. Unmatched ISSNs get a `not_found` row. Request failures are reported in the terminal and incomplete affected rows are marked `error`; the existing data remains.

The script normalizes eight-character ISSNs to `NNNN-NNNN`, retries transient request errors, and writes spreadsheet-safe text values.

## Files

- `openalex_issn.py` — command-line application
- `requirements.txt` — Python dependencies
- `sample_issns.txt` — example input
