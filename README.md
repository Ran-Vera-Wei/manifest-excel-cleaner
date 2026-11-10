# CBP Form 7512 PDF Editor and Manifest Excel Cleaner

This repository contains two in-memory Streamlit apps. The PDF editor is the app for CBP Form 7512 editing; the Excel cleaner remains available as a separate entry point.

## CBP Form 7512 PDF Editor

`pdf_app.py` accepts one or more one-page, portrait Letter-size CBP Form 7512 PDFs. Enter shipment details once and the app applies them to every file; a single result downloads as a PDF, while multiple results download together as a ZIP. It renders each page at 300 dpi, clears the red-marked values, then places the replacement values at the green numbered locations. The source page is flattened to an image so removed content is not left hidden underneath; the output is print-ready, but the original form text is no longer selectable. Processing is in memory and the app does not save uploaded or output files to the server's disk.

- Required inputs: Export AWB #, three-character destination code, Consignee selection, CN/HK port selection, number of packages, gross weight in KG, and value in dollars.
- Choose a client first. Destination codes and consignee choices are then filtered by that client using `client_consignees.json`. TEMU retains the 18 consignee options from the supplied customs workbook; J&T EXPRESS and BDT use their client-specific destinations and consignees.
- The entire original Consignee text area is cleared while its underline is preserved. Replacement text is fixed at 7 pt, wraps to two lines, and is truncated with an ellipsis if it is longer.
- #1 Export AWB # is placed in Car number and initial. #2 is the selected Consignee. #3 is fixed to US. #4 maps CN to `57000 - CHINA(MAINLAND)` and HK to `HONG KONG - 58201`, and also copies the selected value to the Last foreign port field. #5 is fixed to `SHOES, FOOTWEAR WOMEN'S DRESSES MEN'S SHIRTS`. #6-8 are package count, gross weight, and value.
- Existing B/L No. and the Exported from row remain as they are in each uploaded PDF. Every file must be a one-page portrait Letter-size CBP 7512 matching the reference layout.

### Run the PDF editor locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run pdf_app.py
```

### Deploy the PDF editor to Streamlit Community Cloud

The deployed app is public. `client_consignees.json` contains the client, destination, and consignee directory, and the app exposes choices derived from it. The directory is public by explicit approval. Create a Streamlit Community Cloud app using `pdf_app.py` as the main file path. Do not add shipment-specific PDFs or secrets to the public deployment repository.

## Manifest Excel Cleaner

A Python application that runs locally or on Streamlit Community Cloud. It accepts `.xlsx`, `.xlsm`, and legacy `.xls` files, processes them in memory, and provides the result for download in the browser. It does not write the output file to the server's disk. Legacy `.xls` input is converted to `.xlsx` output.

## Features

- Choose a client: **TEMU**, **JTExpress**, or **BDT**.
- Remove fully blank rows from each matched manifest worksheet before deduplicating rows. Cells containing only whitespace count as blank.
- TEMU keeps its current rules: remove duplicates by column C, `BG Number`, and remove sensitive terms from column O, `PRODUCT DESCRIPTION`.
- JTExpress removes columns whose only non-empty value is the header, removes duplicate rows by `TRACKING NO.`, and removes sensitive terms from `CONTENT DESCRIPTION` (or `PRODUCT DESCRIPTION` when present).
- BDT removes columns whose only non-empty value is the header, removes duplicate rows by `Tracking Number`, and removes sensitive terms from `PRODUCT DESCRIPTION`.
- For every client, blank deduplication keys are kept; the first occurrence of each non-blank key is retained.
- Enter a three-character destination code. The app uses the matching destination password only when the uploaded workbook requires an opening password.
- Remove ALCOHOL, ACETONE, ESSENTIAL OIL, PERFUME, PERFUMING, NAIL ART, and NAIL POLISH from the selected description column, keeping the rest of each description and its row.
- Set the font size of used cells in the workbook to 16 pt.
- Set landscape orientation and narrow margins, and fit all columns to one page wide. Long worksheets can continue across multiple pages.
- Name the download `<AWB#>manifest.xlsx` based on the export AWB# (`.xlsm` workbooks retain the `.xlsm` extension).

## Run locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .streamlit\secrets.toml.example .streamlit\secrets.toml
```

Open `.streamlit/secrets.toml` and enter the destination passwords in the corresponding entries. This file is included in `.gitignore`; do not commit actual passwords to a code repository.

```powershell
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

Publish `app.py`, `requirements.txt`, and this README to a GitHub repository. Then select the repository and `app.py` in Streamlit Community Cloud. Under the app's **Settings → Secrets**, add the `[manifest_passwords]` section and passwords using the format shown in `.streamlit/secrets.toml.example`. Do not put real passwords in Python source files, this README, or a public repository.

## How to use

Choose the client, upload a manifest, enter its three-character destination code (the example encrypted workbook uses `SAL`), enter the export AWB#, and click **Process and generate file**. If the workbook is not encrypted, the destination code is not used to unlock it.

Duplex printing depends on the printer and its driver. An Excel workbook cannot reliably select the printer's duplex option for the user. Enable two-sided printing in the print dialog after downloading the file.
