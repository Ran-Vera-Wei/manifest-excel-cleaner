# Manifest Excel Cleaner

A Python application that runs locally or on Streamlit Community Cloud. It accepts `.xlsx` and `.xlsm` files, processes them in memory, and provides the result for download in the browser. It does not write the output file to the server's disk.

## Features

- Enter a three-character destination code. The app uses the matching destination password only when the uploaded workbook requires an opening password.
- Remove duplicate rows using column C, `BG Number`, keeping the first occurrence. Blank BG Numbers are not deduplicated.
- Remove ALCOHOL, ACETONE, ESSENTIAL OIL, PERFUME, PERFUMING, NAIL ART, and NAIL POLISH from column O, `PRODUCT DESCRIPTION`, while keeping the rest of each description and its row.
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

Upload a manifest, enter its three-character destination code (the example workbook uses `SAL`), enter the export AWB#, and click **Process and generate file**. If the workbook is not encrypted, the destination code is not used to unlock it.

Duplex printing depends on the printer and its driver. An Excel workbook cannot reliably select the printer's duplex option for the user. Enable two-sided printing in the print dialog after downloading the file.
