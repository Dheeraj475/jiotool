# JioMart PDF to SCAT Bulk Converter

## Dashboard setup
1. Install Python 3.10 or newer.
2. Open Command Prompt in this folder.
3. Run: `pip install -r requirements.txt`
4. Run: `streamlit run jio_scat_bulk_converter.py`
5. Upload the approved `JIO Scat format.xlsx` template and all invoice PDFs.
6. Select the posting date, click **Generate SCAT Excel**, review any warnings, and download the output.

## Command-line option
`python jio_scat_bulk_converter.py --input ./pdfs --template "JIO Scat format.xlsx" --output JIO_SCAT_Bulk_Output.xlsx --posting-date 27.10.2026`

## Current mapping
- Vendor: 226222
- Business Place: MRA
- Section Code: EMHO
- WH Type / WH Code: I6
- GL Account: 510484
- Tax Code: XD
- Cost Center: CORP91
- Order: OTHERS1000
- SAC rows: 996211, 996819, 999794
- Value Date: invoice month-end

Update these constants in `build_workbook()` if the SAP master-data mapping changes.

## Controls
- Duplicate invoice numbers are blocked.
- Invoice number, date, taxable total, grand total, and SAC-wise taxable values are extracted.
- SAC total is compared with invoice taxable total; differences appear as review warnings.
- PDFs that cannot be read are excluded and shown as errors.
