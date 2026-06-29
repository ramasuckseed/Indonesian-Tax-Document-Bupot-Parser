# Indonesian Tax Document Parser

A command-line application for parsing Indonesian tax documents in bulk, including **Bukti Potong PPh 21**, **Bukti Potong PPh 23/Unifikasi**, and **Tax Invoices (PPN)** into structured Excel reports.

The application automatically detects document types, extracts important tax information, and generates consolidated reports for accounting and tax reconciliation.

---

## Features

- Parse hundreds of PDF documents in one execution
- Automatic document type detection
- Support for:
  - Bukti Potong PPh 21
  - Bukti Potong PPh 23 / PPh Unifikasi
  - Tax Invoice (PPN Keluaran)
  - Tax Invoice (PPN Masukan)
- Extract tax-related information automatically
- Generate Excel reports
- Separate worksheets by document type
- Processing summary
- Error handling for unsupported or scanned PDFs

---

## Extracted Information

### Bukti Potong

- Document Number
- Tax Period
- Taxpayer Name
- Taxpayer NPWP
- Withholding Agent
- Tax Object Code
- Gross Income
- Tax Base (DPP)
- Tax Rate
- Withholding Tax Amount

### Tax Invoice

- Invoice Number
- Invoice Date
- Tax Period
- Seller
- Buyer
- Seller NPWP
- Buyer NPWP
- Product / Service
- DPP
- VAT (PPN)
- Reference Number

---

## Requirements

- Python 3.10+
- pandas
- openpyxl
- pypdf

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

Edit the input directory inside the configuration.

```python
FOLDER_PDF = r"E:\Faktur Pajak All"
```

Run:

```bash
python extract_faktur.py
```

The application will automatically:

- Scan all PDF files
- Detect document types
- Parse tax information
- Generate an Excel report

Output:

```
rekap_faktur.xlsx
```

---

## Output Example

```
rekap_faktur.xlsx

├── Semua
├── PPh21
├── PPh23
├── PPN_Keluaran
└── PPN_Masukan
```

---

## Technologies

- Python
- pandas
- pypdf
- openpyxl
- Regular Expressions (Regex)

---

## Project Structure

```
extract_faktur.py
```

---

## Limitations

- Supports digitally generated PDF documents.
- Scanned PDF files without selectable text are not supported.
- Parsing rules are based on the document layouts available during development and may require updates if official document formats change.

---

## Future Improvements

- XML support
- ZIP archive processing
- CSV export
- Duplicate detection
- Command-line arguments
- Recursive folder scanning
- Logging
- OCR support for scanned PDFs

---

## Disclaimer

This software is intended to assist tax and accounting professionals in extracting information from Indonesian tax documents.

Users should always verify the generated results before using them for tax reporting or regulatory compliance.

---

## License

MIT License
