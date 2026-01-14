# German Tax Invoice Extraction System

A Python-based system to automatically extract tax-relevant information from German vendor PDF invoices and maintain an Excel tracking sheet for easy tax filing.

## Features

- **Automatic PDF Processing**: Extract key information from German invoices (Rechnungen)
- **Excel Tracking**: Maintain a single Excel file with all invoice data
- **Smart Filing**: Automatically organize invoices by vendor in archive folders
- **German Tax Format Support**: Handles German date formats, currency, VAT rates, and umlauts
- **Quality Indicators**: Color-coded rows show extraction confidence
- **Duplicate Detection**: Skip already processed invoices
- **Batch Processing**: Process multiple invoices at once

## Extracted Information

The system extracts:
- Invoice number (Rechnungsnummer)
- Invoice date (Rechnungsdatum)
- Vendor name (Lieferant)
- Net amount (Nettobetrag)
- VAT rate (MwSt.-Satz: 19%, 7%, etc.)
- VAT amount (MwSt.-Betrag)
- Gross total (Bruttobetrag)
- Expense category (for manual classification)

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Setup Steps

1. **Clone the repository**:
   ```bash
   git clone https://github.com/derkaal/bill_analyzer.git
   cd bill_analyzer
   ```

2. **Create a virtual environment** (recommended):
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Verify installation**:
   ```bash
   python invoice_extractor.py --help
   ```

## Usage

### Quick Start

1. **Place PDF invoices** in the `./new/` folder (created automatically on first run)

2. **Run the extraction**:
   ```bash
   python invoice_extractor.py process
   ```

3. **Check the results**:
   - Open `tax_records.xlsx` to see extracted data
   - Processed PDFs are moved to `./archive/<vendor_name>/`

### Available Commands

#### Process Invoices
```bash
python invoice_extractor.py process
```
Processes all PDF files in the `./new/` folder:
- Extracts invoice data
- Adds to Excel tracking sheet
- Moves PDFs to vendor-specific archive folders
- Skips files already processed

#### View Report
```bash
python invoice_extractor.py report
```
Displays summary statistics:
- Total number of invoices
- Total gross amount
- Breakdown by vendor
- Breakdown by month
- Extraction status summary

#### List Pending Files
```bash
python invoice_extractor.py list
```
Lists all PDF files in `./new/` folder without processing them.

## Folder Structure

After first run, the following structure is created:

```
bill_analyzer/
├── invoice_extractor.py      # Main application
├── requirements.txt           # Python dependencies
├── README.md                  # This file
├── .gitignore                 # Git ignore rules
├── new/                       # ⬅️ Drop your invoices here
│   ├── invoice1.pdf
│   └── invoice2.pdf
├── archive/                   # Processed invoices (organized by vendor)
│   ├── telekom_deutschland_gmbh/
│   │   └── invoice1.pdf
│   └── deutsche_post_ag/
│       └── invoice2.pdf
├── tax_records.xlsx           # Excel tracking sheet
└── extraction_log.txt         # Processing log file
```

**Note**: The `new/`, `archive/`, `*.xlsx`, and `extraction_log.txt` are git-ignored as they contain user data.

## Excel Sheet Format

The `tax_records.xlsx` file contains:

| Column | Description | Format |
|--------|-------------|--------|
| Filename | Original PDF filename | Text |
| Date | Invoice date | DD.MM.YYYY |
| Vendor | Vendor/supplier name | Text |
| Invoice_Number | Invoice number | Text |
| Net | Net amount | #,##0.00 € |
| VAT_Rate | VAT percentage | 19% |
| VAT_Amount | VAT amount | #,##0.00 € |
| Gross | Total amount | #,##0.00 € |
| Category | Expense category (dropdown) | Text |
| Extraction_Status | Confidence level | OK / UNCERTAIN / MANUAL_REVIEW_NEEDED |
| Notes | Additional information | Text |

### Color Coding

- **White**: Extraction successful (OK)
- **Yellow**: Uncertain extraction - review recommended
- **Red**: Manual review needed - extraction failed or incomplete

### Expense Categories

The Category column includes a dropdown with common German expense types:
- Büromaterial (Office supplies)
- Software
- Reisekosten (Travel expenses)
- Marketing
- Telefon/Internet
- Miete (Rent)
- Versicherung (Insurance)
- Weiterbildung (Training)
- Beratung (Consulting)
- Sonstiges (Other)

## How It Works

### Extraction Process

1. **Text Extraction**: Uses `pdfplumber` to extract text from PDF
2. **Pattern Matching**: Identifies German invoice fields using regex patterns
3. **Vendor Detection**: Finds company names (looks for GmbH, AG, etc.)
4. **Smart Amount Extraction**: Uses keyword-based context search
   - **Gross Total**: Searches near keywords like "SUMME EUR", "BRUTTO", "Gesamtbetrag"
   - **Net Amount**: Searches near keywords like "NETTO", "Nettobetrag", "Summe Netto"
   - **VAT Amount**: Searches near keywords like "MwSt", "Mehrwertsteuer", "USt"
   - This prevents accidentally extracting line item prices instead of totals
   - Fallback to largest amount if keyword search fails (marked as UNCERTAIN)
5. **Validation**: Checks if Net + VAT ≈ Gross (tolerance: €0.02)
6. **Quality Assessment**: Assigns confidence level to extraction

### Vendor Name Sanitization

Vendor names are sanitized for use as folder names:
- Converted to lowercase
- German umlauts replaced: ä→ae, ö→oe, ü→ue, ß→ss
- Spaces and special characters replaced with underscores
- Example: "Telekom Deutschland GmbH" → "telekom_deutschland_gmbh"

### Duplicate Prevention

The system checks existing Excel records before processing. If a filename already exists, the file is skipped.

## Troubleshooting

### No text extracted from PDF
**Problem**: Some PDFs are scanned images without text layer.
**Solution**: Currently not supported. Future version will include OCR. For now, manually enter data.

### Wrong amounts extracted
**Problem**: Extraction marked as "UNCERTAIN" or amounts don't match.
**Solution**:
1. Check the PDF - it may have unusual formatting
2. Manually correct values in Excel
3. The system validates Net + VAT = Gross

### Vendor name not found
**Problem**: Vendor shows as "Unknown Vendor".
**Solution**: Manually edit the vendor name in Excel. The archive folder can be renamed too.

### Permission errors
**Problem**: "Permission denied" when moving files.
**Solution**:
1. Close any PDF viewers that have the file open
2. Check file permissions
3. On Windows, ensure the file isn't locked

### Processing hangs
**Problem**: Script seems stuck on a PDF.
**Solution**:
1. Press Ctrl+C to cancel
2. Check `extraction_log.txt` for error details
3. Remove or move the problematic PDF
4. Run process command again

## Log Files

Processing details are logged to `extraction_log.txt`:
```
[2026-01-14 10:15:23] INFO: Processed invoice1.pdf -> archive/telekom_deutschland_gmbh/invoice1.pdf
[2026-01-14 10:15:28] WARNING: Vendor name unclear for invoice2.pdf
[2026-01-14 10:15:30] ERROR: Failed to process invoice3.pdf: Invalid PDF
```

## Tips for Best Results

1. **Use digital PDFs**: Scanned PDFs with text layer work better than pure image scans
2. **Process regularly**: Don't let invoices pile up - process weekly or monthly
3. **Review uncertain extractions**: Check yellow-highlighted rows in Excel
4. **Complete categories**: Add expense categories manually for better tax filing
5. **Backup Excel file**: Regularly back up `tax_records.xlsx`

## Limitations

- **No OCR**: Scanned invoices without text layer cannot be processed
- **German invoices only**: Optimized for German invoice formats
- **Generic extraction**: No vendor-specific templates (yet)
- **Manual categories**: Expense categories must be selected manually

## Future Enhancements

- OCR support for scanned invoices
- Vendor-specific extraction templates
- Automatic category classification using AI
- Export to DATEV format for tax advisors
- Web interface for easier use
- Multi-language support

## Requirements

- Python 3.8+
- pdfplumber 0.11.0
- openpyxl 3.1.2

See `requirements.txt` for exact versions.

## License

This project is provided as-is for personal and commercial use.

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review `extraction_log.txt` for error details
3. Open an issue on GitHub

## Contributing

Contributions welcome! Please open an issue first to discuss proposed changes.

---

**Note**: This system helps organize invoices but does not replace professional tax advice. Always consult with a tax advisor (Steuerberater) for tax filing.
