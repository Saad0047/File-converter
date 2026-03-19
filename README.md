# File Converter

File Converter is a desktop application built with Python and Tkinter that converts files across common document, data, spreadsheet, image, presentation, and markup formats.

The app is designed for simple, safe, one-file-at-a-time conversion:
- Pick an input file
- Select a compatible output format
- Click Convert
- Get a new file saved in the same folder

Output files are created with this naming pattern:
original_name_converted.target_extension

---

## Features

- Desktop GUI with file picker
- Dynamic output format list based on selected file type
- Progress indicator and status messages
- Background conversion thread so the UI stays responsive
- Conversion success/error dialogs
- Startup dependency check with install guidance

---

## Supported Conversions

- DOCX -> PDF, TXT, HTML, MD
- PDF -> TXT, HTML, DOCX
- TXT -> PDF, DOCX, HTML, MD
- MD -> HTML, PDF, TXT, DOCX
- HTML -> TXT, PDF, MD, DOCX
- CSV -> JSON, XLSX, TXT, HTML
- JSON -> CSV, TXT, XLSX
- XLSX -> CSV, JSON, TXT
- PNG -> JPG, BMP, TIFF, WEBP, PDF
- JPG -> PNG, BMP, TIFF, WEBP, PDF
- JPEG -> PNG, BMP, TIFF, WEBP, PDF
- BMP -> PNG, JPG, TIFF, WEBP, PDF
- TIFF -> PNG, JPG, BMP, WEBP, PDF
- WEBP -> PNG, JPG, BMP, TIFF, PDF
- GIF -> PNG, JPG, BMP, PDF
- PPTX -> TXT, PDF
- XML -> JSON, TXT

---

## Tech Stack

- Python
- Tkinter (GUI)
- python-docx
- reportlab
- markdown
- beautifulsoup4
- lxml
- openpyxl
- Pillow
- PyPDF2
- python-pptx

---

## Installation

1. Install Python 3.8 or newer.
2. Install dependencies:

    pip install python-docx reportlab markdown beautifulsoup4 lxml openpyxl pillow pypdf2 python-pptx

3. Run the app:

    python file_converter.py

---

## How To Use

1. Launch the app.
2. Click Browse and select a supported file.
3. Choose an output format from the Convert to dropdown.
4. Click Convert.
5. Wait for the success message and open your converted file.

---

## How It Works

- The app detects the input extension.
- It looks up valid target formats from an internal format map.
- It routes conversion to a specific converter function.
- It writes output to the same directory as the source file.
- It handles errors and shows readable messages in the GUI.

---

## Important Notes And Limitations

- This tool focuses on practical content conversion, not perfect visual fidelity.
- Complex formatting may not be preserved across all format pairs.
- PDF text extraction depends on the PDF content structure.
- PDF to DOCX conversion is basic text reconstruction.
- PPTX to PDF is text-based extraction, not full slide rendering.
- JSON to CSV works best with list-of-objects style JSON data.
- Converting to an existing output filename will overwrite that file.

---

## Troubleshooting

- If dependencies are missing, the app shows a startup warning and the packages to install.
- If a conversion pair is unsupported, the app reports that format combination as unsupported.
- If conversion fails, the app shows the exact error message in a dialog.

---

## Project Structure

- file_converter.py: Main application, conversion logic, and GUI
- README.md: Project documentation
- LICENSE: License file

---

## License

This project is licensed under the terms in LICENSE.