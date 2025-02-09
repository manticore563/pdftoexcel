# PDF to Excel Converter

A robust web application for converting PDF files to Excel format with advanced table parsing capabilities.

## Features

- Modern, user-friendly interface
- Advanced table extraction from PDFs
- Multi-page PDF support
- Secure file processing
- Automatic cleanup of temporary files
- Professional Excel formatting with proper styling

## Requirements

- Python 3.11
- Flask
- Pandas
- Tabula-py
- OpenPyXL

## Installation

1. Clone the repository:
```bash
git clone https://github.com/manticore563/pdftoexcel.git
cd pdftoexcel
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python app.py
```

The application will be available at `http://localhost:5000`

## Usage

1. Upload a PDF file by dragging and dropping or using the file browser
2. Click "Convert to Excel"
3. Download the generated Excel file

## License

MIT License
