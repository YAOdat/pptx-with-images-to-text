# PPTX Text Extractor

Extract text from PowerPoint presentations including OCR for images.

## Install

```bash
pip install -r requirements.txt
```

**Note:** For OCR functionality, you'll also need to install Tesseract:
- Ubuntu/Debian: `sudo apt-get install tesseract-ocr tesseract-ocr-ara`
- macOS: `brew install tesseract`
- Windows: Download from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)

## Web Service (Recommended)

Run the Flask web service for easy file upload and download:

```bash
python3 app.py
```

Then open your browser to `http://localhost:5000` or `http://0.0.0.0:5000`

The web interface allows you to:
- Upload PPTX files via drag-and-drop or file picker
- Configure OCR languages
- Download the extracted text file automatically

To host on the internet, you can deploy to:
- Heroku
- AWS Elastic Beanstalk
- Google Cloud Run
- Digital Ocean App Platform
- Railway
- Fly.io

## Command Line Usage

For command-line use:

```bash
python3 extract_pptx_ocr.py presentation.pptx
python3 extract_pptx_ocr.py presentation.pptx output.txt
python3 extract_pptx_ocr.py presentation.pptx output.txt eng+ara
```

## Output

Text extracted to `pptx_text_output.txt` (or custom filename) with:
- Text boxes and tables
- OCR from images  
- Speaker notes

## Features

- ✅ Extract text from text boxes
- ✅ Extract text from tables
- ✅ OCR text from images
- ✅ Extract speaker notes
- ✅ Multi-language OCR support
- ✅ Web interface for easy upload/download
- ✅ Stateless operation (no database required)
