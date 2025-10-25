# PPTX Text Extractor

Extract text from PowerPoint presentations including OCR for images.

## Install

```bash
pip install python-pptx pillow pytesseract
```

## Usage

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
