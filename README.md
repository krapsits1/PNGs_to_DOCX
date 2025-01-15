# Text Extraction Program

This Python program extracts text from `.png` image files in a specified folder using Tesseract OCR and saves the results into a `.docx` file. The program supports Latvian (`lav`) language OCR.

## Features
- Processes `.png` files in a specified folder.
- Extracts text using Tesseract OCR with support for Latvian.
- Saves extracted text to a Word document with organized headings for each file.
- Includes error handling for file processing and output generation.

## Requirements

### Python Libraries
- `pytesseract`
- `Pillow`
- `python-docx`

Install these libraries using:
```bash
pip install pytesseract pillow python-docx
```

### Tesseract OCR
1. Install Tesseract OCR from [Tesseract GitHub](https://github.com/tesseract-ocr/tesseract).
2. Ensure the Latvian language data file (`lav.traineddata`) is installed. Download it from [tessdata](https://github.com/tesseract-ocr/tessdata) and place it in the Tesseract `tessdata` directory:
   - On Windows: `C:\Program Files\Tesseract-OCR\tessdata`
   - On Linux/Mac: `/usr/share/tesseract-ocr/4.00/tessdata`

## How to Use

1. **Set Up the Input Folder**
   - Place all `.png` files in the folder named `vt yo` (or modify the `input_folder` variable in the script to your preferred folder name).

2. **Run the Script**
   Execute the script in your Python environment:
   ```bash
   python script_name.py
   ```

3. **View the Output**
   - The extracted text will be saved in a file named `extracted_text.docx` in the same directory as the script.

## Customization
- **Change Input Folder**: Update the `input_folder` variable to specify a different folder containing `.png` files.
- **Change Output File**: Update the `output_docx` variable to modify the name or location of the output file.
- **Language Support**: To use a different language, modify the `lang` parameter in the `pytesseract.image_to_string` function.

## Error Handling
- If the input folder does not exist, the program will exit with an error message.
- Any issues with individual files will be logged, and the program will continue processing other files.

## Example
### Input Folder Structure
```
vt yo/
  ├── image1.png
  ├── image2.png
  ├── image3.png
```

### Output Word Document
```
extracted_text.docx
  ├── Heading: image1.png
  ├── Text extracted from image1.png
  ├── Heading: image2.png
  ├── Text extracted from image2.png
  ├── Heading: image3.png
  ├── Text extracted from image3.png
```

## Notes
- Ensure Tesseract OCR is correctly installed and accessible in your system's PATH. If not, set the Tesseract path explicitly in the script:
  ```python
  pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
  ```
