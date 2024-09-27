# PyFileConvert

A comprehensive Python script for file format conversion, supporting a wide range of document, image, audio, video, and data formats. It utilizes various libraries to perform conversions between different file types, including PDF, Word documents, images, spreadsheets, audio files, video files, and more. The script also includes features for compression, extraction, and batch processing.

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/RishikSarkar/file-converter.git
   cd file-converter
   ```

2. Create a virtual environment (optional but recommended):
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

4. Install additional system dependencies:
   - FFmpeg (for audio/video conversions)
   - Pandoc (for markdown conversions)
   - Calibre (for e-book conversions)

## Usage

Run the script:
```
python fileconvert.py
```

Follow the prompts to enter input file(s) and output file. Type 'h' for help or 'q' to quit.

### Examples

1. Convert a single file:
   ```
   Enter input file(s): document.docx
   Enter output file: document.pdf
   ```

2. Convert multiple images to PDF:
   ```
   Enter input file(s): image1.jpg image2.png image3.tiff
   Enter output file: combined.pdf
   ```

3. Extract audio from video:
   ```
   Enter input file(s): video.mp4
   Enter output file: audio.mp3
   ```

4. Batch conversion (convert all files in a directory):
   ```
   Enter input file(s): input_directory
   Enter output file: output_directory
   ```

## Supported Formats

- Documents: PDF, DOCX, TXT, MD, HTML, EPUB
- Images: PNG, JPG, JPEG, TIFF, HEIC, WEBP, GIF, BMP
- Spreadsheets: XLSX, CSV
- Data: JSON, YAML, XML
- Archives: ZIP, RAR, 7Z
- Video: MP4, MOV, AVI, MKV, WEBM, FLV, WMV
- Audio: MP3, WAV, OGG, FLAC, AAC, M4A, WMA

## Notes

- Multiple file conversion is only supported for PDF output.
- Some conversions may require additional system dependencies.
- Use the 'h' command when prompted for input files to see all supported formats and possible conversions.