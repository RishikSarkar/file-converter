import os
from PIL import Image
import fitz
import docx2pdf
import csv
import json
import yaml
import pandas as pd
import xml.etree.ElementTree as ET
import subprocess
from docx import Document
from PyPDF2 import PdfReader
import markdown
import zipfile
import rarfile
import py7zr
import pdf2image
from pdf2docx import Converter
import pdfkit
from moviepy.editor import VideoFileClip
from typing import Dict, List
import logging
import pkg_resources

def get_project_requirements():
    with open('requirements.txt', 'w') as f:
        for package in pkg_resources.working_set:
            f.write(f"{package.key}=={package.version}\n")

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def check_ffmpeg():
    try:
        subprocess.run(
            ["ffmpeg", "-version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


ffmpeg_available = check_ffmpeg()

if ffmpeg_available:
    import ffmpeg
    from pydub import AudioSegment


def excel_to_csv(excel_file, csv_file):
    df = pd.read_excel(excel_file)
    df.to_csv(csv_file, index=False)


def csv_to_excel(csv_file, excel_file):
    df = pd.read_csv(csv_file)
    df.to_excel(excel_file, index=False)


def csv_to_json(csv_file, json_file):
    with open(csv_file, "r") as csvfile:
        reader = csv.DictReader(csvfile)
        data = list(reader)

    with open(json_file, "w") as jsonfile:
        json.dump(data, jsonfile, indent=4)


def json_to_csv(json_file, csv_file):
    with open(json_file, "r") as jsonfile:
        data = json.load(jsonfile)

    with open(csv_file, "w", newline="") as csvfile:
        if data:
            writer = csv.DictWriter(csvfile, fieldnames=data[0].keys())
            writer.writeheader()
            writer.writerows(data)


def json_to_yaml(json_file, yaml_file):
    with open(json_file, "r") as jsonfile:
        data = json.load(jsonfile)

    with open(yaml_file, "w") as yamlfile:
        yaml.dump(data, yamlfile, default_flow_style=False)


def yaml_to_json(yaml_file, json_file):
    with open(yaml_file, "r") as yamlfile:
        data = yaml.safe_load(yamlfile)

    with open(json_file, "w") as jsonfile:
        json.dump(data, jsonfile, indent=4)


def word_to_pdf(docx_file, pdf_file):
    docx2pdf.convert(docx_file, pdf_file)


def pdf_to_text(pdf_file, txt_file):
    with open(pdf_file, "rb") as file:
        reader = PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"

    with open(txt_file, "w", encoding="utf-8") as file:
        file.write(text)


def text_to_word(txt_file, docx_file):
    doc = Document()
    with open(txt_file, "r", encoding="utf-8") as file:
        doc.add_paragraph(file.read())
    doc.save(docx_file)


def convert_image(input_path, output_path):
    with Image.open(input_path) as img:
        if output_path.lower().endswith(".pdf"):
            img.save(output_path, "PDF", resolution=100.0)
        else:
            img.save(output_path)


def convert_pdf_to_image(input_path, output_path):
    images = pdf2image.convert_from_path(input_path)
    if images:
        # Save only the first page if the output is a single image file
        images[0].save(output_path)
    else:
        raise ValueError("No images extracted from the PDF")


def convert_pdf(input_path, output_path):
    doc = fitz.open(input_path)
    if output_path.lower().endswith((".png", ".jpg", ".jpeg", ".tiff")):
        page = doc.load_page(0)
        pix = page.get_pixmap()
        pix.save(output_path)
    else:
        doc.save(output_path)
    doc.close()


def convert_video(input_path, output_path):
    if ffmpeg_available:
        stream = ffmpeg.input(input_path)
        stream = ffmpeg.output(stream, output_path)
        ffmpeg.run(stream)
    else:
        raise ValueError("FFmpeg is not available. Video conversion is not supported.")


def convert_docx_to_pdf(input_path, output_path):
    docx2pdf.convert(input_path, output_path)


def convert_audio(input_path, output_path):
    if ffmpeg_available:
        audio = AudioSegment.from_file(input_path)
        audio.export(output_path, format=os.path.splitext(output_path)[1][1:])
    else:
        raise ValueError("FFmpeg is not available. Audio conversion is not supported.")


def convert_data_format(input_path, output_path):
    input_ext = os.path.splitext(input_path)[1].lower()
    output_ext = os.path.splitext(output_path)[1].lower()

    if input_ext in (".json", ".yaml", ".csv", ".xlsx"):
        if input_ext == ".json":
            df = pd.read_json(input_path)
        elif input_ext == ".yaml":
            with open(input_path, "r") as file:
                data = yaml.safe_load(file)
            df = pd.DataFrame(data)
        elif input_ext == ".csv":
            df = pd.read_csv(input_path)
        elif input_ext == ".xlsx":
            df = pd.read_excel(input_path)

        if output_ext == ".json":
            df.to_json(output_path, orient="records", indent=2)
        elif output_ext == ".yaml":
            with open(output_path, "w") as file:
                yaml.dump(df.to_dict(orient="records"), file)
        elif output_ext == ".csv":
            df.to_csv(output_path, index=False)
        elif output_ext == ".xlsx":
            df.to_excel(output_path, index=False)
    elif input_ext == ".xml" and output_ext == ".json":
        tree = ET.parse(input_path)
        root = tree.getroot()
        data = {root.tag: {}}
        for child in root:
            data[root.tag][child.tag] = child.text
        with open(output_path, "w") as file:
            json.dump(data, file, indent=2)
    else:
        raise ValueError(
            f"Unsupported data format conversion: {input_ext} to {output_ext}"
        )


def convert_svg(input_path, output_path):
    try:
        import cairosvg

        cairosvg.svg2png(url=input_path, write_to=output_path)
    except ImportError:
        print("CairoSVG is not installed. SVG conversion is not available.")
        raise ValueError("SVG conversion is not supported without CairoSVG")


def check_command(command):
    try:
        subprocess.run(
            [command, "--version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


pandoc_available = check_command("pandoc")
calibre_available = check_command("ebook-convert")


def markdown_to_html(md_file, html_file):
    with open(md_file, "r", encoding="utf-8") as f:
        md_content = f.read()
    html_content = markdown.markdown(md_content)
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_content)


def markdown_to_pdf(md_file, pdf_file):
    if pandoc_available:
        subprocess.run(["pandoc", md_file, "-o", pdf_file])
    else:
        raise ValueError(
            "Pandoc is not available. Markdown to PDF conversion is not supported."
        )


def html_to_markdown(html_file, md_file):
    if pandoc_available:
        subprocess.run(
            ["pandoc", "-f", "html", "-t", "markdown", "-o", md_file, html_file]
        )
    else:
        raise ValueError(
            "Pandoc is not available. HTML to Markdown conversion is not supported."
        )


def epub_to_pdf(epub_file, pdf_file):
    if calibre_available:
        subprocess.run(["ebook-convert", epub_file, pdf_file])
    else:
        raise ValueError(
            "Calibre is not available. EPUB to PDF conversion is not supported."
        )


def compress_zip(input_file, output_file):
    with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(input_file, os.path.basename(input_file))


def extract_zip(input_file, output_dir):
    with zipfile.ZipFile(input_file, "r") as zipf:
        zipf.extractall(output_dir)


def compress_rar(input_file, output_file):
    with rarfile.RarFile(output_file, "w") as rarf:
        rarf.add(input_file, os.path.basename(input_file))


def extract_rar(input_file, output_dir):
    with rarfile.RarFile(input_file, "r") as rarf:
        rarf.extractall(output_dir)


def compress_7z(input_file, output_file):
    with py7zr.SevenZipFile(output_file, "w") as szf:
        szf.write(input_file, os.path.basename(input_file))


def extract_7z(input_file, output_dir):
    with py7zr.SevenZipFile(input_file, "r") as szf:
        szf.extractall(output_dir)

def pdf_to_word(pdf_file, docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file)
    cv.close()

def images_to_pdf(image_files, pdf_file):
    images = []
    for image_file in image_files:
        img = Image.open(image_file)
        if img.mode == 'RGBA':
            # Convert RGBA images to RGB
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            img = background
        images.append(img)
    
    if images:
        images[0].save(pdf_file, "PDF", resolution=100.0, save_all=True, append_images=images[1:])
    else:
        raise ValueError("No valid images found to convert to PDF")

def html_to_pdf(html_file, pdf_file):
    pdfkit.from_file(html_file, pdf_file)

def extract_audio_from_video(video_file, audio_file):
    video = None
    audio = None
    try:
        video = VideoFileClip(video_file)
        audio = video.audio
        audio.write_audiofile(audio_file)
    finally:
        if audio:
            audio.close()
        if video:
            video.close()

def batch_convert(input_dir, output_dir, input_ext, output_ext):
    for filename in os.listdir(input_dir):
        if filename.endswith(input_ext):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, f"{os.path.splitext(filename)[0]}{output_ext}")
            try:
                convert_file(input_path, output_path)
                print(f"Converted {filename} successfully.")
            except Exception as e:
                print(f"Failed to convert {filename}: {str(e)}")

def convert_file(input_path, output_path):
    logging.info(f"Starting conversion: {input_path} -> {output_path}")
    try:
        if isinstance(input_path, list):
            input_ext = os.path.splitext(input_path[0])[1].lower()
        else:
            input_ext = os.path.splitext(input_path)[1].lower()
        output_ext = os.path.splitext(output_path)[1].lower()

        if isinstance(input_path, list) and output_path.lower().endswith('.pdf'):
            images_to_pdf(input_path, output_path)
        elif input_ext in (".png", ".jpg", ".jpeg", ".tiff") and output_ext == ".pdf":
            images_to_pdf([input_path], output_path)
        elif input_ext == ".docx" and output_ext == ".pdf":
            word_to_pdf(input_path, output_path)
        elif input_ext == ".pdf" and output_ext == ".txt":
            pdf_to_text(input_path, output_path)
        elif input_ext == ".txt" and output_ext == ".docx":
            text_to_word(input_path, output_path)
        elif input_ext in (
            ".png",
            ".jpg",
            ".jpeg",
            ".tiff",
            ".heic",
            ".webp",
            ".gif",
            ".bmp",
        ):
            convert_image(input_path, output_path)
        elif input_ext == ".pdf":
            convert_pdf(input_path, output_path)
        elif input_ext in (".mp4", ".mov", ".avi", ".mkv", ".webm", ".flv", ".wmv"):
            convert_video(input_path, output_path)
        elif input_ext == ".docx" and output_ext == ".pdf":
            convert_docx_to_pdf(input_path, output_path)
        elif input_ext in (".mp3", ".wav", ".ogg", ".flac", ".aac", ".m4a", ".wma"):
            convert_audio(input_path, output_path)
        elif input_ext in (".json", ".yaml", ".csv", ".xlsx", ".xml") and output_ext in (
            ".json",
            ".yaml",
            ".csv",
            ".xlsx",
        ):
            convert_data_format(input_path, output_path)
        elif input_ext == ".svg" and output_ext in (".png", ".jpg", ".jpeg", ".tiff"):
            convert_svg(input_path, output_path)
        elif input_ext in (".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif"):
            convert_image(input_path, output_path)
        elif input_ext == ".pdf" and output_ext in (".png", ".jpg", ".jpeg", ".tiff"):
            convert_pdf_to_image(input_path, output_path)
        elif input_ext == ".xlsx" and output_ext == ".csv":
            excel_to_csv(input_path, output_path)
        elif input_ext == ".csv" and output_ext == ".xlsx":
            csv_to_excel(input_path, output_path)
        elif input_ext == ".csv" and output_ext == ".json":
            csv_to_json(input_path, output_path)
        elif input_ext == ".json" and output_ext == ".csv":
            json_to_csv(input_path, output_path)
        elif input_ext == ".json" and output_ext == ".yaml":
            json_to_yaml(input_path, output_path)
        elif input_ext == ".yaml" and output_ext == ".json":
            yaml_to_json(input_path, output_path)
        elif input_ext == ".md" and output_ext == ".html":
            markdown_to_html(input_path, output_path)
        elif input_ext == ".md" and output_ext == ".pdf":
            markdown_to_pdf(input_path, output_path)
        elif input_ext == ".html" and output_ext == ".md":
            html_to_markdown(input_path, output_path)
        elif input_ext == ".epub" and output_ext == ".pdf":
            epub_to_pdf(input_path, output_path)
        elif output_ext == ".zip":
            compress_zip(input_path, output_path)
        elif input_ext == ".zip":
            extract_zip(input_path, output_path)
        elif output_ext == ".rar":
            compress_rar(input_path, output_path)
        elif input_ext == ".rar":
            extract_rar(input_path, output_path)
        elif output_ext == ".7z":
            compress_7z(input_path, output_path)
        elif input_ext == ".7z":
            extract_7z(input_path, output_path)
        elif input_ext == ".pdf" and output_ext == ".docx":
            pdf_to_word(input_path, output_path)
        elif input_ext == ".html" and output_ext == ".pdf":
            html_to_pdf(input_path, output_path)
        elif input_ext in (".mp4", ".avi", ".mov", ".mkv") and output_ext in (".mp3", ".wav"):
            extract_audio_from_video(input_path, output_path)
        else:
            raise ValueError(f"Unsupported conversion: {input_ext} to {output_ext}")
        
        logging.info(f"Conversion completed successfully: {output_path}")
    except Exception as e:
        logging.error(f"Conversion failed: {str(e)}")
        raise


def get_supported_conversions(input_ext):
    supported = {
        "png": [".jpg", ".jpeg", ".tiff", ".pdf", ".svg", ".bmp", ".gif"],
        "jpg": [".png", ".tiff", ".pdf", ".svg", ".bmp", ".gif"],
        "jpeg": [".png", ".tiff", ".pdf", ".svg", ".bmp", ".gif"],
        "tiff": [".png", ".jpg", ".jpeg", ".pdf", ".svg", ".bmp", ".gif"],
        "heic": [".png", ".jpg", ".jpeg", ".tiff", ".pdf"],
        "webp": [".png", ".jpg", ".jpeg", ".tiff", ".pdf"],
        "gif": [".png", ".jpg", ".jpeg", ".tiff", ".pdf", ".bmp"],
        "bmp": [".png", ".jpg", ".jpeg", ".tiff", ".pdf", ".gif"],
        "pdf": [".png", ".jpg", ".jpeg", ".tiff", ".docx"],
        "docx": [".pdf"],
        "json": [".yaml", ".csv", ".xlsx"],
        "yaml": [".json", ".csv", ".xlsx"],
        "md": [".html", ".pdf"],
        "html": [".md", ".pdf"],
        "epub": [".pdf"],
        "zip": ["."],
        "rar": ["."],
        "7z": ["."],
        "mp4": [".mov", ".avi", ".mkv", ".webm", ".flv", ".wmv", ".mp3", ".wav"],
        "mov": [".mp4", ".avi", ".mkv", ".webm", ".flv", ".wmv", ".mp3", ".wav"],
        "avi": [".mp4", ".mov", ".mkv", ".webm", ".flv", ".wmv", ".mp3", ".wav"],
        "mkv": [".mp4", ".mov", ".avi", ".webm", ".flv", ".wmv", ".mp3", ".wav"],
    }

    if ffmpeg_available:
        supported.update(
            {
                "mp4": [".mov", ".avi", ".mkv", ".webm", ".flv", ".wmv"],
                "mov": [".mp4", ".avi", ".mkv", ".webm", ".flv", ".wmv"],
                "avi": [".mp4", ".mov", ".mkv", ".webm", ".flv", ".wmv"],
                "mkv": [".mp4", ".mov", ".avi", ".webm", ".flv", ".wmv"],
                "webm": [".mp4", ".mov", ".avi", ".mkv", ".flv", ".wmv"],
                "flv": [".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv"],
                "wmv": [".mp4", ".mov", ".avi", ".mkv", ".webm", ".flv"],
                "mp3": [".wav", ".ogg", ".flac", ".aac", ".m4a", ".wma"],
                "wav": [".mp3", ".ogg", ".flac", ".aac", ".m4a", ".wma"],
                "ogg": [".mp3", ".wav", ".flac", ".aac", ".m4a", ".wma"],
                "flac": [".mp3", ".wav", ".ogg", ".aac", ".m4a", ".wma"],
                "aac": [".mp3", ".wav", ".ogg", ".flac", ".m4a", ".wma"],
                "m4a": [".mp3", ".wav", ".ogg", ".flac", ".aac", ".wma"],
                "wma": [".mp3", ".wav", ".ogg", ".flac", ".aac", ".m4a"],
            }
        )

    return supported.get(input_ext.lower().lstrip("."), [])

def get_supported_input_formats() -> Dict[str, List[str]]:
    return {
        "Document": [".docx", ".pdf", ".txt", ".md", ".html", ".epub"],
        "Image": [".png", ".jpg", ".jpeg", ".tiff", ".heic", ".webp", ".gif", ".bmp"],
        "Spreadsheet": [".xlsx", ".csv"],
        "Data": [".json", ".yaml", ".xml"],
        "Archive": [".zip", ".rar", ".7z"],
        "Video": [".mp4", ".mov", ".avi", ".mkv", ".webm", ".flv", ".wmv"],
        "Audio": [".mp3", ".wav", ".ogg", ".flac", ".aac", ".m4a", ".wma"]
    }

def prompt_for_files():
    while True:
        user_input = input("Enter input file(s) (or h for help, q to quit): ").strip()
        
        if user_input.lower() == 'q':
            return None, None
        
        if user_input.lower() == 'h':
            show_help()
            continue
        
        input_files = user_input.split()
        
        if not input_files:
            print("Please enter at least one input file.")
            continue
        
        all_supported_formats = [fmt for formats in get_supported_input_formats().values() for fmt in formats]
        if not all(os.path.splitext(f)[1].lower() in all_supported_formats for f in input_files):
            print("One or more input files have unsupported formats. Use h to see supported formats.")
            continue
        
        output_file = input("Enter output file: ").strip()
        
        if not output_file:
            print("Please enter an output file.")
            continue
        
        return input_files, output_file

def show_help():
    supported_formats = get_supported_input_formats()
    print("\nSupported input formats:")
    for category, formats in supported_formats.items():
        print(f"\n{category}:")
        print(", ".join(formats))
    
    print("\nMultiple file conversion:")
    print("To convert multiple files to PDF, enter the input files separated by spaces.")
    print("Example: file1.jpg file2.png file3.tiff")
    print("Note: Multiple file conversion is only supported for PDF output.")
    
    while True:
        file_type = input("\nEnter a file extension to see possible conversions (or 'q' to quit help): ").strip().lower()
        
        if file_type == 'q':
            break
        
        if not file_type.startswith('.'):
            file_type = '.' + file_type
        
        format_found = False
        for formats in supported_formats.values():
            if file_type in formats:
                format_found = True
                possible_conversions = get_supported_conversions(file_type)
                if possible_conversions:
                    print(f"Possible conversions for {file_type}:")
                    print(", ".join(possible_conversions))
                else:
                    print(f"No conversions available for {file_type}")
                break
        
        if not format_found:
            print(f"{file_type} is not a supported input format.")

def main():
    while True:
        input_files, output_file = prompt_for_files()
        
        if input_files is None and output_file is None:
            print("Exiting the program.")
            break
        
        if len(input_files) > 1 and not output_file.lower().endswith('.pdf'):
            print("Error: Multiple input files are only supported for PDF output.")
            continue

        try:
            convert_file(input_files if len(input_files) > 1 else input_files[0], output_file)
            print(f"Conversion successful: {output_file}")
        except Exception as e:
            print(f"Conversion failed: {str(e)}")

if __name__ == "__main__":
    get_project_requirements()
    main()