import unittest
import os
import tempfile
import shutil
from reportlab.pdfgen import canvas

from fileconvert import (
    excel_to_csv,
    csv_to_excel,
    csv_to_json,
    json_to_csv,
    json_to_yaml,
    yaml_to_json,
    word_to_pdf,
    pdf_to_text,
    text_to_word,
    convert_image,
    convert_pdf_to_image,
    markdown_to_html,
    html_to_markdown,
    compress_zip,
    extract_zip,
    images_to_pdf,
    html_to_pdf,
)

class TestFileConvert(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        for file in os.listdir(self.temp_dir):
            file_path = os.path.join(self.temp_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                import shutil
                shutil.rmtree(file_path)
        os.rmdir(self.temp_dir)

    def test_excel_to_csv(self):
        input_file = os.path.join(self.temp_dir, "test.xlsx")
        output_file = os.path.join(self.temp_dir, "test.csv")
        import pandas as pd
        df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        df.to_excel(input_file, index=False)
        
        excel_to_csv(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_csv_to_excel(self):
        input_file = os.path.join(self.temp_dir, "test.csv")
        output_file = os.path.join(self.temp_dir, "test.xlsx")
        with open(input_file, 'w') as f:
            f.write("A,B\n1,4\n2,5\n3,6")
        
        csv_to_excel(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_csv_to_json(self):
        input_file = os.path.join(self.temp_dir, "test.csv")
        output_file = os.path.join(self.temp_dir, "test.json")
        with open(input_file, 'w') as f:
            f.write("A,B\n1,4\n2,5\n3,6")
        
        csv_to_json(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_json_to_csv(self):
        input_file = os.path.join(self.temp_dir, "test.json")
        output_file = os.path.join(self.temp_dir, "test.csv")
        import json
        data = [{"A": 1, "B": 4}, {"A": 2, "B": 5}, {"A": 3, "B": 6}]
        with open(input_file, 'w') as f:
            json.dump(data, f)
        
        json_to_csv(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_json_to_yaml(self):
        input_file = os.path.join(self.temp_dir, "test.json")
        output_file = os.path.join(self.temp_dir, "test.yaml")
        import json
        data = {"A": [1, 2, 3], "B": [4, 5, 6]}
        with open(input_file, 'w') as f:
            json.dump(data, f)
        
        json_to_yaml(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_yaml_to_json(self):
        input_file = os.path.join(self.temp_dir, "test.yaml")
        output_file = os.path.join(self.temp_dir, "test.json")
        with open(input_file, 'w') as f:
            f.write("A: [1, 2, 3]\nB: [4, 5, 6]")
        
        yaml_to_json(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_word_to_pdf(self):
        input_file = os.path.join(self.temp_dir, "test.docx")
        output_file = os.path.join(self.temp_dir, "test.pdf")
        from docx import Document
        doc = Document()
        doc.add_paragraph("Test document")
        doc.save(input_file)
        
        word_to_pdf(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_pdf_to_text(self):
        input_file = os.path.join(self.temp_dir, "test.pdf")
        output_file = os.path.join(self.temp_dir, "test.txt")
        c = canvas.Canvas(input_file)
        c.drawString(100, 100, "Test PDF")
        c.save()
        
        pdf_to_text(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_text_to_word(self):
        input_file = os.path.join(self.temp_dir, "test.txt")
        output_file = os.path.join(self.temp_dir, "test.docx")
        with open(input_file, 'w') as f:
            f.write("Test text document")
        
        text_to_word(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_convert_image(self):
        input_file = os.path.join(self.temp_dir, "test.png")
        output_file = os.path.join(self.temp_dir, "test.jpg")
        from PIL import Image
        img = Image.new('RGB', (100, 100), color='red')
        img.save(input_file)
        
        convert_image(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_convert_pdf_to_image(self):
        input_file = os.path.join(self.temp_dir, "test.pdf")
        output_file = os.path.join(self.temp_dir, "test.png")
        c = canvas.Canvas(input_file)
        c.drawString(100, 100, "Test PDF")
        c.save()
        
        convert_pdf_to_image(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_markdown_to_html(self):
        input_file = os.path.join(self.temp_dir, "test.md")
        output_file = os.path.join(self.temp_dir, "test.html")
        with open(input_file, 'w') as f:
            f.write("# Test Markdown\n\nThis is a test.")
        
        markdown_to_html(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_html_to_markdown(self):
        if not shutil.which('pandoc'):
            self.skipTest("Pandoc is not installed. Skipping test.")
        input_file = os.path.join(self.temp_dir, "test.html")
        output_file = os.path.join(self.temp_dir, "test.md")
        with open(input_file, 'w') as f:
            f.write("<h1>Test HTML</h1><p>This is a test.</p>")
        
        html_to_markdown(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_compress_zip(self):
        input_file = os.path.join(self.temp_dir, "test.txt")
        output_file = os.path.join(self.temp_dir, "test.zip")
        with open(input_file, 'w') as f:
            f.write("Test file for compression")
        
        compress_zip(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_extract_zip(self):
        input_file = os.path.join(self.temp_dir, "test.zip")
        output_dir = os.path.join(self.temp_dir, "extracted")
        import zipfile
        with zipfile.ZipFile(input_file, 'w') as zipf:
            zipf.writestr("test.txt", "Test file for extraction")
        
        os.mkdir(output_dir)
        extract_zip(input_file, output_dir)
        self.assertTrue(os.path.exists(os.path.join(output_dir, "test.txt")))

    def test_images_to_pdf(self):
        input_files = [
            os.path.join(self.temp_dir, "test1.png"),
            os.path.join(self.temp_dir, "test2.png")
        ]
        output_file = os.path.join(self.temp_dir, "test.pdf")
        from PIL import Image
        for file in input_files:
            img = Image.new('RGB', (100, 100), color='red')
            img.save(file)
        
        images_to_pdf(input_files, output_file)
        self.assertTrue(os.path.exists(output_file))

    def test_html_to_pdf(self):
        if not shutil.which('wkhtmltopdf'):
            self.skipTest("wkhtmltopdf is not installed. Skipping test.")
        input_file = os.path.join(self.temp_dir, "test.html")
        output_file = os.path.join(self.temp_dir, "test.pdf")
        with open(input_file, 'w') as f:
            f.write("<h1>Test HTML</h1><p>This is a test.</p>")
        
        html_to_pdf(input_file, output_file)
        self.assertTrue(os.path.exists(output_file))

if __name__ == '__main__':
    unittest.main()