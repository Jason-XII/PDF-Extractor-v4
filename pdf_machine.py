from PyPDF4.merger import PdfFileReader, PdfFileWriter
from typing import List, Tuple
import fitz
import os


class PDFExtractMachine:
    def __init__(self, pdf_data: List[Tuple[str, int, int]]):
        self.data = pdf_data

    def extract_all(self, output_filename: str):
        writer = PdfFileWriter()
        for data in self.data:
            start, end = data[1], data[2]
            reader = PdfFileReader(open(data[0], 'rb'))
            end = reader.numPages if end == 'max' else end
            for page_num in range(int(start) - 1, int(end)):
                page = reader.getPage(page_num)
                writer.addPage(page)
        with open(output_filename, 'wb') as pdf:
            writer.write(pdf)

    def extract_one(self, in_filename: str, start: int, end: int, out_filename: str):
        writer = PdfFileWriter()
        reader = PdfFileReader(open(in_filename, 'rb'))
        for page_num in range(int(start) - 1, int(end)):
            page = reader.getPage(page_num)
            writer.addPage(page)
        with open(out_filename, 'wb') as pdf:
            writer.write(pdf)


class PDFMergeMachine:
    def __init__(self, pdf_filenames: List[str]):
        self.filenames = pdf_filenames

    def merge(self, output_filename: str):
        writer = PdfFileWriter()
        for filename in self.filenames:
            reader = PdfFileReader(open(filename, 'rb'))
            writer.appendPagesFromReader(reader)
        writer.write(open(output_filename, 'wb'))


class PDFExtractImageMachine:
    def __init__(self, input_pdf_list: str, output_dir: str):
        self.pdf_list = input_pdf_list
        self.output_dir = output_dir
        self.count = 0

    def extract(self):
        for pdf in self.pdf_list:
            doc = fitz.open(pdf)
            for page in doc:
                images = page.getImageList()
                for image in images:
                    self.count += 1
                    pix = fitz.Pixmap(doc, image[0])
                    if pix.n >= 5:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    pix.writePNG(os.path.join(
                        self.output_dir, str(self.count)) + '.png')


class PDFDeleteMachine:
    def __init__(self, filenames: List[str]):
        self.filenames = filenames

    def delete(self, pages: list, output_dir: str):
        for filename in self.filenames:
            output_filename = os.path.join(output_dir, os.path.split(filename)[-1]) if len(self.filenames) > 1 else output_dir
            self.extractor = PDFExtractMachine([(filename, 1, pages[0]-1),
                                                (filename, pages[1]+1, 'max')])
            self.extractor.extract_all(output_filename)