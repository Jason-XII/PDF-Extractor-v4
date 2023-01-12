import os
from typing import List, Tuple

import fitz
from PyPDF4.merger import PdfFileReader, PdfFileWriter

import pdf_redactor


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
                images = page.get_images()
                for image in images:
                    self.count += 1
                    pix = fitz.Pixmap(doc, image[0])
                    if pix.n >= 5:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    with open(os.path.join(
                        self.output_dir, str(self.count)) + '.png', 'wb') as f:
                        f.write(pix.tobytes())


class PDFDeleteMachine:
    def __init__(self, filenames: List[str]):
        self.filenames = filenames

    def delete(self, pages: list, output_dir: str):
        for filename in self.filenames:
            output_filename = os.path.join(output_dir, os.path.split(filename)[-1]) if len(self.filenames) > 1 else output_dir
            self.extractor = PDFExtractMachine([(filename, 1, pages[0]-1),
                                                (filename, pages[1]+1, 'max')])
            self.extractor.extract_all(output_filename)


class PDFRotateMachine:
    def __init__(self, filename: str):
        self.filename = filename
        self.writer = PdfFileWriter()

    def rotate_clockwise(self, start, end, angle, output_filename: str):
        start -= 1
        reader = PdfFileReader(open(self.filename, 'rb'))
        pages = []
        for page in range(start, end):
            print(page)
            rotated_page = reader.getPage(page)
            if angle > 0:
                rotated_page.rotateClockwise(angle)
            else:
                rotated_page.rotateCounterClockwise(abs(angle))
            pages.append(rotated_page)
        for i in range(start):
            print('i', i)
            self.writer.addPage(reader.getPage(i))
        for k in pages:
            self.writer.addPage(k)
        for j in range(end, reader.numPages-1):
            print('j', j)
            self.writer.addPage(reader.getPage(j))
        self.writer.write(open(output_filename, 'wb'))


class PDFReplaceTextMachine:
    def __init__(self, input_filename: str):
        self.input = input_filename
        self.options = pdf_redactor.RedactorOptions()

    def replace_pdf(self, replacement: list, output_filename: str):
        input_s = open(self.input, 'rb')
        self.options.input_stream = input_s
        output_s = open(output_filename, 'wb')
        self.options.output_stream = output_s
        self.options.content_filters = replacement
        pdf_redactor.redactor(self.options)
        input_s.close()
        output_s.close()


class PDFRemoveImageMachine:
    def __init__(self, input_filename: str):
        self.input = input_filename

    def find_possible_watermarks(self):
        document = fitz.open(self.input)
        image_dict = {}
        possible_watermarks = []
        for each_page in document:
            print(dir(each_page))
            image_list = each_page.get_images()
            for info in image_list:
                print(info)
                pix = fitz.Pixmap(document, info[0])
                png = pix.tobytes()  # return picture in png format
                image_dict.setdefault(png, 0)
                image_dict[png] += 1
        for (image, count) in image_dict.items():
            if count >= 2:  # 10页PDF，如果出现4张以上则怀疑是水印
                possible_watermarks.append(image)
        return possible_watermarks

    def remove_image(self, watermark_image: bytes, out_filename: str) -> None:
        document = fitz.open(self.input)
        for each_page in document:
            image_list = each_page.get_images()
            for image_info in image_list:
                pix = fitz.Pixmap(document, image_info[0])
                png = pix.tobytes()  # return picture in png format
                if png == watermark_image:
                    document._deleteObject(image_info[0])
        document.save(out_filename)