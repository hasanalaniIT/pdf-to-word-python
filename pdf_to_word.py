from pdf2docx import Converter


class PDFtoWordConverter:
    def convert_file(self, file_path, save_path):
        cv = Converter(file_path)
        cv.convert(save_path, start=0, end=None)
        cv.close()
