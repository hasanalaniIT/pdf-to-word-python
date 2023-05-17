from pdf2docx import Converter


class PDFtoWordConverter:
    """Responsible for converting PDF files to DOCX format."""

    @staticmethod
    def convert_file(file_path, save_path) -> None:
        """Converts pdf file to docx.

        param file_path: path of PDF file.
        param save_path: the directory where the docx file will be saved into.
        """
        cv = Converter(file_path)
        cv.convert(save_path, start=0, end=None)
        cv.close()
