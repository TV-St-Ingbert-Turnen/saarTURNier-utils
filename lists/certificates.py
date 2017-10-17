from .list_generator import ListGenerator
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger


class CertificateGenerator(ListGenerator):
    def generate(self, team_list):
        self._template = PdfFileReader(open('..templates/certificates/sample.pdf'))
        merger = PdfFileMerger()
        # read config file

        for team in team_list:

            # create new page from template
            # insert team info
            pass


        pass

    def close(self):
        pass

    def __init__(self, path):
        super().__init__(path)

    def write(self):
        pass
