import xlwings as xw
from xlwings.constants import SortOrder
import os

class ArquivoExcel:
    def __init__(self,file_path = None, visibility = True, filtered = False):
        if file_path:
            self.file_path = file_path
            self.visibilidade = visibility
            self.filtrada = filtered
            self.book = xw.Book(self.file_path)
        else:
            self.book = xw.Book()
            self.file_path = None
            self.visibility = visibility
            self.filtered = filtered