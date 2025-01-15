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

    
        #manipulação direta do arquivo

    def create_new_file(self, directory, file_name):
        file_path = os.path.join(directory,file_name)
        self.save_file(file_path)

    def close_file(self):
        self.book.close()
    
    def save_file(self,file_path):
        self.book.save(file_path)
    
    def close_file_without_saving(self):
        self.book.saved = True
        self.close_file()
    