import ArquivoExcel
import os

class TableDivisorSimple(ArquivoExcel.ArquivoExcel):
    def __init__(self,file_path = None, visibility = True, filtered = False):
        super().__init__(file_path=file_path, visibility=visibility, filtered=filtered)

    def spreadsheet_references_save(self,file_divisor):

        self.arquive_name = os.path.basename(file_divisor)

        self.wk_book_division = ArquivoExcel.ArquivoExcel(file_divisor)
        self.tab_division = self.wk_book_division.select_first_tab()

    def division_spreadsheet(self,number_of_division):

        self.total_rows = self.wk_book_division.extract_last_filled_row(self.tab_division,2)
        self.division_rows = self.total_rows//number_of_division
            
    def create_new_files_and_paste(self,number_of_division,book_path):
        header = self.wk_book_division.obtain_final_header(self.tab_division)
        for i in range(1,number_of_division):
            self.create_file_and_paste_division(f"{self.arquive_name}({i}).xlsx",book_path,self.tab_division,f"A2:{header}{self.division_rows}",header)
            self.delete_rows(self.tab_division,f"A2:{header}{self.division_rows}")
                