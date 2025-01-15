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
    
        #manipulação de linhas

    def extract_last_filled_row(self,spreadsheet_tab,column_sheet):
        return spreadsheet_tab.cells(1,column_sheet).end('down').row
    
    def create_row(self,spreadsheet_tab,row_position):
        spreadsheet_tab.cells(row_position,1).api.EntireRow.Insert()

    def delete_content_rows(self,spreadsheet_tab,selection_range):
        delete_cells = spreadsheet_tab.range(selection_range)
        delete_cells.clear_contents()

    def delete_rows(self,spreadsheet_tab,selection_range):
        delete_cells = spreadsheet_tab.range(selection_range)
        delete_cells.api.EntireRow.Delete()

    def delete_filtered_rows(self,spreadsheet_tab,selection_range):
        delete_cells = spreadsheet_tab.range(selection_range)
        delete_cells = delete_cells.api.SpecialCells(12)
        for cell in delete_cells:
            cell.EntireRow.Delete()
        
    def delete_rows_from_condition(self,complete_list,spreadsheet_tab,column_sheet):
        for cell in range(self.extract_last_filled_row(spreadsheet_tab,column_sheet) + 1, 2, -1):
            for value in complete_list:
                cell_value = value.value
                if cell_value == spreadsheet_tab.range(f"G{cell}").value:
                    spreadsheet_tab.range(f"G{cell}").api.EntireRow.Delete()
                    break

    #manipulação de colunas

    def create_column(self,spreadsheet_tab,column_position):
        spreadsheet_tab.cells(1,column_position).api.EntireColumn.Insert()
    
    def name_header(self,spreadsheet_tab,column_position,name_header):
        spreadsheet_tab.cells(1,column_position).value = name_header

    def delete_column(self,spreadsheet_tab,column_position):
        spreadsheet_tab.cells(1,column_position).api.EntireColumn.Delete()

    #uso de formulas 

    def text_join(self,delimiter,array,spreadsheet_tab,cell):
        formula_apply = f'=TEXTJOIN({delimiter},,{array})'
        spreadsheet_tab.range(cell).formula = formula_apply

    def text_join_msp(self,tab_name,cell):
        formula_apply = f'=TEXTJOIN(",",TRUE,UNIQUE(FILTER({tab_name}!F:F,({tab_name}!C:C=C{cell})*({tab_name}!L:L=G{cell}))))'
        return formula_apply
    
    def concat_campus_code(self,spreadsheet_tab,first_cell,second_cell,apply_range):
        formula_apply = f'=CONCAT({first_cell},";campus_code:",{second_cell})'
        spreadsheet_tab.range(apply_range).formula = formula_apply
        self.convert_to_value(apply_range,spreadsheet_tab)


    def concat(self,spreadsheet_tab,first_cell,second_cell,apply_range):
        formula_apply = f'=CONCAT({first_cell},"-",{second_cell})'
        spreadsheet_tab.range(apply_range).formula = formula_apply
        self.convert_to_value(apply_range,spreadsheet_tab)

    def convert_to_value(self,conversion_range,spreadsheet_tab):
        range = spreadsheet_tab.range(conversion_range)
        values = range.value
        range.value = [[val] for val in values]
    
    def xlook_up(self,search_value,search_array,return_array,spreadsheet_tab,formula_range):
        formula_apply = f"=XLOOKUP({search_value},{search_array},{return_array})"
        spreadsheet_tab.range(formula_range).formula = formula_apply
    
