import xlwings as xw
from xlwings.constants import SortOrder
import os
import math

class ArquivoExcel:
    def __init__(self,file_path = None, visibility = True, filtered = False):
        if file_path:
            self.file_path = file_path
            self.visibility = visibility
            self.filtered = filtered
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
        self.book.app.quit()

    def create_tab(self,tab_name):
        self.book.sheets.add(tab_name)

    def select_tab(self,tab_name):
        return self.book.sheets[f'{tab_name}']
    
    def delete_tab(self,tab_name):
        tab_to_delete = self.book.sheets[tab_name]
        tab_to_delete.delete()
    
    def move_sheet_position(self,sheet_move):
        sheet_move.api.Move(Before = self.book.sheets[0].api)
    
    def rename_tab(self,current_name,new_name):
        tab = self.select_tab(current_name)
        tab.name = new_name

    def create_file_and_paste_content(self,path_name,book_path,spreadsheet_copy,range_copy,second_copy_spreadsheet,second_copy_range):
        directory = os.path.dirname(book_path)
        self.wk_book_temporary = ArquivoExcel()
        self.wk_book_temporary.create_new_file(directory,path_name)
        path_save = os.path.join(directory,path_name)
        path_name = path_name.replace(".xlsx","")
        self.wk_book_temporary.rename_tab("Sheet1",f"{path_name}")
        self.tab_temporary =  self.wk_book_temporary.select_tab(f"{path_name}")
        self.wk_book_temporary.copy_and_paste(spreadsheet_copy,self.tab_temporary,range_copy,"A1")
        self.wk_book_temporary.copy_and_paste(second_copy_spreadsheet,self.tab_temporary,second_copy_range,f"E2:E{self.wk_book_temporary.extract_last_filled_row(self.tab_temporary,2)}")
        self.wk_book_temporary.save_file(path_save)
        self.wk_book_temporary.close_file()

        #manipulação de linhas

    def extract_last_filled_row(self,spreadsheet_tab,column_sheet):
        return spreadsheet_tab.cells(1,column_sheet).end('down').row
    
    def extract_end_row_up(self,spreadsheet_tab,last_row):
        return spreadsheet_tab.cells(last_row,2).end('up').row
    
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
        delete_cells = self.select_filtered(spreadsheet_tab,selection_range)
        for cell in delete_cells:
            cell.EntireRow.Delete()

    def delete_row(self,spreadsheet_tab,row,column):
        spreadsheet_tab.cells(row,column).api.EntireRow.Delete()

    def delete_rows_from_condition(self,complete_list,spreadsheet_tab,last_cell):
        for cell in range(last_cell + 1, 0, -1):
            for value in complete_list:
                cell_value = value.value
                if cell_value == spreadsheet_tab.range(f"G{cell}").value:
                    spreadsheet_tab.range(f"G{cell}").api.EntireRow.Delete()
                    break

    def last_row_from(self,spreadsheet_tab,column,row):
        return spreadsheet_tab.cells(row,column).end('down').row

    #manipulação de colunas

    def create_column(self,spreadsheet_tab,column_position):
        spreadsheet_tab.cells(1,column_position).api.EntireColumn.Insert()
    
    def name_header(self,spreadsheet_tab,column_position,name_header):
        spreadsheet_tab.cells(1,column_position).value = name_header

    def delete_column(self,spreadsheet_tab,column_position):
        spreadsheet_tab.cells(1,column_position).api.EntireColumn.Delete()

    def verify_position_value_header(self,spreadsheet_tab,value_search):
        header = spreadsheet_tab.range("A1").expand("right").value
        if value_search in header:
            position_header = header.index(value_search) + 1
        else:
            position_header = -1
        position_header = xw.utils.col_name(position_header)
        return position_header
    

    #uso de formulas 

    def formula_apply(self,spreadsheet_tab,cells,formula):
        spreadsheet_tab.range(cells).formula = formula

    def text_join(self,delimiter,array,spreadsheet_tab,cells):
        formula_apply = f'=TEXTJOIN({delimiter},,{array})'
        self.formula_apply(spreadsheet_tab,cells,formula_apply)

    def text_join_msp(self,tab_name,cell):
        formula_apply = f'=TEXTJOIN(",",TRUE,UNIQUE(FILTER({tab_name}!F:F,({tab_name}!C:C=C{cell})*({tab_name}!L:L=G{cell}))))'
        return formula_apply
    
    def concat_campus_code(self,spreadsheet_tab,first_cell,second_cell,apply_range):
        formula_apply = f'=CONCAT({first_cell},";campus_code:",{second_cell})'
        self.formula_apply(spreadsheet_tab,apply_range,formula_apply)
        self.convert_to_value(apply_range,spreadsheet_tab)

    def concat_campus_code_unique_cell(self,spreadsheet_tab,first_cell,second_cell,apply_range):
        formula_apply = f'=CONCAT({first_cell},";campus_code:",{second_cell})'
        self.formula_apply(spreadsheet_tab,apply_range,formula_apply)
        self.fill_with_value(spreadsheet_tab,apply_range,spreadsheet_tab.range(apply_range).value)

    def concat(self,spreadsheet_tab,first_cell,second_cell,apply_range):
        formula_apply = f'=CONCAT({first_cell},"-",{second_cell})'
        self.formula_apply(spreadsheet_tab,apply_range,formula_apply)
        self.convert_to_value(apply_range,spreadsheet_tab)

    def convert_to_value(self,conversion_range,spreadsheet_tab):
        range = spreadsheet_tab.range(conversion_range)
        values = range.value
        range.value = [[val] for val in values]
    
    def xlook_up(self,search_value,search_array,return_array,spreadsheet_tab,apply_range):
        formula_apply = f"=XLOOKUP({search_value},{search_array},{return_array})"
        self.formula_apply(spreadsheet_tab,apply_range,formula_apply)
    

    #uso de filtros

    def filter_apply(self,spreadsheet_tab,filter_column,filter):
        spreadsheet_tab.range('A1').api.AutoFilter(filter_column,filter)
        self.filtered = True

    def filter_remove(self,spreadsheet_tab,filter_column):
        if self.filtered == True:
            spreadsheet_tab.range('A1').api.AutoFilter(filter_column)
            self.filtered = False

    def sort_table(self,spreadsheet_tab,complete_range,column):
        spreadsheet_tab.range(complete_range).api.Sort(Key1 = spreadsheet_tab.range(column).api,
                                                   Order1 =SortOrder.xlAscending,
                                                   Header = 1,
                                                   Orientation = 1)
    
    def clear_only_filtered(self,spreadsheet_tab,selection_range):
        delete_cells = spreadsheet_tab.range(selection_range)
        delete_cells = self.select_filtered(spreadsheet_tab,selection_range)
        delete_cells.ClearContents()
    
    def delete_only_filtered(self,spreadsheet_tab,selection_range):
        xw.apps.active.api.DisplayAlerts = False
        delete_cells = self.select_filtered(spreadsheet_tab,selection_range)
        delete_cells.EntireRow.Delete()

    def select_filtered(self,spreadsheet_tab,selection_range):
        return spreadsheet_tab.range(selection_range).api.SpecialCells(12)
    
    def verify_filtered(self,search_range,spreadsheet_tab):
        try:
            visibles = self.select_filtered(spreadsheet_tab,search_range)
            if visibles.Count > 0:
                return True
        except Exception as e:
            if "com_error" in str(e):
                return False
            else:
                return False

        

     #outros métodos

    def replace(self,spreadsheet_tab,replace_range,original_value,new_value):
        to_replace = spreadsheet_tab.range(replace_range)
        to_replace.api.Replace(original_value,new_value)

    def remove_duplicates(self,spreadsheet_tab,apply_range,duplicate_column):
        spreadsheet_tab.range(apply_range).api.RemoveDuplicates(Columns = [duplicate_column],
                                                                Header = 0)
                        
    def copy_and_paste(self,copy_tab,paste_tab,copy_range,paste_range):
        copy_tab.range(copy_range).copy()
        paste_tab.range(paste_range).paste()

    def check_name_existence(self,name_from_verify,column,spreadsheet_tab):
        for row in range(2,self.extract_last_filled_row(spreadsheet_tab,column) + 1):
            if self.check_name(name_from_verify,column,spreadsheet_tab,row):
                return True
        return False
    
    def check_name(self,name_from_verify,column,spreadsheet_tab,row):
        if name_from_verify in str(spreadsheet_tab.cells(row,column).value).strip():
                return True
        return False

    def turn_into_text(self,spreadsheet_tab,column,column_cell):
        conversion = spreadsheet_tab.range(f'{column}2:{column}{self.extract_last_filled_row(spreadsheet_tab,column_cell)}')
        conversion.api.TextToColumns(Destination = conversion.api,
        DataType = 1,
        Semicolon = False )

    def fill_with_value(self,spreadsheet_tab,fill_range,value):
        spreadsheet_tab.range(fill_range).value = value

    def check_that_it_does_not_exceed_the_limit(self,total_rows,divide_for,spreadsheet_tab):
        overtake = False
        total_loops = total_rows/divide_for
        total_loops = math.ceil(total_loops)

        for i in range(1,total_loops + 1):
            if (i == 1):
                values = spreadsheet_tab.range(f"Z2:Z{divide_for}").value
                separator = ","
                simul_textjoin = separator.join(values)
                if(len(simul_textjoin) > 32767):
                    overtake = True
                    break
            elif i == total_loops - 1:
                values = spreadsheet_tab.range(f"Z{divide_for * i - 1}:Z{self.extract_last_filled_row(spreadsheet_tab,1)}").value
                values = list(filter(None,values))
                separator = ","
                simul_textjoin = separator.join(values)
                if(len(simul_textjoin) > 32767):
                    overtake = True
                    break
                values = spreadsheet_tab.range(f"Z{divide_for * i - 1}:Z{divide_for * i}").value
                values = list(filter(None,values))
                separator = ","
                simul_textjoin = separator.join(values)
                if(len(simul_textjoin) > 32767):
                    overtake = True
                    break
            else:
                values = spreadsheet_tab.range(f"Z{divide_for * (i - 1) + 1}:Z{divide_for * i}").value
                values = list(filter(None,values))
                separator = ","
                simul_textjoin = separator.join(values)
                if(len(simul_textjoin) > 32767):
                    overtake = True
                    break
        return overtake
    
    def number_usable_in_for(self,total_rows,divide_for):
        total_loops = total_rows/divide_for
        total_loops = math.ceil(total_loops) + 1
        return total_loops