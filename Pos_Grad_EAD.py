import ArquivoExcel
import os
import CompletionMessage
class PosGraduacaoEAD(ArquivoExcel.ArquivoExcel):
    def __init__(self,file_path = None, visibility = True, filtered = False):
        super().__init__(file_path=file_path, visibility=visibility, filtered=filtered)
    
    def spreadsheet_processing(self,file_msp,file_campus,file_relation):

        self.campus_name = os.path.basename(file_campus)

        self.wk_book_msp = ArquivoExcel.ArquivoExcel(file_msp)
        self.tab_msp = self.wk_book_msp.select_tab("Modelo Sem Parar")

        self.wk_book_campus = ArquivoExcel.ArquivoExcel(file_campus)
        self.tab_campus = self.wk_book_campus.select_tab("Sheet 1")

        self.wk_book_relation = ArquivoExcel.ArquivoExcel(file_relation)
        self.tab_relat_uni = self.wk_book_relation.select_tab("UNIPÊ")
        self.tab_relat_pos = self.wk_book_relation.select_tab("POSITIVO")
    
    def check_and_treat_metadata(self):
        self.wk_book_campus.filter_apply(self.tab_campus,7,"=")
        if self.wk_book_campus.verify_filtered(f"G2:G{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}",self.tab_campus):
            self.wk_book_campus.create_tab("Polos com Pendência")
            self.tab_pending = self.wk_book_campus.select_tab("Polos com Pendência")
            self.wk_book_campus.copy_and_paste(self.tab_campus,self.tab_pending,f"A1:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}","A1")
            self.wk_book_campus.delete_filtered_rows(self.tab_campus,f"A2:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}")
        self.wk_book_campus.filter_remove(self.tab_campus,7)
        self.wk_book_campus.turn_into_text(self.tab_campus,"G",1)