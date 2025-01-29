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
    
    def separate_campus_and_apply_xlookup(self):
        self.wk_book_campus.create_tab("Positivo e Unipê")
        self.tab_posi_unip = self.wk_book_campus.select_tab("Positivo e Unipê")
        self.wk_book_campus.filter_apply(self.tab_campus,8,"UNIPÊ - GRADUAÇÃO EAD")
        self.wk_book_campus.copy_and_paste(self.tab_campus,self.tab_posi_unip,f"A2:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}","A1")
        self.wk_book_campus.delete_filtered_rows(self.tab_campus,f"A2:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}")
        self.wk_book_campus.filter_remove(self.tab_campus,8)
        self.wk_book_campus.filter_apply(self.tab_campus,8,"Universidade Positivo")
        self.wk_book_campus.copy_and_paste(self.tab_campus,self.tab_posi_unip,f"A2:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}",f"A{self.wk_book_campus.extract_last_filled_row(self.tab_posi_unip,1) + 1}")
        self.wk_book_campus.delete_filtered_rows(self.tab_campus,f"A2:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}")
        self.wk_book_campus.filter_remove(self.tab_campus,8)

        self.wk_book_relation.name_header(self.tab_relat_uni,3,"xlooup")
        self.wk_book_relation.name_header(self.tab_relat_uni,4,"concat")
        self.wk_book_relation.xlook_up("A2",f"'[{self.campus_name}]Positivo e Unipê'!$G:$G",f"'[{self.campus_name}]Positivo e Unipê'!$A:$A",self.tab_relat_uni,f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}")

        self.wk_book_relation.name_header(self.tab_relat_pos,3,"xlooup")
        self.wk_book_relation.name_header(self.tab_relat_pos,4,"concat")
        self.wk_book_relation.xlook_up("A2",f"'[{self.campus_name}]Positivo e Unipê'!$G:$G",f"'[{self.campus_name}]Positivo e Unipê'!$A:$A",self.tab_relat_pos,f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}")


    def check_NAs_and_treat(self):
        self.wk_book_relation.filter_apply(self.tab_relat_uni,3,"#N/A")
        if self.wk_book_relation.verify_filtered(f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}",self.tab_relat_uni):
            self.wk_book_relation.create_tab("Pendência Campus")
            self.tab_pending_campus = self.wk_book_relation.select_tab("Pendência Campus")
            self.tab_relat_uni.activate()
            self.wk_book_relation.copy_and_paste(self.tab_relat_uni,self.tab_pending_campus,f"A1:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}","A1")
            self.wk_book_relation.delete_filtered_rows(self.tab_relat_uni,f"A2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}")
        self.wk_book_relation.filter_remove(self.tab_relat_uni,3)
        self.wk_book_relation.convert_to_value(f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}",self.tab_relat_uni)
        self.wk_book_relation.concat_campus_code(self.tab_relat_uni,"C2","A2",f"D2:D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}")
        self.wk_book_relation.text_join(",",f"D2:D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}",self.tab_relat_uni,f"D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1) + 1}")
        self.wk_book_relation.fill_with_value(self.tab_relat_uni,f'D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,4)}',self.tab_relat_uni.range(f'D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,4)}').value)
        self.wk_book_relation.filter_apply(self.tab_relat_pos,3,"#N/A")
        if self.wk_book_relation.verify_filtered(self.tab_relat_pos,f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}"):
            self.tab_relat_pos.activate()
            self.wk_book_relation.copy_and_paste(self.tab_relat_pos,self.tab_pending_campus,f"A1:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}","A1")
            self.wk_book_relation.delete_filtered_rows(self.tab_relat_pos,f"A2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}")
        self.wk_book_relation.filter_remove(self.tab_relat_pos,3)
        if self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1) > 2:
            self.wk_book_relation.convert_to_value(f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}",self.tab_relat_pos)
            self.wk_book_relation.concat_campus_code(self.tab_relat_pos,"C2","A2",f"D2:D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}")
            self.wk_book_relation.text_join(",",f"D2:D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}",self.tab_relat_pos,f"D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1) + 1}")
            self.wk_book_relation.fill_with_value(self.tab_relat_pos,f'D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,4)}',self.tab_relat_pos.range(f'D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,4)}').value)
        else:
            self.wk_book_relation.concat_campus_code_unique_cell(self.tab_relat_pos,"C2","A2","D2")

