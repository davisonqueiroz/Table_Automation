import ArquivoExcel
import os
import CompletionMessage
class PosGraduacaoEAD(ArquivoExcel.ArquivoExcel):
    def __init__(self,file_path = None, visibility = True, filtered = False):
        super().__init__(file_path=file_path, visibility=visibility, filtered=filtered)
    
    def spreadsheet_processing(self,file_msp,file_campus,file_relation,file_campus_uni_pos):

        self.campus_name = os.path.basename(file_campus_uni_pos)

        self.wk_book_msp = ArquivoExcel.ArquivoExcel(file_msp)
        self.tab_msp = self.wk_book_msp.select_tab("Modelo Sem Parar")

        self.wk_book_campus = ArquivoExcel.ArquivoExcel(file_campus)
        self.tab_campus = self.wk_book_campus.select_tab("Sheet 1")

        self.wk_book_campus_uni_pos = ArquivoExcel.ArquivoExcel(file_campus_uni_pos)
        self.tab_campus_uni_pos = self.wk_book_campus_uni_pos.select_tab("Sheet 1")

        self.wk_book_relation = ArquivoExcel.ArquivoExcel(file_relation)
        self.tab_relat_uni = self.wk_book_relation.select_tab("UNIPÊ")
        self.tab_relat_pos = self.wk_book_relation.select_tab("POSITIVO")
    
    def check_and_treat_metadata(self):
        self.wk_book_campus.filter_apply(self.tab_campus,7,"=")
        self.polo_pending = False
        if self.wk_book_campus.verify_filtered(f"G2:G{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}",self.tab_campus):
            self.wk_book_campus.create_tab("Polos com Pendência")
            self.tab_pending = self.wk_book_campus.select_tab("Polos com Pendência")
            self.wk_book_campus.copy_and_paste(self.tab_campus,self.tab_pending,f"A1:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}","A1")
            self.wk_book_campus.delete_filtered_rows(self.tab_campus,f"A2:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}")
            self.polo_pending = True
        self.wk_book_campus.filter_remove(self.tab_campus,7)
        self.wk_book_campus.turn_into_text(self.tab_campus,"G",1)
        self.wk_book_campus_uni_pos.turn_into_text(self.tab_campus_uni_pos,"G",1)
        self.total_rows_campus = self.wk_book_campus.extract_last_filled_row(self.tab_campus,2)
    
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
        self.wk_book_relation.xlook_up("A2",f"'[{self.campus_name}]Sheet 1'!$G:$G",f"'[{self.campus_name}]Sheet 1'!$A:$A",self.tab_relat_uni,f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}")

        self.wk_book_relation.name_header(self.tab_relat_pos,3,"xlooup")
        self.wk_book_relation.name_header(self.tab_relat_pos,4,"concat")
        self.wk_book_relation.xlook_up("A2",f"'[{self.campus_name}]Sheet 1'!$G:$G",f"'[{self.campus_name}]Sheet 1'!$A:$A",self.tab_relat_pos,f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}")
        self.wk_book_relation.fill_with_value(self.tab_relat_pos,f"C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,3)}",self.tab_relat_pos.range(f"C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,3)}").value)


    def check_NAs_and_treat(self):
        self.wk_book_relation.filter_apply(self.tab_relat_uni,3,"#N/A")
        self.campus_relat_pending = False
        if self.wk_book_relation.verify_filtered(f"C2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}",self.tab_relat_uni):
            self.wk_book_relation.create_tab("Pendência Campus")
            self.tab_pending_campus = self.wk_book_relation.select_tab("Pendência Campus")
            self.tab_relat_uni.activate()
            self.wk_book_relation.copy_and_paste(self.tab_relat_uni,self.tab_pending_campus,f"A1:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}","A1")
            self.wk_book_relation.delete_filtered_rows(self.tab_relat_uni,f"A2:C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}")
            self.campus_relat_pending = True
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

    def remove_campus_from_exp(self):
        relation_list = self.tab_relat_uni.range(f"A2:A{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,1)}")
        self.wk_book_campus.delete_rows_from_condition(relation_list,self.tab_posi_unip,self.wk_book_campus.extract_last_filled_row(self.tab_posi_unip,1))

        relation_list = self.tab_relat_pos.range(f"A2:A{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,1)}")
        self.wk_book_campus.delete_rows_from_condition(relation_list,self.tab_posi_unip,self.wk_book_campus.extract_last_filled_row(self.tab_posi_unip,1))
        self.wk_book_campus.copy_and_paste(self.tab_posi_unip,self.tab_campus,f"A1:Y{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}",f"A{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1) + 1}")
        self.wk_book_campus.delete_tab("Positivo e Unipê")

    def separate_rows_and_concatenate(self):
        self.row_campus_end = self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)
        self.number_division = 1365
        self.range_for = self.wk_book_campus.number_usable_in_for(self.wk_book_campus.extract_last_filled_row(self.tab_campus,1),self.number_division)
        self.wk_book_campus.concat_campus_code(self.tab_campus,"A2","G2",f"Z2:Z{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}")

        textjoin_validation = self.wk_book_campus.check_that_it_does_not_exceed_the_limit(self.wk_book_campus.extract_last_filled_row(self.tab_campus,1),self.number_division,self.tab_campus)
        if (textjoin_validation == False):
            self.row_campus_end = self.row_campus_end + (self.range_for - 2)
            for row in range(1,self.range_for):
                self.wk_book_campus.create_row(self.tab_campus,row * self.number_division)

                if row == 1:
                    self.wk_book_campus.text_join(",",f"Z2:Z{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}",self.tab_campus,f"AA{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1) + 1}")
                    self.fill_with_value(self.tab_campus,f"AA{self.number_division}",self.tab_campus.range(f"AA{self.number_division}").value)
                elif row == self.range_for - 1: 
                    self.wk_book_campus.text_join(",",f"Z{self.wk_book_campus.extract_end_row_up(self.tab_campus,self.row_campus_end)}:Z{self.row_campus_end}",self.tab_campus,f"AA{self.row_campus_end + 1}")
                    self.fill_with_value(self.tab_campus,f"AA{self.row_campus_end + 1}",self.tab_campus.range(f"AA{self.row_campus_end + 1}").value)
                else:
                    self.wk_book_campus.text_join(",",f"Z{((row - 1 )* self.number_division) + 1}:Z{(self.number_division * row) - 1}",self.tab_campus,f"AA{self.number_division * row}")
                    self.fill_with_value(self.tab_campus,f"AA{self.number_division * row}",self.tab_campus.range(f"AA{self.number_division * row}").value)
        else:
            while(textjoin_validation == True): 
                self.number_division = self.number_division - 5
                textjoin_validation = self.wk_book_campus.check_that_it_does_not_exceed_the_limit(self.wk_book_campus.extract_last_filled_row(self.tab_campus,1),self.number_division,self.tab_campus)

            self.range_for = self.wk_book_campus.number_usable_in_for(self.wk_book_campus.extract_last_filled_row(self.tab_campus,1),self.number_division)
            self.row_campus_end = self.row_campus_end + (self.range_for - 2)
            for row in range(1,self.range_for):
                self.wk_book_campus.create_row(self.tab_campus,row * self.number_division)
                if row == 1:
                    self.wk_book_campus.text_join(",",f"Z2:Z{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}",self.tab_campus,f"AA{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1) + 1}")
                    self.fill_with_value(self.tab_campus,f"AA{self.number_division}",self.tab_campus.range(f"AA{self.number_division}").value)
                elif row == self.range_for -1: 
                    self.wk_book_campus.text_join(",",f"Z{self.wk_book_campus.extract_end_row_up(self.tab_campus,self.row_campus_end)}:Z{self.row_campus_end}",self.tab_campus,f"AA{self.row_campus_end + 1}")
                    self.fill_with_value(self.tab_campus,f"AA{self.row_campus_end + 1}",self.tab_campus.range(f"AA{self.row_campus_end + 1}").value)
                else:
                    self.wk_book_campus.text_join(",",f"Z{((row - 1 )* self.number_division) + 1}:Z{(self.number_division * row) - 1}",self.tab_campus,f"AA{self.number_division * row}")
                    self.fill_with_value(self.tab_campus,f"AA{self.number_division * row}",self.tab_campus.range(f"AA{self.number_division * row}").value)
            

    def separate_universities(self):
        self.wk_book_msp.create_tab("UNIPE E POSITIVO")
        self.tab_uni_msp = self.wk_book_msp.select_tab("UNIPE E POSITIVO")
        self.wk_book_msp.filter_apply(self.tab_msp,2,"UNIPÊ - PÓS-GRADUAÇÃO EAD")
        self.wk_book_msp.copy_and_paste(self.tab_msp,self.tab_uni_msp,f"A1:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}","A1")
        self.wk_book_msp.delete_filtered_rows(self.tab_msp,f"A2:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}")
        self.wk_book_msp.copy_and_paste(self.tab_relat_uni,self.tab_uni_msp,f"D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_uni,4)}",f"E2:E{self.wk_book_msp.extract_last_filled_row(self.tab_uni_msp,2)}")
        self.filter_remove(self.tab_msp,2)
        self.wk_book_msp.filter_apply(self.tab_msp,2,"POSITIVO - PÓS-GRADUAÇÃO EAD")
        self.wk_book_msp.copy_and_paste(self.tab_msp,self.tab_uni_msp,f"A2:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}",f"A{self.wk_book_msp.extract_last_filled_row(self.tab_uni_msp,2) + 1}")
        self.wk_book_msp.delete_filtered_rows(self.tab_msp,f"A2:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}")
        col_pos = self.wk_book_msp.verify_position_value_header(self.tab_uni_msp,"COD CAMPUS")
        self.wk_book_msp.copy_and_paste(self.tab_relat_pos,self.tab_uni_msp,f"D{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,4)}",f"{col_pos}{self.wk_book_msp.extract_last_filled_row(self.tab_uni_msp,5) + 1}:AP{self.wk_book_msp.extract_last_filled_row(self.tab_uni_msp,2)}")
        self.wk_book_msp.copy_and_paste(self.tab_relat_pos,self.tab_uni_msp,f"C{self.wk_book_relation.extract_last_filled_row(self.tab_relat_pos,4)}",f"E{self.wk_book_msp.extract_last_filled_row(self.tab_uni_msp,5) + 1}:E{self.wk_book_msp.extract_last_filled_row(self.tab_uni_msp,2)}")

        self.filter_apply(self.tab_msp,2,"CRUZEIRO DO SUL - PÓS EAD")

    def create_copy_and_separate(self,book_path):
        new_path_name = "POSITIVO_UNIPÊ.xlsx"
        directory = os.path.dirname(book_path)
        self.path_save = os.path.join(directory,new_path_name)
        self.wk_book_unipe_positivo = ArquivoExcel.ArquivoExcel()
        self.wk_book_unipe_positivo.create_new_file(directory,new_path_name)
        self.wk_book_unipe_positivo.rename_tab("Sheet1","UNIPE E POSITIVO")
        self.tab_unipe = self.wk_book_unipe_positivo.select_tab("UNIPE E POSITIVO")
        self.wk_book_unipe_positivo.copy_and_paste(self.tab_uni_msp,self.tab_unipe,f"A1:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}","A1")
        self.wk_book_msp.delete_tab("UNIPE E POSITIVO")

    def create_paths_and_fill_columns(self,book_path):
        for i in range(2,self.range_for):
            if i == self.range_for - 1:
                self.row_campus_end = self.row_campus_end + 1    
                self.wk_book_msp.create_file_and_paste_content(f"CRUZEIRO ({i}).xlsx",book_path,self.tab_msp,f"A1:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}",self.tab_campus,f"AA{self.row_campus_end}")
            else:    
                self.wk_book_msp.create_file_and_paste_content(f"CRUZEIRO ({i}).xlsx",book_path,self.tab_msp,f"A1:BE{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}",self.tab_campus,f"AA{self.number_division * i}")
        self.wk_book_msp.copy_and_paste(self.tab_campus,self.tab_msp,f"AA{self.number_division}",f"E2:E{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}")
    
    def finalize_operation_message(self,window):
        message_shoot = CompletionMessage.MessagesPosGradEad(window)
        if self.polo_pending == True:
            message_shoot.metadata_pending(self.wk_book_campus.extract_last_filled_row(self.tab_pending,1))
        if self.campus_relat_pending == True:
            message_shoot.pending_campus_from_relation(self.wk_book_relation.extract_last_filled_row(self.tab_pending_campus,2))
        if self.polo_pending == False and self.campus_relat_pending == False:
            message_shoot.no_pendings()
    def save_and_close_rest(self,file_msp,file_exp,file_relat):
        self.wk_book_relation.save_file(file_relat)
        self.wk_book_relation.close_file()
        self.wk_book_msp.save_file(file_msp)
        self.wk_book_msp.close_file()
        self.wk_book_campus.save_file(file_exp)
        self.wk_book_campus.close_file()
        self.wk_book_unipe_positivo.save_file(self.path_save)
        self.wk_book_unipe_positivo.close_file_without_saving()
        self.wk_book_campus_uni_pos.close_file_without_saving()

        