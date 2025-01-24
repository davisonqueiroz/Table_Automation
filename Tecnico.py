import ArquivoExcel
import os
import CompletionMessage
class CursoTecnico(ArquivoExcel.ArquivoExcel):
    def __init__(self,file_path = None, visibility = True, filtered = False):
        super().__init__()

    def spreadsheet_processing(self,file_msp,file_campus):

        self.campus_name = os.path.basename(file_msp)

        self.wk_book_msp = ArquivoExcel.ArquivoExcel(file_msp)
        self.tab_msp = self.wk_book_msp.select_tab("Modelo Sem Parar")
        self.tab_polo_portfolio = self.wk_book_msp.select_tab("Polo X Portfólio Técnico")
        self.tab_polo_enfermagem = self.wk_book_msp.select_tab("Polo X Portfólio Tec Enfermagem")

        self.wk_book_campus = ArquivoExcel.ArquivoExcel(file_campus)
        self.tab_campus = self.wk_book_campus.select_tab("Sheet 1")

        #Criar colunas extras
        names_header = ["concat","university_id","campus_id","concat2","ids_concatenados","concat_cursos"]
        for i in range(2,8):
            self.wk_book_msp.create_column(self.tab_polo_portfolio,i)
            self.wk_book_msp.name_header(self.tab_polo_portfolio,names_header[i - 2])
        
        self.wk_book_msp.create_column(self.tab_msp,7)
        self.wk_book_msp.name_header(self.tab_msp,7,"cursos")

        self.wk_book_campus.create_column(self.tab_campus,6)
        self.wk_book_campus.create_column(self.tab_campus,7)
        self.wk_book_campus.name_header(self.tab_campus,6,"concat")
        self.wk_book_campus.name_header(self.tab_campus,7,"concat2")

    def check_nursing_course(self):
        self.wk_book_campus.turn_into_text(self.tab_campus,"I",1)
        if self.wk_book_msp.check_name_existence("ENFERMAGEM",8,self.tab_msp) and self.wk_book_msp.check_name_existence("BRAZ CUBAS",2,self.tab_msp):
            self.wk_book_msp.create_column(self.tab_polo_enfermagem,4)
            self.wk_book_msp.create_column(self.tab_polo_enfermagem,5)
            self.wk_book_msp.name_header(self.tab_polo_enfermagem,4,"concat")
            self.wk_book_msp.name_header(self.tab_polo_enfermagem,5,"campus_id")
            nursing_row = self.wk_book_msp.extract_last_filled_row(self.tab_polo_enfermagem,1)
            self.wk_book_campus.create_tab("BrazCubas")
            self.tab_campus_brazcubas = self.wk_book_campus.select_tab("BrazCubas")
            self.wk_book_campus.filter_apply(self.tab_campus,4,"Brazcubas")
            self.wk_book_campus.copy_and_paste(self.tab_campus,self.tab_campus_brazcubas,f"A1:AA{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}","A1")
            self.wk_book_campus.filter_remove(self.tab_campus,4)
            self.wk_book_msp.xlook_up("F2",f"'[{self.campus_name}]BrazCubas'!$I:$I",f"'[{self.campus_name}]BrazCubas'!$A:$A",self.tab_polo_enfermagem,f"E2:E{nursing_row}")
            self.wk_book_msp.filter_apply(self.tab_polo_enfermagem,5,"#N/A")
            self.nursing_pending =  False
            if self.wk_book_msp.verify_filtered(f"A2:I{nursing_row}",self.tab_polo_enfermagem):
                self.nursing_pending = True
                self.wk_book_msp.create_tab("Pendências Enfermagem")
                self.tab_pend_enf = self.wk_book_msp.select_tab("Pendências Enfermagem")
                self.wk_book_msp.copy_and_paste(self.tab_polo_enfermagem,self.tab_pend_enf,f"A1:I{self.wk_book_msp.extract_last_filled_row(self.tab_polo_enfermagem,1)}","A1")
                self.wk_book_msp.delete_filtered_rows(self.tab_polo_enfermagem,f"A2:I{self.wk_book_msp.extract_last_filled_row(self.tab_polo_enfermagem,1)}")
                self.wk_book_msp.filter_remove(self.tab_polo_enfermagem,5)
                self.wk_book_msp.convert_to_value(f"E2:E{self.wk_book_msp.extract_last_filled_row(self.tab_polo_enfermagem,2)}",self.tab_polo_enfermagem)
                self.wk_book_msp.concat_campus_code(self.tab_polo_enfermagem,"E2","F2",f"D2:D{nursing_row}")
                self.wk_book_msp.convert_to_value(f"D2:D{nursing_row + 1}",self.tab_polo_enfermagem)
                self.wk_book_msp.text_join(",",f"D2:D{nursing_row}",self.tab_polo_enfermagem,f"D{nursing_row + 1}")
            else:
                self.wk_book_msp.filter_remove(self.tab_polo_enfermagem,5)
                self.wk_book_msp.convert_to_value(f"E2:E{self.wk_book_msp.extract_last_filled_row(self.tab_polo_enfermagem,2)}",self.tab_polo_enfermagem)
                self.wk_book_msp.concat_campus_code(self.tab_polo_enfermagem,"E2","F2",f"D2:D{nursing_row}")
                self.wk_book_msp.convert_to_value(f"D2:D{nursing_row + 1}",self.tab_polo_enfermagem)
                self.wk_book_msp.text_join(",",f"D2:D{nursing_row}",self.tab_polo_enfermagem,f"D{nursing_row + 1}")
                self.quantity_nursing = 0
                for i in range(2,self.wk_book_msp.extract_last_filled_row(self.tab_msp,2) + 1):
                    if self.wk_book_msp.check_name("ENFERMAGEM",8,self.tab_msp) and self.wk_book_msp.check_name("BRAZ CUBAS",2,self.tab_msp):
                        self.wk_book_msp.copy_and_paste(self.tab_polo_enfermagem,self.tab_msp,f"D{nursing_row + 1}",f"E{i}")
                        self.quantity_nursing += 1
        else:
            if self.wk_book_msp.check_name_existence("ENFERMAGEM",8,self.tab_msp) and self.wk_book_msp.check_name_existence("CRUZEIRO",2,self.tab_msp):
                for i in range(self.wk_book_msp.extract_last_filled_row(self.tab_msp,2) + 1):
                    if self.wk_book_msp.check_name("TÉCNICO ENFERMAGEM",i,8,self.tab_msp) and self.wk_book_msp.check_name("CRUZEIRO",i,2,self.tab_msp):
                        self.wk_book_msp.delete_row(self.tab_msp,i,8)   
        
    def filter_order_and_fill(self):
        self.wk_book_msp.sort_table(self.tab_polo_portfolio,f"A1:Q{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,1)}","H1")
        self.wk_book_msp.filter_apply(self.tab_polo_portfolio,8,"BRAZ CUBAS - TECNICO EAD")
        filter_row = self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,1)
        self.wk_book_msp.filter_remove(self.tab_polo_portfolio,8)

        self.wk_book_msp.fill_with_value(self.tab_polo_portfolio,f"C2:C{filter_row}",46)
        filter_row += 1
        self.wk_book_msp.fill_with_value(self.tab_polo_portfolio,f"C{filter_row}:C{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,1)}",27)

    def apply_concats(self):
        self.wk_book_msp.concat(self.tab_polo_portfolio,'C2','I2',f'B2:B{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,1)}')
        self.wk_book_msp.concat(self.tab_polo_portfolio,'C2','J2',f'E2:E{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,1)}')
        self.wk_book_campus.concat(self.tab_campus,'H2','I2',f'F2:F{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}')
        self.wk_book_campus.concat(self.tab_campus,'H2','C2',f'G2:G{self.wk_book_campus.extract_last_filled_row(self.tab_campus,1)}')

    def xlooup_and_treat_errors(self):
        self.wk_book_msp.xlook_up("B2",f"'[{self.campus_name}]Sheet 1'!$F:$F",f"'[{self.campus_name}]Sheet 1'!$A:$A",self.tab_polo_portfolio,f"D2:D{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,1)}")    
        self.polo_pending = False

        self.wk_book_msp.filter_apply(self.tab_polo_portfolio,4,"#N/A")
        last_row = self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)
        if self.wk_book_msp.verify_filtered(f"D2:D{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}",self.tab_polo_portfolio):
            self.wk_book_msp.create_tab("Polos com pendência")
            self.tab_pend_polos = self.wk_book_msp.select_tab("Polos com pendência")
            self.wk_book_msp.copy_and_paste(self.tab_polo_portfolio,self.tab_pend_polos,f"A1:Q{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}","A1")
            self.wk_book_msp.delete_filtered_rows(self.tab_polo_portfolio,f"A2:Q{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}")
            self.wk_book_msp.filter_remove(self.tab_polo_portfolio,4) 
            self.wk_book_msp.sort_table(self.tab_polo_portfolio,f"A1:Q{last_row}","H1")
            self.polo_pending == True
            
    
    def treatment_names(self):
        self.wk_book_msp.replace(self.tab_polo_portfolio,f"L2:L{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}",' ','')
        self.wk_book_msp.replace(self.tab_polo_portfolio,f"L2:L{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}",'.','')
        self.wk_book_msp.copy_and_paste(self.tab_msp,self.tab_msp,f'H2:H{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}',f'G2:G{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}')
        self.wk_book_msp.replace(self.tab_msp,f"L2:L{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}",' ','')
        self.wk_book_msp.replace(self.tab_msp,f"L2:L{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}",'.','')
        
    def remove_duplicates_and_apply_textjoin(self):
        self.wk_book_msp.concat(self.tab_polo_portfolio,'D2','L2',f'G2:G{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}')
        self.wk_book_msp.remove_duplicates(self.tab_polo_portfolio,f'A2:Q{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}',7)
        self.concat_campus_code(self.tab_polo_portfolio,'D2','I2',f'F2:F{self.wk_book_msp.extract_last_filled_row(self.tab_polo_portfolio,2)}')
        self.convert_to_value(f'F2:F{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}',self.tab_polo_portfolio)

        for cell in range (2,self.wk_book_msp.extract_last_filled_row(self.tab_msp,2) + 1):
            if not self.wk_book_msp.check_name("ENFERMAGEM",cell,8,self.tab_msp):
                formula = self.wk_book_msp.text_join_msp("'Polo X Portfólio Técnico'",cell)
                self.wk_book_msp.formula_apply(self.tab_msp,f"E{cell}",formula)

    def verify_if_have_pending(self):
        self.msp_row = self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)
        self.wk_book_msp.filter_apply(self.tab_msp,5,"#CALC!")
        self.courses_pending = False
        if self.verify_filtered(f"E2:E{self.wk_book_msp.filter_apply(self.tab_msp,2)}",self.tab_msp):
            self.wk_book_msp.create_tab("Cursos com Pendência")
            self.tab_courses_pend = self.wk_book_msp.select_tab("Cursos com Pendência")
            self.wk_book_msp.copy_and_paste(self.tab_msp,self.tab_courses_pend,f"A1:BD{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}","A1")
            self.wk_book_msp.delete_filtered_rows(self.tab_msp,f"A2:BD{self.wk_book_msp.extract_last_filled_row(self.tab_msp,2)}")
            self.wk_book_msp.filter_remove(self.tab_msp,5)
            self.all_courses_pending = False
            self.courses_pending = True
            if self.wk_book_msp.extract_last_filled_row(self.tab_courses_pend,2) == self.msp_row - self.quantity_nursing:
                self.all_courses_pending = True
            
    def finalize_operation_message(self,window):
        message_shoot = CompletionMessage.MessagesTecnico(window)
        if self.polo_pending == True:
            message_shoot.polo_pending()
        if self.all_courses_pending == True:
            message_shoot.all_courses_pending()
        elif self.courses_pending == True:
            message_shoot.couses_pending(self.wk_book_msp.extract_last_filled_row(self.tab_courses_pend,2) - 1)
        if self.nursing_pending == True:
            message_shoot.nursing_pending(self.wk_book_msp.extract_last_filled_row(self.tab_pend_enf,2) - 1)
        
        

