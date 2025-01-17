import ArquivoExcel
import os
class CursoTecnico(ArquivoExcel.ArquivoExcel):
    def __init__(self,file_path = None, visibility = True, filtered = False):
        super().__init__()

    def spreadsheet_processing(self,file_msp,file_campus):

        campus_name = os.path.basename(file_msp)

        wk_book_msp = ArquivoExcel.ArquivoExcel(file_msp)
        tab_msp = wk_book_msp.select_tab("Modelo Sem Parar")
        tab_polo_portfolio = wk_book_msp.select_tab("Polo X Portfólio Técnico")
        tab_polo_enfermagem = wk_book_msp.select_tab("Polo X Portfólio Tec Enfermagem")

        wk_book_campus = ArquivoExcel.ArquivoExcel(file_campus)
        tab_campus = wk_book_campus.select_tab("Sheet 1")

        #Criar colunas extras
        names_header = ["concat","university_id","campus_id","concat2","ids_concatenados","concat_cursos"]
        for i in range(2,8):
            wk_book_msp.create_column(tab_polo_portfolio,i)
            wk_book_msp.name_header(tab_polo_portfolio,names_header[i - 2])
        
        wk_book_msp.create_column(tab_msp,7)
        wk_book_msp.name_header(tab_msp,7,"cursos")

        wk_book_campus.create_column(tab_campus,6)
        wk_book_campus.create_column(tab_campus,7)
        wk_book_campus.name_header(tab_campus,6,"concat")
        wk_book_campus.name_header(tab_campus,7,"concat2")

