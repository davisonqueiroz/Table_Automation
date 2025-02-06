from tkinter import filedialog
from tkinter import *
import Tecnico
import os
import Pos_Grad_EAD

class MenuSelection:
    def __init__(self):
        pass
    
    def main_window(self):
        self.root = Tk()
        self.root.title("Cruzeiro do Sul")
        self.root.geometry("368x428+600+200")

        #Labels
        Title_label = Label(self.root,text= "Selecione qual modelo deseja padronizar :",font=("Arial", 13))
        Title_label.pack(pady=50)

        #Buttons
        Btn_Tecnico = Button(self.root,text= "Técnico",width=25,height=2,font=("Arial", 12),command=self.tecnico_command)
        Btn_Tecnico.pack(pady= 5.2)
        
        Btn_GradEad = Button(self.root,text= "Pós - Graduação EaD",width=25,height=2,font=("Arial", 12),command=self.pos_grad_EAD_command)
        Btn_GradEad.pack(pady= 5.2)

        Btn_PosEad = Button(self.root,text= "Pós - Graduação Presencial",width=25,height=2,font=("Arial", 12))
        Btn_PosEad.pack(pady= 5.2)


        self.root.mainloop()

    def tecnico_command(self):
        self.root.withdraw()
        bookMsp = filedialog.askopenfilename(title="Selecione a tabela Msp",filetypes=(("Arquivo Excel",".xlsx*"),))
        if not bookMsp:
            self.root.deiconify()
            return
        else:
            bookCampus = filedialog.askopenfilename(title="Selecione a tabela de Campus",filetypes=(("Arquivo Excel",".xlsx*"),))
        if not bookCampus:
            self.root.deiconify()
            return
        
        arquive_tec = Tecnico.CursoTecnico(file_path=bookMsp)
        arquive_tec.spreadsheet_processing(bookMsp,bookCampus)
        arquive_tec.check_nursing_course()
        arquive_tec.filter_order_and_fill()
        arquive_tec.apply_concats()
        arquive_tec.xlooup_and_treat_errors()
        arquive_tec.treatment_names()
        arquive_tec.remove_duplicates_and_apply_textjoin()
        arquive_tec.verify_if_have_pending()
        arquive_tec.finalize_operation_message(self.root)
        msp_save = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Arquivo Excel", "*.xlsx")],initialfile= os.path.basename(bookMsp))
        arquive_tec.saving_files(msp_save)

    def pos_grad_EAD_command(self):
        self.root.withdraw()
        bookMsp = filedialog.askopenfilename(title="Selecione a tabela Msp",filetypes=(("Arquivo Excel",".xlsx*"),))
        if not bookMsp:
            self.root.deiconify()
            return
        else:
            bookExp = filedialog.askopenfilename(title="Selecione a tabela de Campus",filetypes=(("Arquivo Excel",".xlsx*"),))
        if not bookExp:
            self.root.deiconify()
            return
        else:
            bookRelPolos = filedialog.askopenfilename(title="Selecione a tabela de Relação de Polos da IES",filetypes=(("Arquivo Excel",".xlsx*"),))
        if not bookRelPolos:
            self.root.deiconify()
            return
        arquive_pos_grad = Pos_Grad_EAD.PosGraduacaoEAD(file_path=bookMsp)
        arquive_pos_grad.spreadsheet_processing(bookMsp,bookExp,bookRelPolos)
        arquive_pos_grad.check_and_treat_metadata()
        arquive_pos_grad.separate_campus_and_apply_xlookup()
        arquive_pos_grad.check_NAs_and_treat()
        arquive_pos_grad.remove_campus_from_exp()
        arquive_pos_grad.separate_rows_and_concatenate()
        arquive_pos_grad.separate_universities()
        arquive_pos_grad.create_copy_and_separate(bookMsp)
        arquive_pos_grad.create_paths_and_fill_columns(bookMsp)
        arquive_pos_grad.save_and_close_rest(bookMsp,bookExp,bookRelPolos)