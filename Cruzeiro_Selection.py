from customtkinter import filedialog
import Tecnico
import os
import Pos_Grad_EAD
import Menu_Selection
import customtkinter as ctk

class CruzeiroMenuSelection:
    def __init__(self):
        self.window = ctk.CTk()

        self.window.title("Cruzeiro do Sul")
        width = 470
        height = 520
        #resolução do sistema
        width_screen = self.window.winfo_screenwidth()
        height_screen =  self.window.winfo_screenheight()
        #posicionamento da janela
        pos_x = int(width_screen/2 - width/2)
        pos_y = int(height_screen/2 - height/2)
        self.window.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        self.window._set_appearance_mode("light")
        self.window.resizable(False,False)
        self.window.config(background="#FFFAFA")

        #Labels
        text_menu_cruzeiro =  ctk.CTkLabel(self.window,text= " Selecione qual modelo padronizar:",text_color= "black",font=("Arial Black", 21),bg_color= "#FFFAFA")
        text_menu_cruzeiro.place(x= 38,y= 55)

         #Botoes

        btn_cruzeiro = ctk.CTkButton(self.window,text="Técnico",command=self.tecnico_command,height= 70,width=280,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 25))
        btn_cruzeiro.place(x= 95, y = 160)

        btn_cruzeiro = ctk.CTkButton(self.window,text="Pós-Graduação EaD",command=self.pos_grad_EAD_command,height= 70,width=280,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 20))
        btn_cruzeiro.place(x= 95, y = 240)

        btn_cruzeiro = ctk.CTkButton(self.window,text="Pós-Grad Presencial",height= 70,width=280,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 21))
        btn_cruzeiro.place(x= 95, y = 320)

        btn_cruzeiro = ctk.CTkButton(self.window,text="Voltar",height= 70,width=280,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 25),command=self.return_command)
        btn_cruzeiro.place(x= 95, y = 400)

        self.window.mainloop()

    def tecnico_command(self):
        self.window.withdraw()
        while True:
            bookMsp = filedialog.askopenfilename(title="Selecione a tabela Msp",filetypes=(("Arquivo Excel",".xlsx*"),))
            if not bookMsp:
                self.window.deiconify()
                return
            else:
                bookCampus = filedialog.askopenfilename(title="Selecione a tabela de Campus",filetypes=(("Arquivo Excel",".xlsx*"),))
            if not bookCampus:
                continue
            break
        
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
        self.window.withdraw()
        while True:
            bookMsp = filedialog.askopenfilename(title="Selecione a tabela Msp",filetypes=(("Arquivo Excel",".xlsx*"),))
            if not bookMsp:
                self.window.deiconify()
                return
            else:
                bookExp = filedialog.askopenfilename(title="Selecione a tabela de Campus",filetypes=(("Arquivo Excel",".xlsx*"),))
            if not bookExp:
                continue
            else:
                bookRelPolos = filedialog.askopenfilename(title="Selecione a tabela de Relação de Polos da IES",filetypes=(("Arquivo Excel",".xlsx*"),))
            if not bookRelPolos:
                continue
            else:
                bookExpPosUni = filedialog.askopenfilename(title="Selecione a tabela de exp da Positivo e Unipe",filetypes=(("Arquivo Excel",".xlsx*"),))
            if not bookExpPosUni:
                continue
            break

        arquive_pos_grad = Pos_Grad_EAD.PosGraduacaoEAD(file_path=bookMsp)
        arquive_pos_grad.spreadsheet_processing(bookMsp,bookExp,bookRelPolos,bookExpPosUni)
        arquive_pos_grad.check_and_treat_metadata()
        arquive_pos_grad.separate_campus_and_apply_xlookup()
        arquive_pos_grad.check_NAs_and_treat()
        arquive_pos_grad.remove_campus_from_exp()
        arquive_pos_grad.separate_rows_and_concatenate()
        arquive_pos_grad.separate_universities()
        arquive_pos_grad.create_copy_and_separate(bookMsp)
        arquive_pos_grad.create_paths_and_fill_columns(bookMsp)
        arquive_pos_grad.finalize_operation_message(self.root)
        arquive_pos_grad.save_and_close_rest(bookMsp,bookExp,bookRelPolos)

    def return_command(self):
        self.window.destroy()
        Menu_Selection.MenuSelection()