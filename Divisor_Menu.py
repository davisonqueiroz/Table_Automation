from customtkinter import filedialog
import os
import Menu_Selection
import TableDivisor
import customtkinter as ctk

class TableDivisorMenu:
    def __init__(self):
        self.window = ctk.CTk()
        self.number = None

        self.window.title("Divisor de Tabelas")
        width = 470
        height = 320
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
        text_select =  ctk.CTkLabel(self.window,text= " Selecione a planilha:",text_color= "black",font=("Arial Black", 16),bg_color= "#FFFAFA")
        text_select.place(x=18,y= 36)

        text_number_divisions =  ctk.CTkLabel(self.window,text= " Dividir em:",text_color= "black",font=("Arial Black", 16),bg_color= "#FFFAFA")
        text_number_divisions.place(x=20,y= 77)
        self.msg_error = ""
        self.text_number_divisions =  ctk.CTkLabel(self.window,text_color= "red",font=("Arial Black", 16),bg_color= "#FFFAFA")

        #Botoes

        btn_archive_selection = ctk.CTkButton(self.window,text="Escolher Arquivo",command=self.arquive_command,height= 30,width=90,text_color= "black",corner_radius= 80,
                                     fg_color="#A9A9A9",bg_color= "#FFFAFA",font=("Arial", 14))
        btn_archive_selection.place(x= 215, y = 35)

        btn_generate = ctk.CTkButton(self.window,text="Gerar",command=self.generate_division_command,height= 70,width=180,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 25))
        btn_generate.place(x= 145, y = 200)

        #ComboBox
        numbers = ["--Selecione--","2","3","4","5","6"]
        cb_numbers = ctk.CTkComboBox(self.window,values=numbers,command=self.number_picker)
        cb_numbers.place(x=140,y=80)

        self.window.mainloop()


    def arquive_command(self):
        self.bookDivision = filedialog.askopenfilename(title="Selecione a tabela que deseja dividir",filetypes=(("Arquivo Excel",".xlsx*"),))
        if self.bookDivision:
            self.text_number_divisions.place_forget()
            self.name_reference = self.bookDivision
        else:
            self.update_error(" Necessário selecionar arquivo")
            self.bookDivision = None
    
    def update_error(self,msg_error):
        self.text_number_divisions.configure(text= f"ERRO! {msg_error}")
        self.text_number_divisions.place(x=27,y= 150)
    
    def number_picker(self,choice):
        if choice == "--Selecione--" or not choice:
            self.number = None
        else:
            self.number = int(choice)
    
    def generate_division_command(self):
        if hasattr(self,"bookDivision") and self.number  is not None:
            self.arquive_division = TableDivisor.TableDivisorSimple(file_path=self.bookDivision)
            self.arquive_division.spreadsheet_references_save(self.name_reference)
            self.arquive_division.division_spreadsheet(self.number)
            self.arquive_division.create_new_files_and_paste(self.number,self.name_reference)
            self.text_number_divisions.place_forget()
        elif self.number is None and not hasattr(self,"bookDivision"):
            self.update_error(" Necessário selecionar arquivo e divisões")
        elif not hasattr(self,"bookDivision"):
            self.update_error(" Necessário selecionar arquivo")
        elif self.number is None:
            self.update_error(" Necessário selecionar divisões")
