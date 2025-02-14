import Cruzeiro_Selection
import customtkinter as ctk
import Divisor_Menu


class MenuSelection:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("Table Automatization")
        width = 400
        height = 470
        self.window.update_idletasks()
        #resolução do sistema
        width_screen = self.window.winfo_screenwidth()
        height_screen =  self.window.winfo_screenheight()
        #posicionamento da janela
        pos_x = (width_screen - width)//2
        pos_y = (height_screen - height)//2
        self.window.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        self.window._set_appearance_mode("light")
        self.window.resizable(False,False)
        self.window.config(background="#FFFAFA")

        #Labels

        text_menu =  ctk.CTkLabel(self.window,text= " Bem vindo ao Menu",text_color= "black",font=("Arial Black", 29),bg_color= "#FFFAFA")
        text_menu.place(x= 38,y= 55)

        #Botoes

        btn_cruzeiro = ctk.CTkButton(self.window,text=" Cruzeiro do Sul",command=self.cruzeiro_command,height= 70,width=280,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 25))
        btn_cruzeiro.place(x= 60, y = 160)

        btn_to_divide = ctk.CTkButton(self.window,text=" Dividir tabela",height= 70,width=280,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial", 25),command=self.divisor_command)
        btn_to_divide.place(x= 60, y = 245)

        btn_exit = ctk.CTkButton(self.window,text=" Sair",height= 50,width=110,text_color= "white",corner_radius= 80,
                                     fg_color="#0000FF",bg_color= "#FFFAFA",font=("Arial Black", 18),command=self.quit_command,hover_color="red")
        btn_exit.place(x= 276, y = 400)

        self.window.mainloop()

    def cruzeiro_command(self):
        self.window.destroy()
        Cruzeiro_Selection.CruzeiroMenuSelection()

    def divisor_command(self):
        self.window.destroy()
        Divisor_Menu.TableDivisorMenu()

    def quit_command(self):
        self.window.quit()