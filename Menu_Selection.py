from tkinter import filedialog
from tkinter import *
import Cruzeiro_Selection

class MenuSelection:
    def __init__(self):
        self.root = Tk()
        self.root.title("Menu Principal")
        width = 380
        height = 440
        #resolução do sistema
        width_screen = self.root.winfo_screenwidth()
        height_screen =  self.root.winfo_screenheight()
        #posicionamento da janela
        pos_x = int(width_screen/2 - width/2)
        pos_y = int(height_screen/2 - height/2)
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

        #Labels
        Title_label = Label(self.root,text= "Selecione uma das seguintes opções : ",font=("Arial", 13))
        Title_label.pack(pady=50)

        #Buttons
        Btn_Cruzeiro = Button(self.root,text= "Cruzeiro do Sul",width=25,height=2,font=("Arial", 12),command=self.cruzeiro_command)
        Btn_Cruzeiro.pack(pady= 5.2)
        
        Btn_to_divide = Button(self.root,text= "Dividir tabela",width=25,height=2,font=("Arial", 12))
        Btn_to_divide.pack(pady= 5.2)

        Btn_to_divide = Button(self.root,text= "Sair",width=25,height=2,font=("Arial", 12),command= self.quit_command)
        Btn_to_divide.pack(pady= 5.2)
        self.root.mainloop()

    def cruzeiro_command(self):
        self.root.destroy()
        Cruzeiro_Selection.CruzeiroMenuSelection()

    def quit_command(self):
        self.root.quit()