from tkinter import filedialog
from tkinter import *
import Cruzeiro_Selection

class MenuSelection:
    def __init__(self):
        self.root = Tk()
        self.root.title("Menu Principal")
        self.root.geometry("368x428+600+200")

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