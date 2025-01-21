from tkinter import messagebox

class MessagesTecnico():
    def __init__(self,root):
        self.root = root

    def no_course_pending(self):
        messagebox.showinfo(title= "Pendência de Cursos", message= "Não foram encontrados cursos com pendência")
    
    def couses_pending(self,amount_pending):
        messagebox.showinfo(title= "Pendência de Cursos", message= f"Foram encontrados{amount_pending} cursos com pendência, os quais foram separados na aba 'Cursos com Pendência'")

    def all_courses_pending(self):
        messagebox.showerror(title= "Todos cursos pendentes", message= "Todos os cursos(exceto Enfermagem) estão como pendentes. Verifique o motivo do erro antes de subir a tabela")

