import ArquivoExcel

class CursoTecnico(ArquivoExcel.ArquivoExcel):
    def __init__(self,file_path = None, visibility = True, filtered = False):
        super().__init__()