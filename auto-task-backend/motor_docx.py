# motor_docx.py
from generadores import GENERADORES

class GeneradorWord:
    def __init__(self, nombre_archivo="Tarea_Generada.docx", tipo_tarea="general"):
        self.nombre_archivo = nombre_archivo
        GeneradorClase = GENERADORES.get(tipo_tarea, GENERADORES['general'])
        self.generador = GeneradorClase(nombre_archivo)

    def agregar_texto(self, texto):
        for linea in texto.split('\n'):
            self.generador.procesar_linea(linea)
        # Cerrar tabla pendiente al final
        if self.generador._dentro_tabla and self.generador._tabla_buffer:
            self.generador._plasmar_tabla_markdown(self.generador._tabla_buffer)
            self.generador._tabla_buffer = []
            self.generador._dentro_tabla = False

    def guardar(self):
        self.generador.guardar()
