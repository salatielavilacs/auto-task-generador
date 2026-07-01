# generadores/ingles.py
import re
from .base import BaseGenerador

class GeneradorIngles(BaseGenerador):
   
    def procesar_linea(self, linea):
        linea = linea.strip()
        if not linea:
            return

        if linea.startswith("### "):
            self.agregar_titulo(linea[4:].strip(), nivel=2)
        elif linea.startswith("## "):
            self.agregar_titulo(linea[3:].strip(), nivel=2)
        elif linea.startswith("# "):
            self.agregar_titulo(linea[2:].strip(), nivel=1)

        elif re.match(r'^\d+\.\d+\.\s+', linea):
            self.agregar_titulo(linea, nivel=2)
        elif re.match(r'^\d+\.\s+[A-ZÁÉÍÓÚÜÑ]', linea):
            self.agregar_titulo(linea, nivel=1)

        elif linea.startswith("- ") or linea.startswith("• "):
            self.agregar_lista(linea[2:])

        elif re.match(r'^\d+\.\s+', linea):
            self.agregar_lista(re.sub(r'^\d+\.\s+', '', linea))

        else:
            linea = re.sub(r'\*\*(.*?)\*\*', r'\1', linea)
            linea = re.sub(r'\*(.*?)\*', r'\1', linea)
            linea = re.sub(r'__(.*?)__', r'\1', linea)
            self.agregar_parrafo(linea)