# generadores/__init__.py
# Todos los tipos usan GeneradorUniversal — un solo generador para todo.

from .base import BaseGenerador
from .generador_universal import GeneradorUniversal

# Alias de compatibilidad (por si algún código importa los nombres antiguos)
GeneradorGeneral     = GeneradorUniversal
GeneradorEstadistica = GeneradorUniversal
GeneradorCiudadania  = GeneradorUniversal
GeneradorIngles      = GeneradorUniversal

# Mapeo de tipos: todos apuntan al mismo generador
GENERADORES = {
    'general':     GeneradorUniversal,
    'matematicas': GeneradorUniversal,
    'estadistica': GeneradorUniversal,
    'ingles':      GeneradorUniversal,
    'ciudadania':  GeneradorUniversal,
}
