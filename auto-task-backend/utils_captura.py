from PIL import ImageGrab

def guardar_portapapeles(nombre="temp_imagen.png"):
    img = ImageGrab.grabclipboard()
    if img:
        img.save(nombre, 'PNG')
        return nombre
    return None