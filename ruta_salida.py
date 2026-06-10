import os
def obtener_ruta_salida(base_path):
    """Genera el nombre Formato_Tom(n).xlsx si el archivo ya existe."""
    filename = "Formato_Tom.xlsx"
    full_path = os.path.join(base_path, filename)
    counter = 1
    
    while os.path.exists(full_path):
        full_path = os.path.join(base_path, f"Formato_Tom({counter}).xlsx")
        counter += 1
    return full_path