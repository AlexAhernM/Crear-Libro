import tkinter as tk
from tkinter import filedialog
root = tk.Tk()
root.title("Generador de Formato TOM")
root.geometry("500x350")
from procesar import procesar_datos

def seleccionar_archivo():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_archivo.delete(0, tk.END)
    entry_archivo.insert(0, path)

def seleccionar_destino():
    path = filedialog.askdirectory()
    entry_destino.delete(0, tk.END)
    entry_destino.insert(0, path)


tk.Label(root, text="Libro de Origen:").pack(pady=5)
entry_archivo = tk.Entry(root, width=50)
entry_archivo.pack()
tk.Button(root, text="Buscar", command=seleccionar_archivo).pack()

tk.Label(root, text="Fecha Inicio (dd/mm/aaaa):").pack(pady=5)
entry_inicio = tk.Entry(root)
entry_inicio.pack()

tk.Label(root, text="Fecha Término (dd/mm/aaaa):").pack(pady=5)
entry_fin = tk.Entry(root)
entry_fin.pack()

tk.Label(root, text="Carpeta de Destino:").pack(pady=5)
entry_destino = tk.Entry(root, width=50)
entry_destino.pack()
tk.Button(root, text="Buscar", command=seleccionar_destino).pack()

tk.Button(root, text="PROCESAR", command=lambda: procesar_datos(entry_archivo, entry_destino, entry_fin, entry_inicio), bg="green", fg="white", height=2).pack(pady=20)

if __name__ == "__main__":
    root.mainloop()