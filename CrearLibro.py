import pandas as pd
import customtkinter
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import datetime

def obtener_ruta_salida(base_path):
    """Genera el nombre Formato_Tom(n).xlsx si el archivo ya existe."""
    filename = "Formato_Tom.xlsx"
    full_path = os.path.join(base_path, filename)
    counter = 1
    
    while os.path.exists(full_path):
        full_path = os.path.join(base_path, f"Formato_Tom({counter}).xlsx")
        counter += 1
    return full_path

def buscar_datos_nave_hacia_atras(df, index_actual, col_nave, col_trg, col_loa, max_filas=20):
    """
    Busca hasta max_filas hacia atrás una fila con el mismo nombre de nave
    y con TRG informado. Si lo encuentra, devuelve TRG y LOA.
    """
    nombre_actual = str(df.loc[index_actual, col_nave]).strip()

    # Posición real del índice actual dentro del dataframe
    posicion_actual = df.index.get_loc(index_actual)

    # Buscar hasta 20 filas hacia atrás
    inicio = max(0, posicion_actual - max_filas)

    for pos in range(posicion_actual - 1, inicio - 1, -1):
        idx_busqueda = df.index[pos]

        nombre_busqueda = str(df.loc[idx_busqueda, col_nave]).strip()

        if nombre_busqueda == nombre_actual:
            trg_busqueda = pd.to_numeric(df.loc[idx_busqueda, col_trg], errors='coerce')

            if not pd.isna(trg_busqueda):
                loa_busqueda = pd.to_numeric(df.loc[idx_busqueda, col_loa], errors='coerce')

                return {
                    "trg": trg_busqueda,
                    "loa": loa_busqueda
                }

    return None


def procesar_datos():
    ruta_input = entry_archivo.get()
    fecha_ini_str = entry_inicio.get()
    fecha_ter_str = entry_fin.get()
    ruta_output_dir = entry_destino.get()

    # Formato Fecha (dd/mm/yyyy h:mm)
    # Nota: En Excel 's' es para segundos, usa 'mm' para minutos
    formato_fecha = 'dd/mm/yyyy h:mm'

    # Formato decimal (Numérico, 2 decimales, sin miles)
    formato_dec = '0.00'

    # Formato Entero (Entero, con separador de miles)
    # El '#' indica dígito opcional, el '0' indica dígito obligatorio
    formato_entero = '#,##0'
    formato_multa = '#,##0.00'

    border_horizontal = Border(
    top=Side(border_style="thin", color="000000"),
    bottom=Side(border_style="thin", color="000000"))

    relleno_rosa =   PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

    relleno_celeste = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid") 

    #Color Amarillo Claro (Hexadecimal: FFFFE0 o FFF9C4 para un crema suave)
    relleno_amarillo = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")                                                                                  

    # Estilo Liviano
    fuente_liviana = Font(name='Aptos Narrow', size=9, color='808080')
    alineacion_sangria = Alignment(indent=1, vertical='center')


    fuente_resultados = Font(name='Aptos NArrow', size=11, bold=True)
    fuente_datos = Font(name = "Aptos Narrow", size = 10, bold = False)

    # 2. Definición de Borde Nulo (para asegurar que el cuerpo no tenga nada)
    sin_bordes = Border(
    top=Side(border_style=None),
    bottom=Side(border_style=None),
    left=Side(border_style=None),
    right=Side(border_style=None))

    if not all([ruta_input, fecha_ini_str, fecha_ter_str, ruta_output_dir]):
        messagebox.showwarning("Error", "Por favor complete todos los campos.")
        return

    try:
        # Convertir fechas de entrada
        f_inicio = datetime.strptime(fecha_ini_str, "%d/%m/%Y")
        f_termino = datetime.strptime(fecha_ter_str, "%d/%m/%Y")

        # Cargar Excel (Asumiendo que no tiene encabezados o detectándolos)
        df = pd.read_excel(ruta_input)
        
        # Mapeo de columnas (A=0, C=2, L=11, M=12, N=13)
        # Col C: Nave, Col L: TRG, Col M: Atraque
        col_nave = df.columns[2]
        col_servicio = df.columns[4]
        col_tipo_nave = df.columns[5]
        col_term = df.columns[6]
        col_sitio = df.columns[7]
        col_loa = df.columns[9]
        col_trg = df.columns[11]
        col_atraque = df.columns[12]
        col_desatraque = df.columns[13]
        col_ton = df.columns[31]
        col_xs = df.columns[53]
        col_cu_met = df.columns[56]
        col_zn = df.columns[57]
        
        # Convertir columna de atraque a datetime
        df[col_atraque] = pd.to_datetime(df[col_atraque],  dayfirst=True,   # Indica que el día va primero (dd/mm/aaaa)
        errors='coerce'  # Si hay un error (texto o celda vacía), lo convierte en "NaT" (Not a Time) en lugar de fallar
        )
        
        # Filtrar por rango,  que el terminal contenga "ATI"
        mask = (
            (df[col_atraque] >= f_inicio) & 
            (df[col_atraque] <= f_termino) & 
            (df[col_term].astype(str).str.contains("ATI", na=False)) # Nuevo filtro ATI
        )
       
        df_filtrado = df.loc[mask].sort_values(by=col_atraque)

        if df_filtrado.empty:
            messagebox.showinfo("Sin resultados", "No se encontraron registros en el rango especificado")
            return

        # Crear nuevo libro
        wb = Workbook()
        # Eliminar hoja por defecto
        default_sheet = wb.active
        wb.remove(default_sheet)
        

        for index, row in df_filtrado.iterrows():
            nombre_nave = str(row[col_nave]).strip()

            # Revisar si el TRG actual viene vacío
            trg_val = pd.to_numeric(row[col_trg], errors='coerce')
            loa_val = pd.to_numeric(row[col_loa], errors='coerce')

            datos_recuperados = None
            es_segunda_recalada_o_repetida = False

            if pd.isna(trg_val):
                datos_recuperados = buscar_datos_nave_hacia_atras(
                    df=df,
                    index_actual=index,
                    col_nave=col_nave,
                    col_trg=col_trg,
                    col_loa=col_loa,
                    max_filas=20
                )

                if datos_recuperados is not None:
                    trg_val = datos_recuperados["trg"]
                    loa_val = datos_recuperados["loa"]
                    es_segunda_recalada_o_repetida = True

            # Crear nombre base de hoja
            sheet_name_base = "".join(x for x in nombre_nave if x.isalnum() or x in " -_")[:31]

            # Si se recuperó TRG desde una fila anterior, agregar (2), (3), etc.
            if es_segunda_recalada_o_repetida:
                contador = 2

                while True:
                    sufijo = f"({contador})"
                    largo_max_base = 31 - len(sufijo)
                    sheet_name = f"{sheet_name_base[:largo_max_base]}{sufijo}"

                    if sheet_name not in wb.sheetnames:
                        break

                    contador += 1
            else:
                sheet_name = sheet_name_base

                # Si por cualquier razón ya existe una hoja con ese nombre,
                # también se numera para evitar error de Excel.
                if sheet_name in wb.sheetnames:
                    contador = 2

                    while True:
                        sufijo = f"({contador})"
                        largo_max_base = 31 - len(sufijo)
                        sheet_name = f"{sheet_name_base[:largo_max_base]}{sufijo}"

                        if sheet_name not in wb.sheetnames:
                            break

                        contador += 1

            ws = wb.create_sheet(title=sheet_name)

            # 1. Desactivar cuadrícula
            ws.sheet_view.showGridLines = False

            #aca deseo que desde la columna H hasta la columna G tengan un ancho de 10
            ws.column_dimensions['B'].width = 27
            ws.column_dimensions['C'].width = 21

            # Ajustar ancho de 10 para columnas desde la H hasta la O
            columnas_cuadro = ['G','H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
            for letra in columnas_cuadro:
                ws.column_dimensions[letra].width = 10

            # --- Formato Numérico para el Cuadro de Bodegas (H20:O27) ---
            # Recorremos desde la fila 20 hasta la 27
            for r in range(20, 28): 
                # Recorremos desde la columna 8 (H) hasta la 15 (O)
                for c in range(7, 15): 
                    celda_cuadro = ws.cell(row=r, column=c)
                    
                    # Aplicamos el formato definido previamente
                    celda_cuadro.number_format = formato_multa
                    
                    # Opcional: Alinear a la derecha para que los decimales queden ordenados
                    celda_cuadro.alignment = Alignment(horizontal='right')

            
            # --- DATOS GENERALES (Texto) ---
            ws.cell(row=3, column=2, value="Nave")
            ws.cell(row=3, column=3, value=nombre_nave)

            ws.cell(row=4, column=2, value="Servicio")
            ws.cell(row=4, column=3, value=str(row[col_servicio]))

            ws.cell(row=5, column=2, value="Viaje")
            #ws.cell(row=5, column=3, value=row[col_viaje]) # Asumiendo columna de viaje

            ws.cell(row=17, column=4, value="Ton/hora")
            ws.cell(row=18, column=4, value="Ton")
            ws.cell(row=19, column=4, value="Hr")
            ws.cell(row=20, column=4, value="Hr")      

            c_loa = ws.cell(row=7, column=3, value=float(row[col_loa]))
            c_loa.number_format = formato_dec
            # LOA
            ws.cell(row=7, column=2, value="LOA")

            if pd.isna(loa_val):
                c_loa = ws.cell(row=7, column=3, value="")
            else:
                c_loa = ws.cell(row=7, column=3, value=float(loa_val))
                c_loa.number_format = formato_dec

            # TRG
            ws.cell(row=8, column=2, value="TRG")

            if pd.isna(trg_val):
                c_trg = ws.cell(row=8, column=3, value="")
            else:
                c_trg = ws.cell(row=8, column=3, value=int(trg_val))
                c_trg.number_format = formato_entero

            ws.cell(row=9, column=2, value="Sitio")        
            ws.cell(row=9, column=3, value=int(row[col_sitio]))

            ws.cell(row=10, column=2, value="Fecha Requerida de Atraque")
            c_fech_req = ws.cell(row=10, column=3)
            c_fech_req.number_format = formato_fecha

            ws.cell(row=11, column=2, value="Fecha amarra primera espia")        
            c_atraque = ws.cell(row=11, column=3, value=row[col_atraque])
            c_atraque.number_format = formato_fecha
            
            ws.cell(row=12, column=2, value="Fecha desamarra ultima espia")        
            c_desatraque = ws.cell(row=12, column=3, value=row[col_desatraque])
            c_desatraque.number_format = formato_fecha

            ws.cell(row=13, column=2, value="Tiempo de Espera")
            ws.cell(row=13, column=3, value='=IF(ISBLANK(C10),"",24*(C11-C10))')

            ws.cell(row=14, column=2, value="Tiempo de ocupación")
            tiempo_ocupacion = ws.cell(row=14, column=3, value="=24*(C12-C11)")
            tiempo_ocupacion.number_format = formato_dec

            for celda in ws["B14:C14"][0]: 
                celda.fill = relleno_rosa

            ws.cell(row=16, column=2, value="Toneladas Transferidas")
            tons = ws.cell(row=16, column=3, value=float(row[col_ton]))
            tons.number_format = formato_entero
            
            ws.cell(row=17, column=2, value="Factor de Rendimiento")
            

            ws.cell(row=18, column=2, value="Bodega Dominante")
            ton_dominante = ws.cell(row=18, column=3, value="=MAX(G27:M27)")
            ton_dominante.number_format = formato_dec

            ws.cell(row=27, column=6, value="Totales")
           
           # --- DATOS GENERALES (Fila 6: Tipo Nave) ---
            ws.cell(row=6, column=2, value="Tipo Nave")
            
            # Paso 1: Convertir a número de forma segura (evita el error de base 10)
            tipo_nave_val = pd.to_numeric(row[col_tipo_nave], errors='coerce')
            tipo_nave_val = int(0 if pd.isna(tipo_nave_val) else tipo_nave_val) # Si es vacío, asume 0

            # Caso Tipo Nave 32 (Container Ship)
            if tipo_nave_val == 32:
                ws.cell(row=6, column=3, value="Container Ship")
                # Sumamos columnas 32 a 39 (el límite 40 es para incluir la 39)
                suma_embarque = pd.to_numeric(row.iloc[32:40], errors='coerce').fillna(0).sum()
                ws.cell(row=17, column=3, value=21)
                ws.cell(row=17, column=4, value="ctnr/hora")
                ws.cell(row=18, column=4, value="ctnr")

                teus=ws.cell(row=18, column=3, value=suma_embarque)
                teus.number_format = formato_entero

            # Caso Tipo Nave 13 (Bulk Carrier)
            elif tipo_nave_val == 13:
                ws.cell(row=6, column=3, value="Bulk Carrier")
                
                # Conversión segura para Zinc
                val_zn = pd.to_numeric(row[col_zn], errors='coerce')
                val_zn = 0 if pd.isna(val_zn) else val_zn
                
                if val_zn > 0:
                    ws.cell(row=17, column=3, value=250)

            ws.cell(row=19, column=2, value="Tiempo de Ocupacion Maximo")
            tom = ws.cell(row=19, column=3, value="=IFERROR(C18/C17, 0)+2")
            tom.number_format = formato_dec
            

            ws.cell(row=20, column=2, value="Tiempo de Inactividad")
            tnt = ws.cell(row=20, column=3, value="=SUM(C21:C27)") 
            tnt.number_format = formato_dec

            ws.cell(row=21, column=2, value="Atraque y Recepcion Nave")
            ws.cell(row=22, column=2, value="Colacion")
            ws.cell(row=23, column=2, value="Descanso Legal Amanecida")
            ws.cell(row=24, column=2, value="Tiempo Muerto Remate")
            ws.cell(row=25, column=2, value="Espera de carga / camiones")
            ws.cell(row=25, column=2, value="Tiempo a la Gira")
            ws.cell(row=26, column=2, value="Atraque 2da recalada")
            ws.cell(row=27, column=2, value="Tiempo Muerto Rem. 2da recalada")
            ws.cell(row=28, column=2, value="Fuerza Mayor")

            # Aplicar formato liviano + fondo amarillo al rango deseado
            # Cambia "B18:B27" por el rango que necesites resaltar
            for fila in ws["B21:C28"]:
                for celda in fila:
                    celda.font = fuente_liviana
                    celda.alignment = alineacion_sangria
                    celda.fill = relleno_amarillo
            for fila in ws["C21:C28"]:
                for celda in fila:
                    celda.number_format = formato_dec

            ws.cell(row=29, column=2, value="Tiempo Maximo Permitido (Real)")
            t_max= ws.cell(row=29, column=3, value="=C19+C20")
            t_max.number_format = '0.00'

            ws.cell(row=30, column=2, value="GAP")
            gap = ws.cell(row=30, column=3, value='=IF(C17=0,"Sin determinar aun",C29-C14)')
            gap.number_format = '0.00'

            ws.cell(row=31, column=2, value="*Si GAP (+) no aplica multa. Si GAP (-) aplica multa")

            for celda in ws["B31:C31"][0]: 
                celda.fill = relleno_celeste

            tit_tom= ws.cell(row=33, column=2, value="Multa Tom")
            tit_tom.font = fuente_resultados

            ws.cell(row=34, column=2, value="Multa Base")
            ws.cell(row=34, column=3, value=0.035)

            ws.cell(row=35, column=2, value="Factor USPPI")
            ws.cell(row=35, column=3, value=1.803)

            ws.cell(row=36, column=2, value="Multa Final")
            m_final= ws.cell(row=36, column=3, value="=C34*C35")
            m_final.number_format = '0.000'
        
            calc_tom= ws.cell(row=38, column=2, value="Calculo Multa Tom")
            calc_tom.font = fuente_resultados

            ws.cell(row=39, column=2, value="Multa Final")
            ws.cell(row=39, column=3, value="=C36")

            ws.cell(row=40, column=2, value="Horas Exceso")
            ws.cell(row=40, column=3, value='=IFERROR(C30*-1, "Sin definir aun")')
            ws.cell(row=41, column=2, value="TRG")
            trg=ws.cell(row=41, column=3, value="=C8")
            trg.number_format = formato_entero

            ws.cell(row=42, column=2, value="MULTA TOM").border = border_horizontal
            m_apl = ws.cell(row=42, column=3, value='=IF(C18=0,"Sin Revision",IF(C40<0,0,C41*C40*C39))')
           
            m_apl.number_format = formato_multa
            m_apl.border = border_horizontal
        
            for celda in ws[42][1:3]: # Accede directamente a la fila 42, columnas B(1) y C(2)
                celda.fill = relleno_rosa


        # --- CONFIGURACIÓN DEL CUADRO BODEGAS  (F18:N27) ---
            
            # 1. Definir los encabezados del cuadro (Fila 19, de F a N)
            encabezados_cuadro = ["", "BOD 1", "BOD 2", "BOD 3", "BOD 4", "BOD 5", "BOD 6", "BOD 7", "TOTALES"]
            

            for i, texto in enumerate(encabezados_cuadro):
                celda = ws.cell(row=19, column=6 + i, value=texto)
                celda.font = fuente_resultados
                celda.alignment = Alignment(horizontal='right')
                celda.border = border_horizontal

            # 4. Configuración del Cuerpo y Totales (Filas 19 a 27)
            for r in range(20, 29): # De 20 hasta 28
                for c in range(6, 15): # Columnas F a N
                    celda_actual = ws.cell(row=r, column=c)
                    
                    if r == 27:
                        # Fila de TOTALES: Aplicar borde horizontal arriba y abajo
                        celda_actual.font = fuente_resultados
                        celda_actual.border = border_horizontal
                    else:
                        # Cuerpo de datos: Sin bordes, solo fuente normal
                        celda_actual.font = fuente_datos
                        celda_actual.border = sin_bordes

            #Revisar al parecer hay un error porque al abri excel envia un mensaje de error para reparar formula

            # 3. FÓRMULAS DE SUMA POR FILA (Columna N)
            # Sumar desde la columna F (6) hasta la M (13) para cada fila
            for r in range(19, 27): # Filas de datos (debajo del encabezado y sobre el total)
                # Ejemplo: =SUM(F19:M19)
                ws.cell(row=r, column=14).value = f"=SUM(G{r}:M{r})"

            # 4. FÓRMULAS DE SUMA POR COLUMNA (Fila 27)
            # Sumar verticalmente desde la fila 18 a la 26 para cada columna
            columnas_a_sumar = ['G', 'H', 'I', 'J', 'K', 'L', 'M','N']

            for letra in columnas_a_sumar:
                # La 'G' es el índice 0 en nuestra lista. 
                # Para que caiga en la columna 7 de Excel: 0 + 7 = 7.
                col_num = columnas_a_sumar.index(letra) + 7                
                # La suma debe ir desde la fila 20 (primer dato) hasta la 26 (último dato)
                ws.cell(row=27, column=col_num).value =f"=SUM({letra}20:{letra}26)"

        # Guardar archivo
        ruta_final = obtener_ruta_salida(ruta_output_dir)
        wb.save(ruta_final)
        messagebox.showinfo("Éxito", f"Archivo creado en:\n{ruta_final}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

# --- Interfaz Gráfica ---
root = tk.Tk()
root.title("Generador de Formato TOM")
root.geometry("500x350")

def seleccionar_archivo():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_archivo.delete(0, tk.END)
    entry_archivo.insert(0, path)

def seleccionar_destino():
    path = filedialog.askdirectory()
    entry_destino.delete(0, tk.END)
    entry_destino.insert(0, path)

# Layout
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

tk.Button(root, text="PROCESAR", command=procesar_datos, bg="green", fg="white", height=2).pack(pady=20)


if __name__ == "__main__":
    root.mainloop()