from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from tkinter import messagebox
import pandas as pd
from datetime import datetime
from ruta_salida import obtener_ruta_salida
from estilos import (formato_fecha, formato_dec, formato_entero, 
                         border_horizontal, border_completo, relleno_gris)
from encabezados import crear_encabezados, obtener_columnas
from excel_sheets import crear_hojas
def procesar_datos(entry_archivo, entry_destino, entry_fin, entry_inicio):
    ruta_input = entry_archivo.get()
    fecha_ini_str = entry_inicio.get()
    fecha_ter_str = entry_fin.get()
    ruta_output_dir = entry_destino.get()

    if not all([ruta_input, fecha_ini_str, fecha_ter_str, ruta_output_dir]):
        messagebox.showwarning("Error", "Por favor complete todos los campos.")
        return

    try:
        # Convertir fechas de entrada
        f_inicio = datetime.strptime(fecha_ini_str, "%d/%m/%Y")
        f_termino = datetime.strptime(fecha_ter_str, "%d/%m/%Y")

        # Cargar Excel (Asumiendo que no tiene encabezados o detectándolos)
        df = pd.read_excel(ruta_input)
        
        cols = obtener_columnas(df)

        col_nave = cols["nave"]
        col_servicio = cols["servicio"]
        col_tipo_nave = cols["tipo_nave"]
        col_term = cols["term"]
        col_sitio = cols["sitio"]
        col_loa = cols["loa"]
        col_trg = cols["trg"]
        col_atraque = cols["atraque"]
        col_desatraque = cols["desatraque"]
        col_ton = cols["ton"]
        col_cu_met = cols["cu_met"]
        col_zn = cols["zn"]
        col_cu = cols["cu"]
        col_cs = cols["cs"]
        col_break = cols["break"]

        
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

        # Crear hoja resumen al principio del libro
        ws_res = wb.create_sheet(title="Res", index=0)
        ws_res.sheet_view.showGridLines = False

                
        crear_encabezados (ws_res, border_completo)
        

        hojas_naves = crear_hojas(wb, df_filtrado, df,  col_nave, col_trg, col_loa, col_servicio, col_sitio, col_atraque, col_desatraque,
                col_tipo_nave, col_ton, col_cs, col_cu_met, col_break, col_zn, col_cu)

        referencias_res = [
                            (2, None, None),          # Nº
                            (3, "C3", None),          # Nave
                            (4, "C6", None),          # Tipo Nave
                            (5, "C8", formato_entero),
                            (6, "C9", None),
                            (7, "C11", formato_fecha),
                            (8, "C12", formato_fecha),
                            (9, "C14", formato_dec),
                            (10, "C18", formato_dec),
                            (11, "C17", formato_dec),
                            (12, "C19", formato_dec),
                            (13, "C20", formato_dec),
                            (14, "C43", formato_dec),
                            (15, "C45", formato_dec),
                            (16, "C16", formato_entero),
                            (17, "C47", formato_dec),
                            (18, "G16", None),
                        ]
        
        fila_res = 2

        for nombre_hoja in hojas_naves:
            ref_hoja = f"'{nombre_hoja}'"
            
            for col_res, celda_origen, formato in referencias_res:
                if col_res == 2:
                    celda = ws_res.cell(row=fila_res, column=col_res, value=fila_res - 1)
                else:
                    celda = ws_res.cell(row=fila_res, column=col_res, value=f"={ref_hoja}!{celda_origen}")

                if formato:
                    celda.number_format = formato

                celda.alignment = Alignment(horizontal="center", vertical="center")
            

            fila_res += 1
        ultima_fila_res = ws_res.max_row

        # Fila donde irá el total
        fila_total = ultima_fila_res + 2

        ws_res.cell(row=fila_total, column=4, value="TOTALES")  # Columna D
        ws_res.cell(row=fila_total, column=4).font = Font(name="Aptos Narrow", size=12, bold=True)
        ws_res.cell(row=fila_total, column=4).alignment = Alignment(horizontal="right", vertical="center")

        
        for celda in ws_res[fila_total][3:16]: # Accede directamente a la fila 47, columnas D(1) y P(2)
                celda.fill = relleno_gris
                celda.border = border_horizontal
                
        ws_res.cell(row=fila_total, column=4).border = Border(left=Side(border_style="thin", color="000000"),
                                                              top=Side(border_style="thin", color="000000"),
                                                              bottom=Side(border_style="thin", color="000000"))

        celda_total_multa = ws_res.cell(row=fila_total,column=15,value=f"=SUM(O2:O{ultima_fila_res})")
        celda_total_multa.border = border_completo
        celda_total_multa.number_format = formato_dec
        celda_total_multa.font = Font(name="Aptos Narrow", size=12, bold=True)
        celda_total_multa.alignment = Alignment(horizontal="center", vertical="center")

        celda_total_ton = ws_res.cell(row=fila_total, column=16, value=f"=SUM(P2:P{ultima_fila_res})")
        celda_total_ton.border = border_completo
        celda_total_ton.number_format = formato_entero
        celda_total_ton.font = Font(name="Aptos Narrow", size=12, bold=True)
        celda_total_ton.alignment = Alignment(horizontal="center", vertical="center")

        


        for fila in ws_res.iter_rows(
            min_row=2,
            max_row=ultima_fila_res,
            min_col=2,   # B
            max_col=18   # R
        ):
            for celda in fila:
                celda.border = border_completo

        # Guardar archivo
        ruta_final = obtener_ruta_salida(ruta_output_dir)
        wb.save(ruta_final)
        messagebox.showinfo("Éxito", f"Archivo creado en:\n{ruta_final}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
    
    return