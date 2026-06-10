
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill



    # ---------------formatos ---------

formato_fecha = 'dd/mm/yyyy h:mm'

formato_dec = '#,##0.00'

formato_entero = '#,##0'

# ----------------bordes -------------------

border_horizontal = Border(
top=Side(border_style="thin", color="000000"),
bottom=Side(border_style="thin", color="000000"))

border_completo = Border(
top=Side(border_style="thin", color="000000"),
bottom=Side(border_style="thin", color="000000"),
left=Side(border_style="thin", color="000000"),
right=Side(border_style="thin", color="000000")
)
sin_bordes = Border(
top=Side(border_style=None),
bottom=Side(border_style=None),
left=Side(border_style=None),
right=Side(border_style=None))

# ----------------rellenos --------------------

relleno_rosa =   PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

relleno_celeste = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid")

relleno_gris =    PatternFill(start_color="F2F2F2", end_color = "F2F2F2", fill_type="solid") 

relleno_amarillo = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")                                                                                  

# ------------fuente y alineacion --------------
fuente_liviana = Font(name='Aptos Narrow', size=9, color='808080')
alineacion_sangria = Alignment(indent=1, vertical='center')

fuente_resultados = Font(name='Aptos NArrow', size=11, bold=True)
fuente_datos = Font(name = "Aptos Narrow", size = 10, bold = False)





def crear_encabezados (ws_res, border_completo):

    # Encabezados hoja Res desde B1 hasta R1
    encabezados_res = [
        "Nº",
        "Nave",
        "Tipo Nave",
        "TRG",
        "Sitio",
        "Primera Espia",
        "Ultima Espia",
        "T.O",
        "B. Dom",
        "F.Rend",
        "TOM",
        "TNT",
        "GAP",
        "Multa Inicial",
        "Ton",
        "Vel (ton/hr)",
        "TIPO DE CARGA"
    ]
    ws_res.column_dimensions['B'].width = 4
    ws_res.column_dimensions['C'].width = 24
    ws_res.column_dimensions['D'].width = 16
    ws_res.column_dimensions['E'].width = 10
    ws_res.column_dimensions['F'].width = 7
    ws_res.column_dimensions['G'].width = 16
    ws_res.column_dimensions['H'].width = 16
    ws_res.column_dimensions['I'].width = 8
    ws_res.column_dimensions['J'].width = 9
    ws_res.column_dimensions['K'].width = 9
    ws_res.column_dimensions['L'].width = 8
    ws_res.column_dimensions['M'].width = 8
    ws_res.column_dimensions['N'].width = 8
    ws_res.column_dimensions['O'].width = 12
    ws_res.column_dimensions['P'].width = 10
    ws_res.column_dimensions['Q'].width = 10
    ws_res.column_dimensions['R'].width = 35

    for i, texto in enumerate(encabezados_res, start=2):  # B=2 hasta R=18
        celda = ws_res.cell(row=1, column=i, value=texto)
        celda.font = Font(name="Aptos Narrow", size=10, bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.fill = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid")
        celda.border = border_completo