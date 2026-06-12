from openpyxl.styles import Font, Alignment, Border, Side


def crear_totales_res(
    ws_res,
    ultima_fila_res,
    formato_dec,
    formato_entero,
    border_completo,
    border_horizontal,
    relleno_gris
):
    """
    Crea la fila de totales en la hoja Res.
    Suma Multa TOM, Ton, T.E y Multa TE.
    """

    # Fila donde irá el total
    fila_total = ultima_fila_res + 2

    # Texto TOTALES
    celda_titulo = ws_res.cell(row=fila_total, column=4, value="TOTALES")  # Columna D
    celda_titulo.font = Font(name="Aptos Narrow", size=12, bold=True)
    celda_titulo.alignment = Alignment(horizontal="right", vertical="center")
    celda_titulo.border = Border(
        left=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )

    # Formato general de la fila de totales desde D hasta R
    for celda in ws_res[fila_total][3:18]:  # D a R
        celda.fill = relleno_gris
        celda.border = border_horizontal

    # Totales hoja Res
    totales_res = [
        (15, "O", formato_dec),      # Multa TOM
        (16, "P", formato_entero),   # Ton
        (17, "Q", formato_dec),      # T.E
        (18, "R", formato_dec),      # Multa TE
    ]

    for columna, letra_columna, formato in totales_res:
        celda_total = ws_res.cell(
            row=fila_total,
            column=columna,
            value=f"=SUM({letra_columna}2:{letra_columna}{ultima_fila_res})"
        )

        celda_total.border = border_completo
        celda_total.number_format = formato
        celda_total.font = Font(name="Aptos Narrow", size=12, bold=True)
        celda_total.alignment = Alignment(horizontal="center", vertical="center")

    return fila_total