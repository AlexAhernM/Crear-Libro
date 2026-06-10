import pandas as pd

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