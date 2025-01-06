import pandas as pd

def process_file(df):
    # Verificar que el DataFrame tenga al menos 3 columnas
    if df.shape[1] < 3:
        print("El archivo no tiene suficientes columnas.")
        return

    # Procesar columnas
    # Primera columna a formato fecha, convirtiendo la fecha al formato 'dd/mm/yyyy'
    df['DATE_TRANSACTION'] = pd.to_datetime(df['DATE_TRANSACTION'], dayfirst=True).dt.strftime('%d/%m/%Y')
    # Segunda y tercera columna a texto (sin decimales)
    # Convertir a string y eliminar el punto decimal en las columnas que tienen valores flotantes
    df.iloc[:, 1] = df.iloc[:, 1].fillna("").astype(str).apply(lambda x: str(int(float(x))) if x else "")
    df.iloc[:, 2] = df.iloc[:, 2].fillna("").astype(str).apply(lambda x: str(int(float(x))) if x else "")

    return df

