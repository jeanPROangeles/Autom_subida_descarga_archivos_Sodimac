import os
import shutil

def eliminar_carpeta(ruta_carpeta):
    # Verifica si la carpeta existe y luego elimínala
    if os.path.exists(ruta_carpeta) and os.path.isdir(ruta_carpeta):
        shutil.rmtree(ruta_carpeta)
        print(f'La carpeta "{ruta_carpeta}" ha sido eliminada.')
    else:
        print(f'La carpeta "{ruta_carpeta}" no existe.')

