import pandas as pd
import os

def load_database(workfolder_path, db_name):
    """Carga la base de datos desde el archivo Excel."""
    db_path = os.path.join(workfolder_path, 'database', db_name)
    try:
        with open(db_path, mode='rb') as fp:
            df = pd.read_excel(fp, sheet_name='PRODUCTOS', engine='openpyxl')
        print("Base de datos cargada correctamente.")
        return df
    except Exception as e:
        print(f"Error al cargar la base de datos: {e}")
        return None

def filter_products(df):
    """Filtra los productos por aliado."""
    EPM = 'EMPRESAS PÚBLICAS DE MEDELLÍN'
    COMFAMA = 'COMFAMA'

    epm_products = df[df.ALIADO == EPM].PRODUCTO.to_list()
    comfama_products = df[df.ALIADO == COMFAMA].PRODUCTO.to_list()
    other_products = df[(df.ALIADO != COMFAMA) & (df.ALIADO != EPM)].PRODUCTO.to_list()

    products = {
        'EPM': tuple(epm_products),
        'COMFAMA': tuple(comfama_products),
        'OTROS': tuple(other_products)
    }
    return products

def main():
    workfolder_path = r"C:\Users\correo.ejemplo\OneDrive\NotasTrasladosPY"
    db_name = 'BD Productos Conciliaciones V1.xlsx'

    # Cargar la base de datos
    df = load_database(workfolder_path, db_name)

    if df is not None:
        # Filtrar y obtener los productos
        products = filter_products(df)
        print("Productos actualizados correctamente.")

        # Mostrar los productos obtenidos
        for aliado, productos in products.items():
            print(f"Productos para {aliado}: {productos}")
    else:
        print("No se han podido actualizar los productos.")

if __name__ == "__main__":
    main()
