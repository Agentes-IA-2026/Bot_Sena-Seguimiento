import pandas as pd

def cargar_datos_aprendices(ruta_excel):
    try:
        # Leemos el archivo Excel
        df = pd.read_excel(ruta_excel)
        
        # Convertimos los datos a una lista de diccionarios para que sea fácil de usar
        aprendices = df.to_dict(orient='records')
        
        print(f"✅ Se cargaron {len(aprendices)} aprendices correctamente.")
        return aprendices
    
    except Exception as e:
        print(f"❌ Error al leer el Excel: {e}")
        return []

# Esta parte es solo para probar que funcione
if __name__ == "__main__":
    # Prueba local (ajusta el nombre si tu excel se llama distinto)
    datos = cargar_datos_aprendices("assets/listado_aprendices.xlsx")
    print(datos[0] if datos else "No hay datos")