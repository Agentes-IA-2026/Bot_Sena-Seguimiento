import os
from Core.document_analyzer import extraer_nombres_evidencias

def probar_lectura_guia():
    print("--- 🔍 Probando Motor de Análisis de PDF ---")
    
    # Ruta a tu guía de inducción (ajusta si el nombre cambió un poco)
    ruta_guia = "assets/Guias_Referencia/Asistencia_comercial/Guía_00_Inducción.pdf"
    
    if os.path.exists(ruta_guia):
        print(f"✅ Archivo encontrado: {ruta_guia}")
        evidencias = extraer_nombres_evidencias(ruta_guia)
        
        if evidencias:
            print(f"\n🎯 El bot detectó {len(evidencias)} evidencias requeridas:")
            for ev in evidencias:
                print(f"  - {ev}")
        else:
            print("\n⚠️ No se detectaron nombres de archivos con extensión (.pdf, .png, etc.)")
            print("Revisa si el PDF tiene el texto protegido o como imagen.")
    else:
        print(f"❌ No se encontró el archivo en: {ruta_guia}")
        print("Verifica que el nombre en la carpeta sea EXACTAMENTE igual.")

if __name__ == "__main__":
    probar_lectura_guia()