import os
from Core.drive_manager import conectar_drive, verificar_evidencias_en_carpeta
from Core.document_analyzer import extraer_nombres_evidencias

def prueba_final_seguimiento():
    print("--- 🤖 PROCESO DE AUDITORÍA FAROCLICK ---")
    
    # Configuración de prueba
    ID_CARPETA = "16cwnoo9H4kVuVWu3qbAxs_mRrGsUCPpH"
    RUTA_GUIA = "assets/Guias_Referencia/Asistencia_comercial/Guía_00_Inducción.pdf"
    
    # 1. Obtener qué evidencias buscar desde el PDF
    print(f"📄 Analizando guía: {os.path.basename(RUTA_GUIA)}")
    evidencias_requeridas = extraer_nombres_evidencias(RUTA_GUIA)
    
    # 2. Conectar a Drive
    service = conectar_drive()
    
    if service and evidencias_requeridas:
        print(f"✅ Conexión establecida. Revisando {len(evidencias_requeridas)} archivos...")
        
        # 3. Cruzar datos
        reporte = verificar_evidencias_en_carpeta(service, ID_CARPETA, evidencias_requeridas)
        
        print("\n📝 REPORTE DEL APRENDIZ:")
        for archivo, estado in reporte.items():
            status = "✔️ RECIBIDO" if estado else "❌ PENDIENTE"
            print(f"  [{status}] - {archivo}")
            
        completado = sum(reporte.values())
        print(f"\n📊 Avance: {completado}/{len(evidencias_requeridas)} evidencias.")
    else:
        print("❌ Fallo en la conexión o no se encontraron evidencias en el PDF.")

if __name__ == "__main__":
    prueba_final_seguimiento()