from sincronizador_tabla import cargar_actividades
from bot.auditor import MotorAuditoria, EstadoAuditoria, resumen


def main():
    # 1. Cargar actividades de la guía EXACTA como aparece en Supabase
    #    (según tu consulta: Guía_01_Diagnóstico_Empresarial)
    actividades_guia01 = cargar_actividades(guia="Guía_01_Diagnóstico_Empresarial")

    # 2. Cargar TODAS las actividades parametrizadas (todas las guías/programas)
    todas_actividades = cargar_actividades()  # universo completo

    # 3. Simular un aprendiz que solo tiene archivos de otra guía
    archivos_simulados = [
        {
            "nombre": "02_3_1_Segmentacion_cliente.xlsx",
            "extension": ".xlsx",
            "carpeta_padre": "Guia2",
            "folder_path": "Portafolio/Guia2",
            "drive_file_id": "x1",
        },
        {
            "nombre": "foto_estudio.jpg",
            "extension": ".jpg",
            "carpeta_padre": "Fotos",
            "folder_path": "Portafolio/Fotos",
            "drive_file_id": "x2",
        },
    ]

    # 4. Ejecutar el motor de auditoría:
    #    - actividades_guia01 = actividades esperadas de Guía_01_Diagnóstico_Empresarial
    #    - archivos_simulados = archivos encontrados en "Drive"
    #    - todas_actividades  = universo completo para buscar evidencias cruzadas
    motor = MotorAuditoria(
        actividades_guia01,
        archivos_simulados,
        todas_actividades
    )
    resultados = motor.auditar()

    # 5. Resumen general
    stats = resumen(resultados)
    print(
        f"OK: {stats.get('ok', 0)}  |  "
        f"Cruzadas: {stats.get('cruzadas', 0)}  |  "
        f"Faltan: {stats.get('faltantes', 0)}"
    )

    # 6. Listar detalles de las evidencias cruzadas
    for r in resultados:
        if r["estado"] == EstadoAuditoria.EVIDENCIA_CRUZADA.value:
            print(f"\n⚠️  CRUZADA: {r['actividad_id']}")
            print(f"   Archivo encontrado : {r['archivo_encontrado']}")
            print(f"   Observación        : {r['observaciones']}")


if __name__ == "__main__":
    main()