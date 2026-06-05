# CHANGELOG

## v1.0.3 — Índice de archivos y reanudación de procesos
- Índice de archivos (`BuildFileIndex`) para búsqueda eficiente de adjuntos por nombre base
- Indexación recursiva de subcarpetas al seleccionar la carpeta de adjuntos
- Soporte para múltiples archivos con el mismo nombre base en distintas rutas
- Reanudación automática desde el primer registro pendiente (`PENDING`, `ERROR`, `DISPLAY` o vacío)
- Omisión de filas ya enviadas (`SENT`) en ejecuciones posteriores
- Columnas de estado renombradas a `STATUS` y `STATUS_DATE`
- Valores de estado normalizados: `Sent`, `Display`, `Error`
- Retardo entre envíos con `Application.Wait` en lugar de bucle activo
- Limpieza explícita de objetos COM y manejo de errores por fila con marca de estado
- Restauración de macros de entrada `PROCESAR_PREVISUALIZAR` y `PROCESAR_ENVIAR`

---

## v1.0.2 — Paginación en vista previa
- Paginación configurable en modo previsualización (correos por página)
- Pausa interactiva entre páginas para revisar borradores antes de continuar

---

## v1.0.1 — Validación y refactorización de módulos
- Renombrado de módulos: `hermes.bas`, `ProcessRow.bas`, `Utils.bas`, `Variables.bas`
- Validación previa de hoja `BASE` y columnas obligatorias (`ValidateSheet`)
- Constante `DEFAULT_SHEET_NAME` para centralizar el nombre de la hoja
- Columnas dinámicas de estado con detección y creación automática (`GetStatusColumns`)
- Reemplazo de variables usando `.Text` para preservar formato de celdas
- Reemplazo protegido de variables en HTML respetando imágenes incrustadas (`ApplyVariablesToHTMLProtected`)
- Registro de estado diferenciado por modo de ejecución

---

## v1.0.0 — Primera versión pública
- Arquitectura modular separada en 4 módulos .bas
- Soporte para múltiples plantillas .OFT (una por fila)
- Variables dinámicas ilimitadas mediante diccionarios
- Reemplazo automático de variables en el HTML del correo
- Adjuntos detectados por nombre base
- Marcas de procesamiento en columnas dinámicas
- Documentación en `/docs`
- Archivos de ejemplo incluidos en `/examples`

---

## Pendientes para futuras versiones
- Validación automática de adjuntos faltantes
- Manejo de logs externos
- Exportación de reportes de envío
- Compatibilidad con Gmail API (modo avanzado)
