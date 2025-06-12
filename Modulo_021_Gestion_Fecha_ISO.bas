Attribute VB_Name = "Modulo_021_Gestion_Fecha_ISO"

Option Explicit
Public Function fun830_ObtenerFechaHoraActual() As String
    
    '******************************************************************************
    ' FUNCIÓN: fun830_ObtenerFechaHoraActual
    ' FECHA Y HORA DE CREACIÓN: 2025-06-12 10:39:42 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Devuelve la fecha y hora actual en formato yyyyMMdd_hhmmss para uso
    ' en generación de nombres de archivos, hojas, logs y timestamps del sistema.
    ' Función auxiliar reutilizable desde cualquier parte del proyecto.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables de control de errores
    ' 2. Configuración de optimización de rendimiento
    ' 3. Obtención de fecha y hora actual del sistema
    ' 4. Formateo de fecha en formato yyyyMMdd
    ' 5. Formateo de hora en formato hhmmss
    ' 6. Concatenación de fecha y hora con separador underscore
    ' 7. Validación del resultado generado
    ' 8. Retorno del timestamp formateado
    ' 9. Restauración del entorno en caso de error
    ' 10. Manejo exhaustivo de errores con información detallada
    '
    ' PARÁMETROS: Ninguno
    ' RETORNA: String - Fecha y hora en formato "yyyyMMdd_hhmmss"
    ' EJEMPLO: "20250612_103942"
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para fecha y hora
    Dim dtFechaHoraActual As Date
    Dim strFecha As String
    Dim strHora As String
    Dim strResultado As String
    
    ' Variables para optimización
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnEnableEventsOriginal As Boolean
    
    ' Inicialización
    strFuncion = "fun830_ObtenerFechaHoraActual"
    fun830_ObtenerFechaHoraActual = ""
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables de control de errores
    '--------------------------------------------------------------------------
    lngLineaError = 30
    strFecha = ""
    strHora = ""
    strResultado = ""
    
    '--------------------------------------------------------------------------
    ' 2. Configuración de optimización de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 40
    ' Guardar estado actual para restaurar después
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnEnableEventsOriginal = Application.EnableEvents
    
    ' Optimizar rendimiento (aunque para esta función no es crítico)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 3. Obtención de fecha y hora actual del sistema
    '--------------------------------------------------------------------------
    lngLineaError = 50
    ' Usar función NOW() de VBA que es compatible con todas las versiones
    dtFechaHoraActual = Now()
    
    ' Validar que se obtuvo una fecha válida
    If dtFechaHoraActual = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 8301, strFuncion, _
            "Error al obtener fecha y hora actual del sistema"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Formateo de fecha en formato yyyyMMdd
    '--------------------------------------------------------------------------
    lngLineaError = 60
    ' Usar función FORMAT compatible con Excel 97-365
    strFecha = Format(dtFechaHoraActual, "yyyymmdd")
    
    ' Validar formato de fecha
    If Len(strFecha) <> 8 Then
        Err.Raise ERROR_BASE_IMPORT + 8302, strFuncion, _
            "Error en formato de fecha. Longitud esperada: 8, obtenida: " & Len(strFecha)
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Formateo de hora en formato hhmmss
    '--------------------------------------------------------------------------
    lngLineaError = 70
    ' Usar formato de 24 horas para evitar problemas con AM/PM
    strHora = Format(dtFechaHoraActual, "hhmmss")
    
    ' Validar formato de hora
    If Len(strHora) <> 6 Then
        Err.Raise ERROR_BASE_IMPORT + 8303, strFuncion, _
            "Error en formato de hora. Longitud esperada: 6, obtenida: " & Len(strHora)
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Concatenación de fecha y hora con separador underscore
    '--------------------------------------------------------------------------
    lngLineaError = 80
    ' Usar CHR(95) para el caracter underscore como solicitado
    strResultado = strFecha & Chr(95) & strHora
    
    '--------------------------------------------------------------------------
    ' 7. Validación del resultado generado
    '--------------------------------------------------------------------------
    lngLineaError = 90
    ' Validar longitud total del resultado (8 + 1 + 6 = 15 caracteres)
    If Len(strResultado) <> 15 Then
        Err.Raise ERROR_BASE_IMPORT + 8304, strFuncion, _
            "Error en longitud del resultado. Esperada: 15, obtenida: " & Len(strResultado)
    End If
    
    ' Validar que contiene el separador underscore en la posición correcta
    If Mid(strResultado, 9, 1) <> Chr(95) Then
        Err.Raise ERROR_BASE_IMPORT + 8305, strFuncion, _
            "Error en formato del resultado. Separador underscore no encontrado en posición 9"
    End If
    
    '--------------------------------------------------------------------------
    ' 8. Retorno del timestamp formateado
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun830_ObtenerFechaHoraActual = strResultado
    
    '--------------------------------------------------------------------------
    ' 9. Restauración del entorno
    '--------------------------------------------------------------------------
    lngLineaError = 110
    Application.EnableEvents = blnEnableEventsOriginal
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 10. Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Restaurar configuración del entorno
    On Error Resume Next
    Application.EnableEvents = blnEnableEventsOriginal
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    On Error GoTo 0
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Fecha/Hora obtenida: " & CStr(dtFechaHoraActual) & vbCrLf & _
                      "Fecha formateada: " & strFecha & vbCrLf & _
                      "Hora formateada: " & strHora & vbCrLf & _
                      "Resultado parcial: " & strResultado & vbCrLf & _
                      "Compatibilidad: Excel 97/2003/2007/365"
    
    ' Registrar error en el log si la función está disponible
    On Error Resume Next
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    On Error GoTo 0
    
    ' Retornar cadena vacía en caso de error
    fun830_ObtenerFechaHoraActual = ""
End Function

Public Function fun831_CalcularDuracionTarea(ByVal strFechaInicio As String, _
                                            ByVal strFechaFin As String) As String
    
    '******************************************************************************
    ' FUNCIÓN MEJORADA: fun831_CalcularDuracionTarea
    ' FECHA Y HORA DE CREACIÓN: 2025-06-12 13:35:44 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Calcula la duración entre dos timestamps en formato yyyyMMdd_hhmmss
    ' y devuelve un mensaje formateado con fechas de inicio/fin y duración.
    ' VERSIÓN MEJORADA que incluye fechas legibles en el mensaje de salida.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables de control de errores
    ' 2. Validación básica de parámetros de entrada
    ' 3. Validación de formato usando función auxiliar simplificada
    ' 4. Conversión manual de strings a componentes de fecha/hora
    ' 5. Construcción de fechas usando DateSerial y TimeSerial
    ' 6. Validación lógica de fechas (inicio <= fin)
    ' 7. Formateo de fechas para mostrar en formato legible
    ' 8. Cálculo de diferencia usando DateDiff
    ' 9. Conversión de diferencia total a componentes individuales
    ' 10. Construcción del mensaje completo con fechas y duración
    ' 11. Validación del resultado y retorno
    '
    ' PARÁMETROS:
    ' - strFechaInicio (String): Fecha/hora inicio en formato "yyyyMMdd_hhmmss"
    ' - strFechaFin (String): Fecha/hora fin en formato "yyyyMMdd_hhmmss"
    '
    ' RETORNA: String - Mensaje completo con fechas y duración formateada
    ' EJEMPLO:
    ' "INFORMACIÓN DE DURACIÓN DE TAREA
    '  Fecha y hora de inicio: 12/06/2025 13:30:00
    '  Fecha y hora de finalización: 12/06/2025 14:35:15
    '  La tarea ha consumido
    '  1 días
    '  1 horas
    '  5 minutos
    '  15 segundos"
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para fechas y tiempos
    Dim dtFechaInicio As Date
    Dim dtFechaFin As Date
    Dim dblDiferenciaTotal As Double
    
    ' Variables para componentes de fecha inicio
    Dim intAnoInicio As Integer
    Dim intMesInicio As Integer
    Dim intDiaInicio As Integer
    Dim intHoraInicio As Integer
    Dim intMinutoInicio As Integer
    Dim intSegundoInicio As Integer
    
    ' Variables para componentes de fecha fin
    Dim intAnoFin As Integer
    Dim intMesFin As Integer
    Dim intDiaFin As Integer
    Dim intHoraFin As Integer
    Dim intMinutoFin As Integer
    Dim intSegundoFin As Integer
    
    ' Variables para cálculo de duración
    Dim vdd As Long    ' Días
    Dim vhh As Long    ' Horas
    Dim vmm As Long    ' Minutos
    Dim vss As Long    ' Segundos
    
    ' Variables para formateo de fechas legibles
    Dim strFechaInicioLegible As String
    Dim strFechaFinLegible As String
    
    ' Variable para resultado
    Dim strResultado As String
    
    ' Inicialización
    strFuncion = "fun831_CalcularDuracionTarea"
    fun831_CalcularDuracionTarea = ""
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables de control de errores
    '--------------------------------------------------------------------------
    lngLineaError = 30
    vdd = 0
    vhh = 0
    vmm = 0
    vss = 0
    strResultado = ""
    dtFechaInicio = 0
    dtFechaFin = 0
    strFechaInicioLegible = ""
    strFechaFinLegible = ""
    
    '--------------------------------------------------------------------------
    ' 2. Validación básica de parámetros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 40
    ' Validar que los parámetros no estén vacíos
    If Len(Trim(strFechaInicio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 8311, strFuncion, _
            "Parámetro strFechaInicio está vacío"
    End If
    
    If Len(Trim(strFechaFin)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 8312, strFuncion, _
            "Parámetro strFechaFin está vacío"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Validación de formato usando función auxiliar simplificada
    '--------------------------------------------------------------------------
    lngLineaError = 50
    ' Validar longitud
    If Len(strFechaInicio) <> 15 Then
        Err.Raise ERROR_BASE_IMPORT + 8313, strFuncion, _
            "Longitud incorrecta en strFechaInicio. Esperada: 15, obtenida: " & Len(strFechaInicio)
    End If
    
    If Len(strFechaFin) <> 15 Then
        Err.Raise ERROR_BASE_IMPORT + 8314, strFuncion, _
            "Longitud incorrecta en strFechaFin. Esperada: 15, obtenida: " & Len(strFechaFin)
    End If
    
    ' Validar separador underscore
    If Mid(strFechaInicio, 9, 1) <> Chr(95) Then
        Err.Raise ERROR_BASE_IMPORT + 8315, strFuncion, _
            "Separador underscore no encontrado en posición 9 de strFechaInicio"
    End If
    
    If Mid(strFechaFin, 9, 1) <> Chr(95) Then
        Err.Raise ERROR_BASE_IMPORT + 8316, strFuncion, _
            "Separador underscore no encontrado en posición 9 de strFechaFin"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Conversión manual de strings a componentes de fecha/hora
    '--------------------------------------------------------------------------
    lngLineaError = 60
    ' Extraer componentes de fecha inicio
    intAnoInicio = CInt(Mid(strFechaInicio, 1, 4))      ' Posición 1-4: Ano
    intMesInicio = CInt(Mid(strFechaInicio, 5, 2))      ' Posición 5-6: mes
    intDiaInicio = CInt(Mid(strFechaInicio, 7, 2))      ' Posición 7-8: día
    intHoraInicio = CInt(Mid(strFechaInicio, 10, 2))    ' Posición 10-11: hora
    intMinutoInicio = CInt(Mid(strFechaInicio, 12, 2))  ' Posición 12-13: minuto
    intSegundoInicio = CInt(Mid(strFechaInicio, 14, 2)) ' Posición 14-15: segundo
    
    ' Extraer componentes de fecha fin
    intAnoFin = CInt(Mid(strFechaFin, 1, 4))
    intMesFin = CInt(Mid(strFechaFin, 5, 2))
    intDiaFin = CInt(Mid(strFechaFin, 7, 2))
    intHoraFin = CInt(Mid(strFechaFin, 10, 2))
    intMinutoFin = CInt(Mid(strFechaFin, 12, 2))
    intSegundoFin = CInt(Mid(strFechaFin, 14, 2))
    
    '--------------------------------------------------------------------------
    ' 5. Construcción de fechas usando DateSerial y TimeSerial
    '--------------------------------------------------------------------------
    lngLineaError = 70
    ' Construir fecha inicio
    dtFechaInicio = DateSerial(intAnoInicio, intMesInicio, intDiaInicio) + _
                    TimeSerial(intHoraInicio, intMinutoInicio, intSegundoInicio)
    
    ' Construir fecha fin
    dtFechaFin = DateSerial(intAnoFin, intMesFin, intDiaFin) + _
                 TimeSerial(intHoraFin, intMinutoFin, intSegundoFin)
    
    '--------------------------------------------------------------------------
    ' 6. Validación lógica de fechas (inicio <= fin)
    '--------------------------------------------------------------------------
    lngLineaError = 80
    If dtFechaInicio > dtFechaFin Then
        Err.Raise ERROR_BASE_IMPORT + 8317, strFuncion, _
            "La fecha de inicio debe ser anterior o igual a la fecha de fin. " & _
            "Inicio: " & strFechaInicio & " (" & CStr(dtFechaInicio) & "), " & _
            "Fin: " & strFechaFin & " (" & CStr(dtFechaFin) & ")"
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Formateo de fechas para mostrar en formato legible
    '--------------------------------------------------------------------------
    lngLineaError = 85
    ' Formatear fecha inicio en formato legible dd/mm/yyyy hh:mm:ss
    strFechaInicioLegible = Format(dtFechaInicio, "dd/mm/yyyy hh:mm:ss")
    
    ' Formatear fecha fin en formato legible dd/mm/yyyy hh:mm:ss
    strFechaFinLegible = Format(dtFechaFin, "dd/mm/yyyy hh:mm:ss")
    
    ' Validar que los formatos se generaron correctamente
    If Len(strFechaInicioLegible) = 0 Then
        strFechaInicioLegible = "Error al formatear fecha de inicio"
    End If
    
    If Len(strFechaFinLegible) = 0 Then
        strFechaFinLegible = "Error al formatear fecha de fin"
    End If
    
    '--------------------------------------------------------------------------
    ' 8. Cálculo de diferencia usando DateDiff
    '--------------------------------------------------------------------------
    lngLineaError = 90
    ' Calcular diferencia total en segundos usando DateDiff
    ' Método más robusto y compatible con todas las versiones
    dblDiferenciaTotal = (dtFechaFin - dtFechaInicio) * 86400 ' 86400 segundos por día
    
    '--------------------------------------------------------------------------
    ' 9. Conversión de diferencia total a componentes individuales
    '--------------------------------------------------------------------------
    lngLineaError = 100
    ' Calcular días completos
    vdd = Int(dblDiferenciaTotal / 86400)
    dblDiferenciaTotal = dblDiferenciaTotal - (vdd * 86400)
    
    ' Calcular horas completas del resto
    vhh = Int(dblDiferenciaTotal / 3600)
    dblDiferenciaTotal = dblDiferenciaTotal - (vhh * 3600)
    
    ' Calcular minutos completos del resto
    vmm = Int(dblDiferenciaTotal / 60)
    
    ' Los segundos restantes
    vss = Int(dblDiferenciaTotal - (vmm * 60))
    
    ' Asegurar que los valores están en rangos correctos
    If vhh >= 24 Then
        vdd = vdd + Int(vhh / 24)
        vhh = vhh Mod 24
    End If
    
    If vmm >= 60 Then
        vhh = vhh + Int(vmm / 60)
        vmm = vmm Mod 60
    End If
    
    If vss >= 60 Then
        vmm = vmm + Int(vss / 60)
        vss = vss Mod 60
    End If
    
    '--------------------------------------------------------------------------
    ' 10. Construcción del mensaje completo con fechas y duración
    '--------------------------------------------------------------------------
    lngLineaError = 110
    ' Construir mensaje completo con información de fechas y duración
    strResultado = "INFORMACIÓN DE DURACIÓN DE TAREA" & vbCrLf & vbCrLf & _
                   "Fecha y hora de inicio: " & strFechaInicioLegible & vbCrLf & _
                   "Fecha y hora de finalización: " & strFechaFinLegible & vbCrLf & vbCrLf & _
                   "La tarea ha consumido " & vbCrLf & _
                   CStr(vdd) & " días" & vbCrLf & _
                   CStr(vhh) & " horas" & vbCrLf & _
                   CStr(vmm) & " minutos" & vbCrLf & _
                   CStr(vss) & " segundos" & vbCrLf
    
    '--------------------------------------------------------------------------
    ' 11. Validación del resultado y retorno
    '--------------------------------------------------------------------------
    lngLineaError = 120
    ' Validar que el resultado no esté vacío
    If Len(strResultado) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 8318, strFuncion, _
            "Error al generar mensaje de resultado"
    End If
    
    ' Retornar resultado
    fun831_CalcularDuracionTarea = strResultado
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Fecha Inicio: " & strFechaInicio & vbCrLf & _
                      "Fecha Fin: " & strFechaFin & vbCrLf & _
                      "Fecha Inicio convertida: " & CStr(dtFechaInicio) & vbCrLf & _
                      "Fecha Fin convertida: " & CStr(dtFechaFin) & vbCrLf & _
                      "Fecha Inicio legible: " & strFechaInicioLegible & vbCrLf & _
                      "Fecha Fin legible: " & strFechaFinLegible & vbCrLf & _
                      "Componentes Inicio - Ano: " & CStr(intAnoInicio) & _
                      ", Mes: " & CStr(intMesInicio) & ", Día: " & CStr(intDiaInicio) & _
                      ", Hora: " & CStr(intHoraInicio) & ", Min: " & CStr(intMinutoInicio) & _
                      ", Seg: " & CStr(intSegundoInicio) & vbCrLf & _
                      "Componentes Fin - Ano: " & CStr(intAnoFin) & _
                      ", Mes: " & CStr(intMesFin) & ", Día: " & CStr(intDiaFin) & _
                      ", Hora: " & CStr(intHoraFin) & ", Min: " & CStr(intMinutoFin) & _
                      ", Seg: " & CStr(intSegundoFin) & vbCrLf & _
                      "Diferencia calculada (segundos): " & CStr(dblDiferenciaTotal) & vbCrLf & _
                      "Resultado parcial - Días: " & CStr(vdd) & ", Horas: " & CStr(vhh) & _
                      ", Minutos: " & CStr(vmm) & ", Segundos: " & CStr(vss)
    
    ' Registrar error en el log si la función está disponible
    On Error Resume Next
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    On Error GoTo 0
    
    ' Retornar cadena de error descriptiva en lugar de vacía
    fun831_CalcularDuracionTarea = "Error al calcular duración: " & Err.Description & vbCrLf & _
                                   "Fecha Inicio: " & strFechaInicio & vbCrLf & _
                                   "Fecha Fin: " & strFechaFin
End Function


Public Function fun832_ValidarFormatoFechaHora(ByVal strFechaHora As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun832_ValidarFormatoFechaHora
    ' FECHA Y HORA DE CREACIÓN: 2025-06-12 10:39:42 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Valida que una cadena de texto tenga el formato correcto yyyyMMdd_hhmmss
    ' Función auxiliar para verificar timestamps generados por fun830_ObtenerFechaHoraActual
    '
    ' PARÁMETROS:
    ' - strFechaHora (String): Cadena a validar
    '
    ' RETORNA: Boolean (True si formato correcto, False si incorrecto)
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim strCaracter As String
    
    ' Validar longitud
    If Len(strFechaHora) <> 15 Then
        fun832_ValidarFormatoFechaHora = False
        Exit Function
    End If
    
    ' Validar separador underscore en posición 9
    If Mid(strFechaHora, 9, 1) <> Chr(95) Then
        fun832_ValidarFormatoFechaHora = False
        Exit Function
    End If
    
    ' Validar que los demás caracteres sean numéricos
    For i = 1 To 15
        If i <> 9 Then ' Saltar el separador underscore
            strCaracter = Mid(strFechaHora, i, 1)
            If strCaracter < "0" Or strCaracter > "9" Then
                fun832_ValidarFormatoFechaHora = False
                Exit Function
            End If
        End If
    Next i
    
    fun832_ValidarFormatoFechaHora = True
    Exit Function
    
ErrorHandler:
    fun832_ValidarFormatoFechaHora = False
End Function


Public Function fun833_ConvertirStringADate(ByVal strFechaHora As String) As Date
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun833_ConvertirStringADate
    ' FECHA Y HORA DE CREACIÓN: 2025-06-12 10:39:42 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Convierte una cadena en formato yyyyMMdd_hhmmss a tipo Date de VBA
    ' Compatible con todas las versiones de Excel y configuraciones regionales
    '
    ' PARÁMETROS:
    ' - strFechaHora (String): Cadena con formato yyyyMMdd_hhmmss
    '
    ' RETORNA: Date - Fecha y hora convertida, 0 si error
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    ' Variables para extracción de componentes
    Dim strAno As String
    Dim strMes As String
    Dim strDia As String
    Dim strHora As String
    Dim strMinuto As String
    Dim strSegundo As String
    
    ' Variables para conversión
    Dim intAno As Integer
    Dim intMes As Integer
    Dim intDia As Integer
    Dim intHora As Integer
    Dim intMinuto As Integer
    Dim intSegundo As Integer
    
    Dim dtResultado As Date
    
    ' Validar formato primero
    If Not fun832_ValidarFormatoFechaHora(strFechaHora) Then
        fun833_ConvertirStringADate = 0
        Exit Function
    End If
    
    ' Extraer componentes de fecha
    strAno = Mid(strFechaHora, 1, 4)        ' Posición 1-4: Ano
    strMes = Mid(strFechaHora, 5, 2)        ' Posición 5-6: mes
    strDia = Mid(strFechaHora, 7, 2)        ' Posición 7-8: día
    
    ' Extraer componentes de hora (después del underscore)
    strHora = Mid(strFechaHora, 10, 2)      ' Posición 10-11: hora
    strMinuto = Mid(strFechaHora, 12, 2)    ' Posición 12-13: minuto
    strSegundo = Mid(strFechaHora, 14, 2)   ' Posición 14-15: segundo
    
    ' Convertir a números
    intAno = CInt(strAno)
    intMes = CInt(strMes)
    intDia = CInt(strDia)
    intHora = CInt(strHora)
    intMinuto = CInt(strMinuto)
    intSegundo = CInt(strSegundo)
    
    ' Validar rangos lógicos
    If intAno < 1900 Or intAno > 3000 Then GoTo ErrorHandler
    If intMes < 1 Or intMes > 12 Then GoTo ErrorHandler
    If intDia < 1 Or intDia > 31 Then GoTo ErrorHandler
    If intHora < 0 Or intHora > 23 Then GoTo ErrorHandler
    If intMinuto < 0 Or intMinuto > 59 Then GoTo ErrorHandler
    If intSegundo < 0 Or intSegundo > 59 Then GoTo ErrorHandler
    
    ' Crear fecha usando DateSerial y TimeSerial (compatible con todas las versiones)
    dtResultado = DateSerial(intAno, intMes, intDia) + _
                  TimeSerial(intHora, intMinuto, intSegundo)
    
    fun833_ConvertirStringADate = dtResultado
    Exit Function
    
ErrorHandler:
    fun833_ConvertirStringADate = 0
End Function


Public Function fun834_ExtraerFechaDeCadena(ByVal strFechaHora As String) As String
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun834_ExtraerFechaDeCadena
    ' FECHA Y HORA DE CREACIÓN: 2025-06-12 10:39:42 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Extrae solo la parte de fecha (yyyyMMdd) de una cadena con formato yyyyMMdd_hhmmss
    '
    ' PARÁMETROS:
    ' - strFechaHora (String): Cadena con formato yyyyMMdd_hhmmss
    '
    ' RETORNA: String - Solo la fecha en formato yyyyMMdd
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    ' Validar formato primero
    If Not fun832_ValidarFormatoFechaHora(strFechaHora) Then
        fun834_ExtraerFechaDeCadena = ""
        Exit Function
    End If
    
    ' Extraer los primeros 8 caracteres (fecha)
    fun834_ExtraerFechaDeCadena = Left(strFechaHora, 8)
    Exit Function
    
ErrorHandler:
    fun834_ExtraerFechaDeCadena = ""
End Function


Public Function fun835_ExtraerHoraDeCadena(ByVal strFechaHora As String) As String
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun835_ExtraerHoraDeCadena
    ' FECHA Y HORA DE CREACIÓN: 2025-06-12 10:39:42 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Extrae solo la parte de hora (hhmmss) de una cadena con formato yyyyMMdd_hhmmss
    '
    ' PARÁMETROS:
    ' - strFechaHora (String): Cadena con formato yyyyMMdd_hhmmss
    '
    ' RETORNA: String - Solo la hora en formato hhmmss
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    ' Validar formato primero
    If Not fun832_ValidarFormatoFechaHora(strFechaHora) Then
        fun835_ExtraerHoraDeCadena = ""
        Exit Function
    End If
    
    ' Extraer los últimos 6 caracteres (hora)
    fun835_ExtraerHoraDeCadena = Right(strFechaHora, 6)
    Exit Function
    
ErrorHandler:
    fun835_ExtraerHoraDeCadena = ""
End Function


Public Function fun802_CrearHojaDelimitadores(wb As Workbook, nombreHoja As String) As Worksheet

    ' =============================================================================
    ' FUNCIÓN AUXILIAR 802: CREAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripción: Crea una nueva hoja con el nombre especificado y la deja visible
    ' Parámetros: wb (Workbook), nombreHoja (String)
    ' Retorna: Worksheet (referencia a la hoja creada, Nothing si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lineaError As Long
    
    lineaError = 300
    
    ' Verificar parámetros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Exit Function
    End If
    
    lineaError = 310
    
    ' Verificar que el libro no esté protegido (importante para entornos cloud)
    If wb.ProtectStructure Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Debug.Print "ERROR: No se puede crear hoja, libro protegido - Función: fun802_CrearHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 320
    
    ' Crear nueva hoja al final del libro (método compatible con todas las versiones)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    lineaError = 330
    
    ' Asignar nombre a la hoja
    ws.Name = nombreHoja
    
    lineaError = 340
    
    ' Asegurar que la hoja esté visible
    ws.Visible = xlSheetVisible
    
    lineaError = 350
    
    ' Configuración adicional para compatibilidad con entornos cloud
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    ' Retornar referencia a la hoja creada
    Set fun802_CrearHojaDelimitadores = ws
    
    lineaError = 360
    
    Exit Function
    
ErrorHandler:
    Set fun802_CrearHojaDelimitadores = Nothing
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun802_CrearHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "PARÁMETRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun802_SheetExists(ByVal strSheetName As String) As Boolean
    
    '========================================================================
    ' FUNCION AUXILIAR: fun802_SheetExists
    ' Descripcion : Verifica de forma segura si existe una hoja (worksheet)
    '               con el nombre indicado en el libro actual
    '               antes de entrar a trabajar con ella
    ' Fecha       : 2025-06-01
    ' Retorna     : Boolean
    '========================================================================
    
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    fun802_SheetExists = False
    Set ws = ThisWorkbook.Worksheets(strSheetName)
    If Not ws Is Nothing Then
        fun802_SheetExists = True
    End If
    Exit Function
ErrorHandler:
    fun802_SheetExists = False
End Function

Public Function fun811_DetectarThousandsSeparatorLegacy() As String

    ' =============================================================================
    ' FUNCIÓN AUXILIAR 811: DETECTAR THOUSANDS SEPARATOR (MÉTODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripción: Método alternativo para detectar separador de miles en versiones antiguas
    ' Parámetros: Ninguno
    ' Retorna: String (carácter del separador de miles)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detección
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1200
    
    ' Método alternativo: formatear un número grande y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = Format(1000, "#,##0")
    
    lineaError = 1210
    
    ' El separador de miles es el segundo carácter en números de 4 dígitos
    If Len(numeroFormateado) >= 2 Then
        fun811_DetectarThousandsSeparatorLegacy = Mid(numeroFormateado, 2, 1)
    Else
        ' Si no hay separador visible, asumir coma por defecto
        fun811_DetectarThousandsSeparatorLegacy = ","
    End If
    
    lineaError = 1220
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir coma por defecto
    fun811_DetectarThousandsSeparatorLegacy = ","
    
    ' Información detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCIÓN: fun811_DetectarThousandsSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "LÍNEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "LÍNEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function F010_Abrir_Hoja_Inicial() As Integer

    ' =============================================================================
    ' FUNCION: F010_Abrir_Hoja_Inicial
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Funcion para navegar a la hoja inicial del libro
    ' PARAMETROS: Ninguno
    ' RETORNO: Integer (0=exito, >0=error)
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Configurar la variable vHojaInicial con "00_Ejecutar_Procesos"
    ' 2. Verificar que el libro de trabajo este disponible
    ' 3. Buscar la hoja especificada en la coleccion de hojas del libro
    ' 4. Si la hoja existe, activarla y posicionarse en celda A1
    ' 5. Si la hoja no existe, retornar codigo de error
    ' 6. Retornar codigo de resultado

    On Error GoTo ErrorHandler
    
    Dim vResultado As Integer
    Dim vHojaEncontrada As Boolean
    Dim vContadorHojas As Integer
    Dim vNombreHojaActual As String
    Dim vLineaError As Integer
    
    vResultado = 0
    vHojaEncontrada = False
    vContadorHojas = 0
    vLineaError = 10
    
    ' Paso 1: Configurar la variable vHojaInicial con "00_Ejecutar_Procesos"
    vHojaInicial = "00_Ejecutar_Procesos"
    vLineaError = 20
    
    ' Paso 2: Verificar que el libro de trabajo este disponible
    vLineaError = 30
    If ThisWorkbook Is Nothing Then
        vResultado = 1001 ' Error: Libro de trabajo no disponible
        GoTo ErrorHandler
    End If
    
    ' Paso 3: Buscar la hoja especificada en la coleccion de hojas del libro
    vLineaError = 40
    For vContadorHojas = 1 To ThisWorkbook.Worksheets.Count
        vNombreHojaActual = ThisWorkbook.Worksheets(vContadorHojas).Name
        If StrComp(vNombreHojaActual, vHojaInicial, vbTextCompare) = 0 Then
            vHojaEncontrada = True
            Exit For
        End If
    Next vContadorHojas
    
    ' Paso 4: Si la hoja existe, activarla y posicionarse en celda A1
    vLineaError = 50
    If vHojaEncontrada Then
        ThisWorkbook.Worksheets(vHojaInicial).Activate
        vLineaError = 55
        ThisWorkbook.Worksheets(vHojaInicial).Range("A1").Select
        vResultado = 0 ' Exito
    Else
        ' Paso 5: Si la hoja no existe, retornar codigo de error
        vResultado = 1002 ' Error: Hoja no encontrada
    End If
    
    ' Paso 6: Retornar codigo de resultado
    F010_Abrir_Hoja_Inicial = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F010_Abrir_Hoja_Inicial" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Hoja objetivo: " & vHojaInicial
    
    MsgBox vMensajeError, vbCritical, "Error F010_Abrir_Hoja_Inicial"
    
    If vResultado = 0 Then
        vResultado = 9999 ' Error no especificado
    End If
    
    F010_Abrir_Hoja_Inicial = vResultado
    
End Function


