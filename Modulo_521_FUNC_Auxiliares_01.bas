Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_01"

Option Explicit

Public Function fun801_LogMessage(ByVal strMessage As String, _
                                Optional ByVal blnIsError As Boolean = False, _
                                Optional ByVal strFileName As String = "", _
                                Optional ByVal strSheetName As String = "") As Boolean
        
    '------------------------------------------------------------------------------
    ' FUNCI�N: fun801_LogMessage
    ' PROP�SITO: Sistema integral de logging para registrar eventos y errores
    '
    ' PAR�METROS:
    ' - strMessage (String): Mensaje a registrar
    ' - blnIsError (Boolean, Opcional): True=ERROR, False=INFO (defecto: False)
    ' - strFileName (String, Opcional): Archivo relacionado (defecto: "NA")
    ' - strSheetName (String, Opcional): Hoja relacionada (defecto: "NA")
    '
    ' RETORNA: Boolean - True si exitoso, False si error
    '
    ' FUNCIONALIDADES:
    ' - Crea hoja de log autom�ticamente con formato profesional
    ' - Timestamp ISO, usuario del sistema, tipo de evento
    ' - Formato condicional para errores (fondo rojo)
    ' - Filtros autom�ticos y ajuste de columnas
    '
    ' COMPATIBILIDAD: Excel 97-365, Office Online, SharePoint, Teams
    '
    ' EJEMPLO: Call fun801_LogMessage("Operaci�n completada", False, "datos.csv")
    '------------------------------------------------------------------------------
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para el log
    Dim wsLog As Worksheet
    Dim lngLastRow As Long
    Dim strDateTime As String
    Dim strUserName As String
    Dim strLogType As String
    
    ' Inicializaci�n
    strFuncion = "fun801_LogMessage"
    fun801_LogMessage = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Verificar hoja de log
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If Not fun802_SheetExists(gstrHoja_Log) Then
        If Not F002_Crear_Hoja(gstrHoja_Log) Then
            MsgBox "Error al crear la hoja de log", vbCritical
            Exit Function
        End If
        
        ' Crear y formatear encabezados
        With ThisWorkbook.Sheets(gstrHoja_Log)
            ' Establecer textos de encabezados exactamente como se solicita
            .Range("A1").Value = "Date/Time"
            .Range("B1").Value = "User"
            .Range("C1").Value = "Type"
            .Range("D1").Value = "File"
            .Range("E1").Value = "Sheet"
            .Range("F1").Value = "Message"
            
            ' Formato de encabezados
            With .Range("A1:F1")
                .Font.Bold = True
                .Font.Size = 11
                .Font.Name = "Calibri"
                .Interior.Color = RGB(200, 200, 200)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' Formato espec�fico para la columna de fecha
            .Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
            
            ' Ajustar anchos de columna
            .Columns("A").ColumnWidth = 20  ' Date/Time
            .Columns("B").ColumnWidth = 15  ' User
            .Columns("C").ColumnWidth = 15  ' Type
            .Columns("D").ColumnWidth = 40  ' File
            .Columns("E").ColumnWidth = 20  ' Sheet
            .Columns("F").ColumnWidth = 60  ' Message
            
            ' Filtros autom�ticos
            .Range("A1:F1").AutoFilter
        End With
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Preparar datos para el log
    '--------------------------------------------------------------------------
    lngLineaError = 55
    Set wsLog = ThisWorkbook.Sheets(gstrHoja_Log)
    
    ' Obtener �ltima fila
    lngLastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Preparar datos (reemplazar valores vac�os con "NA")
    strDateTime = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    strUserName = IIf(Environ("USERNAME") = "", "NA", Environ("USERNAME"))
    strLogType = IIf(blnIsError, "ERROR", "INFO")
    strFileName = IIf(Len(Trim(strFileName)) = 0, "NA", strFileName)
    strSheetName = IIf(Len(Trim(strSheetName)) = 0, "NA", strSheetName)
    strMessage = IIf(Len(Trim(strMessage)) = 0, "NA", strMessage)
    
    '--------------------------------------------------------------------------
    ' 3. Escribir en el log
    '--------------------------------------------------------------------------
    lngLineaError = 70
    With wsLog
        ' Escribir datos
        .Cells(lngLastRow, 1).Value = strDateTime    ' Date/Time
        .Cells(lngLastRow, 2).Value = strUserName    ' User
        .Cells(lngLastRow, 3).Value = strLogType     ' Type
        .Cells(lngLastRow, 4).Value = strFileName    ' File
        .Cells(lngLastRow, 5).Value = strSheetName   ' Sheet
        .Cells(lngLastRow, 6).Value = strMessage     ' Message
        
        ' Formato de la nueva fila
        With .Range(.Cells(lngLastRow, 1), .Cells(lngLastRow, 6))
            ' Formato general
            .Font.Name = "Calibri"
            .Font.Size = 10
            .VerticalAlignment = xlTop
            .WrapText = True
            
            ' Bordes
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            
            ' Formato condicional para errores
            If blnIsError Then
                .Interior.Color = RGB(255, 200, 200)
                .Font.Bold = True
            End If
        End With
        
        ' Asegurar formato de fecha en la columna A
        .Cells(lngLastRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    
    fun801_LogMessage = True
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    MsgBox strMensajeError, vbCritical, "Error en sistema de logging"
    fun801_LogMessage = False
End Function

Public Function F002_Crear_Hoja(ByVal strNombreHoja As String) As Boolean

    '******************************************************************************
    ' M�dulo: F002_Crear_Hoja
    ' Fecha y Hora de Creaci�n: 2025-05-26 09:17:15 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Funci�n para crear hojas en el libro con formato y configuraci�n est�ndar
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para manejo de hojas
    Dim ws As Worksheet
    
    ' Inicializaci�n
    strFuncion = "F002_Crear_Hoja"
    F002_Crear_Hoja = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Verificar si la hoja ya existe
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If fun802_SheetExists(strNombreHoja) Then
        F002_Crear_Hoja = True
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Crear nueva hoja
    '--------------------------------------------------------------------------
    lngLineaError = 40
    Application.ScreenUpdating = False
    
    ' Crear hoja al final del libro
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    ' Asignar nombre
    ws.Name = strNombreHoja
    
    ' Configuraci�n b�sica
    'With ws
    '    ' Ajustar vista
    '    .DisplayGridlines = True
    '    .DisplayHeadings = True
    '
    '    ' Configurar primera vista
    '    .Range("A1").Select
    '
    '    ' Ajustar ancho de columnas est�ndar
    '    .Columns.StandardWidth = 10
    '
    '    ' Configurar �rea de impresi�n
    '    .PageSetup.PrintArea = ""
    'End With
    
    Application.ScreenUpdating = True
    
    F002_Crear_Hoja = True
    Exit Function

GestorErrores:
    ' Restaurar configuraci�n
    Application.ScreenUpdating = True
    
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F002_Crear_Hoja = False
End Function

Public Function fun801_LimpiarHoja(ByVal strNombreHoja As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun801_LimpiarHoja
    ' FECHA Y HORA DE CREACI�N: 2025-05-28 17:50:26 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa    '
    ' PROP�SITO:
    ' Limpia de forma segura y eficiente todo el contenido de una hoja de c�lculo
    ' espec�fica, preservando el formato y estructura, pero eliminando todos los
    ' datos y valores almacenados en las celdas utilizadas.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(strNombreHoja)
    
    Application.ScreenUpdating = False
    ws.UsedRange.ClearContents
    Application.ScreenUpdating = True
    
    fun801_LimpiarHoja = True
    Exit Function
    
GestorErrores:
    fun801_LimpiarHoja = False
End Function

Public Function fun802_SeleccionarArchivo(ByVal strPrompt As String) As String
    
    '******************************************************************************
    ' FUNCI�N: fun802_SeleccionarArchivo (VERSI�N MEJORADA)
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' FECHA MODIFICACI�N: 2025-06-01
    ' PROP�SITO:
    ' Proporciona una interfaz de usuario intuitiva para seleccionar archivos de
    ' texto (TXT y CSV) con sistema de carpetas de respaldo autom�tico.
    '
    ' L�GICA DE CARPETAS DE RESPALDO:
    ' 1. Carpeta del archivo Excel actual
    ' 2. %TEMP% (si hay error)
    ' 3. %TMP% (si hay error)
    ' 4. %USERPROFILE% (si hay error)
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para carpetas de respaldo
    Dim strCarpetaInicial As String
    Dim strCarpetaActual As String
    Dim intIntentoActual As Integer
    Dim blnCarpetaValida As Boolean
    
    ' Inicializaci�n
    strFuncion = "fun802_SeleccionarArchivo"
    fun802_SeleccionarArchivo = ""
    lngLineaError = 0
    intIntentoActual = 1
    blnCarpetaValida = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Intentar obtener carpetas de respaldo en orden de prioridad
    '--------------------------------------------------------------------------
    Do While intIntentoActual <= 4 And Not blnCarpetaValida
        lngLineaError = 40 + intIntentoActual
        
        Select Case intIntentoActual
            Case 1: ' Carpeta del archivo Excel actual
                strCarpetaActual = fun808_ObtenerCarpetaSistema("EXCEL_PATH_CURRENT_BOOK")
                
            Case 2: ' Variable de entorno %TEMP%
                strCarpetaActual = fun808_ObtenerCarpetaSistema("TEMP")
                
            Case 3: ' Variable de entorno %TMP%
                strCarpetaActual = fun808_ObtenerCarpetaSistema("TMP")
                
            Case 4: ' Variable de entorno %USERPROFILE%
                strCarpetaActual = fun808_ObtenerCarpetaSistema("USERPROFILE")
        End Select
        
        ' Verificar si la carpeta es v�lida y accesible
        If fun809_ValidarCarpeta(strCarpetaActual) Then blnCarpetaValida = True
            strCarpetaInicial = strCarpetaActual
        Else
            intIntentoActual = intIntentoActual + 1
        End If
    Loop
    
    ' Si no se pudo obtener ninguna carpeta v�lida, usar carpeta por defecto
    If Not blnCarpetaValida Then
        strCarpetaInicial = ""
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Mostrar di�logo de selecci�n de archivo
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    On Error GoTo GestorErrores
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = strPrompt
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt;*.csv"
        .AllowMultiSelect = False
        
        ' Establecer carpeta inicial si es v�lida
        If Len(strCarpetaInicial) > 0 Then
            .InitialFileName = strCarpetaInicial & "\"
        End If
        
        If .Show = -1 Then
            fun802_SeleccionarArchivo = .SelectedItems(1)
        Else
            fun802_SeleccionarArchivo = ""
        End If
    End With
    
    Exit Function
    
GestorErrores:
    ' Log del error
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Intento actual: " & intIntentoActual
    
    fun801_LogMessage strMensajeError, True
    fun802_SeleccionarArchivo = ""
End Function

Public Function fun803_ImportarArchivo(ByRef wsDestino As Worksheet, _
                                     ByVal strFilePath As String, _
                                     ByVal strColumnaInicial As String, _
                                     ByVal lngFilaInicial As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun803_ImportarArchivo
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' PROP�SITO:
    ' Importa el contenido completo de archivos de texto plano (TXT/CSV) l�nea por
    ' l�nea hacia una hoja de Excel espec�fica, colocando cada l�nea del archivo
    ' en una celda individual seg�n la posici�n inicial definida. Funci�n core
    ' del sistema de importaci�n de datos de presupuesto.
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim objFSO As Object
    Dim objFile As Object
    Dim strLine As String
    Dim lngRow As Long
    
    ' Inicializaci�n
    strFuncion = "fun803_ImportarArchivo"
    fun803_ImportarArchivo = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros
    '--------------------------------------------------------------------------
    lngLineaError = 20
    If wsDestino Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, "Hoja de destino no v�lida"
    End If
    
    If Len(strFilePath) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, "Ruta de archivo no v�lida"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Configurar objetos para lectura de archivo
    '--------------------------------------------------------------------------
    lngLineaError = 35
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFilePath, 1) ' ForReading = 1
    
    '--------------------------------------------------------------------------
    ' 3. Leer archivo l�nea por l�nea
    '--------------------------------------------------------------------------
    lngLineaError = 45
    lngRow = lngFilaInicial
    
    While Not objFile.AtEndOfStream
        strLine = objFile.ReadLine
        wsDestino.Range(strColumnaInicial & lngRow).Value = strLine
        lngRow = lngRow + 1
    Wend
    
    '--------------------------------------------------------------------------
    ' 4. Limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 60
    objFile.Close
    Set objFile = Nothing
    Set objFSO = Nothing
    
    fun803_ImportarArchivo = True
    Exit Function

GestorErrores:
    ' Limpieza en caso de error
    If Not objFile Is Nothing Then
        objFile.Close
        Set objFile = Nothing
    End If
    Set objFSO = Nothing
    
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun803_ImportarArchivo = False
End Function

Public Function fun804_DetectarRangoDatos(ByRef ws As Worksheet, _
                                         ByRef lngLineaInicial As Long, _
                                         ByRef lngLineaFinal As Long) As Boolean
    '******************************************************************************
    ' FUNCI�N: fun804_DetectarRangoDatos
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' PROP�SITO:
    ' Detecta autom�ticamente el rango exacto de datos en una columna espec�fica
    ' de una hoja de c�lculo, identificando la primera y �ltima fila que contienen
    ' informaci�n. Funci�n esencial para determinar l�mites de procesamiento
    ' despu�s de la importaci�n de datos.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim rngBusqueda As Range
    Dim lngColumna As Long
    
    ' Obtener n�mero de columna
    lngColumna = Range(vColumnaInicial_Importacion & "1").Column
    
    ' Configurar rango de b�squeda
    Set rngBusqueda = ws.Columns(lngColumna)
    
    With rngBusqueda
        ' Encontrar primera celda con datos
        Set rngBusqueda = .Find(What:="*", _
                               After:=.Cells(.Cells.Count), _
                               LookIn:=xlFormulas, _
                               LookAt:=xlPart, _
                               SearchOrder:=xlByRows, _
                               SearchDirection:=xlNext)
        
        If Not rngBusqueda Is Nothing Then
            lngLineaInicial = rngBusqueda.Row
            
            ' Encontrar �ltima celda con datos
            Set rngBusqueda = .Find(What:="*", _
                                   After:=.Cells(1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlPrevious)
            
            lngLineaFinal = rngBusqueda.Row
            fun804_DetectarRangoDatos = True
        Else
            lngLineaInicial = 0
            lngLineaFinal = 0
            fun804_DetectarRangoDatos = False
        End If
    End With
    Exit Function
    
GestorErrores:
    lngLineaInicial = 0
    lngLineaFinal = 0
    fun804_DetectarRangoDatos = False
End Function

Public Function fun801_VerificarExistenciaHoja(wb As Workbook, nombreHoja As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 801: VERIFICAR EXISTENCIA DE HOJA
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Verifica si una hoja existe en el libro especificado
    ' Par�metros: wb (Workbook), nombreHoja (String)
    ' Retorna: Boolean (True si existe, False si no existe)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim lineaError As Long
    
    lineaError = 200
    fun801_VerificarExistenciaHoja = False
    
    ' Verificar par�metros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Exit Function
    End If
    
    lineaError = 210
    
    ' Recorrer todas las hojas del libro (m�todo compatible con Excel 97)
    For i = 1 To wb.Worksheets.Count
        If UCase(wb.Worksheets(i).Name) = UCase(nombreHoja) Then
            fun801_VerificarExistenciaHoja = True
            Exit For
        End If
    Next i
    
    lineaError = 220
    
    Exit Function
    
ErrorHandler:
    fun801_VerificarExistenciaHoja = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun801_VerificarExistenciaHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Sub fun804_LimpiarContenidoHoja(ws As Worksheet)
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 804: LIMPIAR CONTENIDO DE HOJA
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Limpia todo el contenido de una hoja espec�fica
    ' Par�metros: ws (Worksheet)
    ' Retorna: Nada (Sub procedure)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 500
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        Exit Sub
    End If
    
    lineaError = 510
    
    ' Verificar que la hoja no est� protegida
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    lineaError = 520
    
    ' Limpiar todo el contenido de la hoja (m�todo compatible con todas las versiones)
    ws.Cells.Clear
    
    lineaError = 530
    
    Exit Sub
    
ErrorHandler:
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun804_LimpiarContenidoHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub

Public Function fun805_DetectarUseSystemSeparators() As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 805: DETECTAR USE SYSTEM SEPARATORS
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta si Excel est� usando separadores del sistema
    ' Par�metros: Ninguno
    ' Retorna: String ("True" o "False")
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variable para almacenar el resultado
    Dim resultado As String
    Dim lineaError As Long
    
    lineaError = 600
    
    ' Detectar configuraci�n actual de Use System Separators
    ' Usar compilaci�n condicional para compatibilidad con versiones
    
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lineaError = 610
        If Application.UseSystemSeparators Then
            resultado = "True"
        Else
            resultado = "False"
        End If
    #Else
        ' Excel 97, 2003 y anteriores
        lineaError = 620
        resultado = fun809_DetectarUseSystemSeparatorsLegacy()
    #End If
    
    lineaError = 630
    
    fun805_DetectarUseSystemSeparators = resultado
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, intentar m�todo alternativo
    fun805_DetectarUseSystemSeparators = fun809_DetectarUseSystemSeparatorsLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun805_DetectarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun806_DetectarDecimalSeparator() As String

    ' =============================================================================
    ' FUNCI�N AUXILIAR 806: DETECTAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta el separador decimal actual de Excel
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador decimal)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    
    ' Detectar separador decimal actual (compatible con todas las versiones)
    fun806_DetectarDecimalSeparator = Application.DecimalSeparator
    
    lineaError = 710
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar m�todo alternativo
    fun806_DetectarDecimalSeparator = fun810_DetectarDecimalSeparatorLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun806_DetectarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun807_DetectarThousandsSeparator() As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 807: DETECTAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta el separador de miles actual de Excel
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador de miles)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    
    ' Detectar separador de miles actual (compatible con todas las versiones)
    fun807_DetectarThousandsSeparator = Application.ThousandsSeparator
    
    lineaError = 810
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar m�todo alternativo
    fun807_DetectarThousandsSeparator = fun811_DetectarThousandsSeparatorLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun807_DetectarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun809_DetectarUseSystemSeparatorsLegacy() As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 809: DETECTAR USE SYSTEM SEPARATORS (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar Use System Separators en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String ("True" o "False")
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para comparaci�n
    Dim separadorSistema As String
    Dim separadorExcel As String
    Dim lineaError As Long
    
    lineaError = 1000
    
    ' Obtener separador decimal del sistema (Windows) ' M�todo compatible con Excel 97 y 2003
    separadorSistema = Mid(CStr(1.1), 2, 1)
    
    lineaError = 1010
    
    ' Obtener separador decimal de Excel
    separadorExcel = Application.DecimalSeparator
    
    lineaError = 1020
    
    ' Si coinciden, probablemente Use System Separators est� activado
    If separadorSistema = separadorExcel Then
        fun809_DetectarUseSystemSeparatorsLegacy = "True"
    Else
        fun809_DetectarUseSystemSeparatorsLegacy = "False"
    End If
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir False por defecto
    fun809_DetectarUseSystemSeparatorsLegacy = "False"
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun809_DetectarUseSystemSeparatorsLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun810_DetectarDecimalSeparatorLegacy() As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 810: DETECTAR DECIMAL SEPARATOR (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar separador decimal en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador decimal)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detecci�n
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1100
    
    ' M�todo alternativo: formatear un n�mero y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = CStr(1.1)
    
    lineaError = 1110
    
    ' El separador decimal es el segundo car�cter en el formato est�ndar
    If Len(numeroFormateado) >= 2 Then
        fun810_DetectarDecimalSeparatorLegacy = Mid(numeroFormateado, 2, 1)
    Else
        ' Fallback: asumir punto por defecto
        fun810_DetectarDecimalSeparatorLegacy = "."
    End If
    
    lineaError = 1120
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir punto por defecto
    fun810_DetectarDecimalSeparatorLegacy = "."
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun810_DetectarDecimalSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Sub fun804_Aplicar_Formato_Inventario_Fila(vHojaInventario As Worksheet, vFila As Integer, vEsVisible As Boolean)

    ' =============================================================================
    ' FUNCION AUXILIAR: fun804_Aplicar_Formato_Inventario_Fila
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Aplica formato a una fila del inventario segun visibilidad
    ' PARAMETROS: vHojaInventario (Worksheet), vFila (Integer), vEsVisible (Boolean)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vRangoFila As Range
    
    ' Definir rango de la fila (columnas 2 a 4)
    Set vRangoFila = vHojaInventario.Range("B" & vFila & ":D" & vFila)
    
    If vEsVisible Then
        ' Fila visible: sin color de fondo
        vRangoFila.Interior.ColorIndex = xlNone
        vHojaInventario.Cells(vFila, 4).Value = ">> visible <<"
    Else
        ' Fila oculta: fondo gris medio
        vRangoFila.Interior.Color = RGB(128, 128, 128)
        vHojaInventario.Cells(vFila, 4).Value = "OCULTA"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

Public Function fun805_Es_Hoja_Protegida(vNombreHoja As String) As Boolean
    
    ' =============================================================================
    ' FUNCION AUXILIAR: fun805_Es_Hoja_Protegida
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Verifica si una hoja esta en la lista de hojas protegidas
    ' PARAMETROS: vNombreHoja (String)
    ' RETORNO: Boolean (True=protegida, False=no protegida)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vHojasProtegidas(1 To 6) As String
    Dim i As Integer
    
    ' Lista de hojas protegidas
    vHojasProtegidas(1) = "00_Ejecutar_Procesos"
    vHojasProtegidas(2) = "01_Inventario"
    vHojasProtegidas(3) = "05_Username"
    vHojasProtegidas(4) = "06_Delimitadores_Originales"
    vHojasProtegidas(5) = "09_Report_PL"
    vHojasProtegidas(6) = "10_Report_PL_AH"
    
    fun805_Es_Hoja_Protegida = False
    
    For i = 1 To 6
        If StrComp(vNombreHoja, vHojasProtegidas(i), vbTextCompare) = 0 Then
            fun805_Es_Hoja_Protegida = True
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    fun805_Es_Hoja_Protegida = False
    
End Function

Public Sub fun806_Eliminar_Hoja_Segura(vNombreHoja As String)
    
    ' =============================================================================
    ' SUB AUXILIAR: fun806_Eliminar_Hoja_Segura
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Elimina una hoja de forma segura con control de errores
    ' PARAMETROS: vNombreHoja (String)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vAlertas As Boolean
    
    ' Desactivar alertas para evitar confirmaciones
    vAlertas = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    ' Eliminar la hoja
    ThisWorkbook.Worksheets(vNombreHoja).Delete
    
    ' Restaurar alertas
    Application.DisplayAlerts = vAlertas
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = vAlertas
    ' No mostrar error, simplemente continuar
    
End Sub


