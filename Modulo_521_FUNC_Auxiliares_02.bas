Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_02"

Option Explicit

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
    ' FUNCI�N AUXILIAR 811: DETECTAR THOUSANDS SEPARATOR (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar separador de miles en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador de miles)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detecci�n
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1200
    
    ' M�todo alternativo: formatear un n�mero grande y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = Format(1000, "#,##0")
    
    lineaError = 1210
    
    ' El separador de miles es el segundo car�cter en n�meros de 4 d�gitos
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
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun811_DetectarThousandsSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun802_CrearHojaDelimitadores(wb As Workbook, nombreHoja As String) As Worksheet

    ' =============================================================================
    ' FUNCI�N AUXILIAR 802: CREAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Crea una nueva hoja con el nombre especificado y la deja visible
    ' Par�metros: wb (Workbook), nombreHoja (String)
    ' Retorna: Worksheet (referencia a la hoja creada, Nothing si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lineaError As Long
    
    lineaError = 300
    
    ' Verificar par�metros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Exit Function
    End If
    
    lineaError = 310
    
    ' Verificar que el libro no est� protegido (importante para entornos cloud)
    If wb.ProtectStructure Then
        Set fun802_CrearHojaDelimitadores = Nothing
        Debug.Print "ERROR: No se puede crear hoja, libro protegido - Funci�n: fun802_CrearHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 320
    
    ' Crear nueva hoja al final del libro (m�todo compatible con todas las versiones)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    lineaError = 330
    
    ' Asignar nombre a la hoja
    ws.Name = nombreHoja
    
    lineaError = 340
    
    ' Asegurar que la hoja est� visible
    ws.Visible = xlSheetVisible
    
    lineaError = 350
    
    ' Configuraci�n adicional para compatibilidad con entornos cloud
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    ' Retornar referencia a la hoja creada
    Set fun802_CrearHojaDelimitadores = ws
    
    lineaError = 360
    
    Exit Function
    
ErrorHandler:
    Set fun802_CrearHojaDelimitadores = Nothing
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun802_CrearHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun803_HacerHojaVisible(ws As Worksheet) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 803: HACER HOJA VISIBLE
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Verifica la visibilidad de una hoja y la hace visible si est� oculta
    ' Par�metros: ws (Worksheet)
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 400
    fun803_HacerHojaVisible = True
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        fun803_HacerHojaVisible = False
        Exit Function
    End If
    
    lineaError = 410
    
    ' Verificar que el libro permite cambiar visibilidad (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede cambiar visibilidad, libro protegido - Funci�n: fun803_HacerHojaVisible - " & Now()
        Exit Function
    End If
    
    lineaError = 420
    
    ' Verificar el estado actual de visibilidad y actuar seg�n corresponda
    Select Case ws.Visible
        Case xlSheetVisible
            ' La hoja ya est� visible, no hacer nada
            Debug.Print "INFO: Hoja " & ws.Name & " ya est� visible - Funci�n: fun803_HacerHojaVisible - " & Now()
            
        Case xlSheetHidden, xlSheetVeryHidden
            ' La hoja est� oculta, hacerla visible
            ws.Visible = xlSheetVisible
            Debug.Print "INFO: Hoja " & ws.Name & " se hizo visible - Funci�n: fun803_HacerHojaVisible - " & Now()
            
        Case Else
            ' Estado desconocido, forzar visibilidad
            ws.Visible = xlSheetVisible
            Debug.Print "INFO: Hoja " & ws.Name & " visibilidad forzada - Funci�n: fun803_HacerHojaVisible - " & Now()
    End Select
    
    lineaError = 430
    
    Exit Function
    
ErrorHandler:
    fun803_HacerHojaVisible = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun803_HacerHojaVisible" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun804_ConvertirValorACadena(valor As Variant) As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 804: CONVERTIR VALOR A CADENA
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Convierte un valor de celda a cadena de texto de forma segura
    ' Par�metros: valor (Variant)
    ' Retorna: String (valor convertido o cadena vac�a si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    Dim resultado As String
    
    lineaError = 500
    
    ' Verificar si el valor es Nothing o Empty
    If IsEmpty(valor) Or IsNull(valor) Then
        resultado = ""
    ElseIf IsError(valor) Then
        resultado = ""
    Else
        ' Convertir a cadena
        resultado = CStr(valor)
        ' Eliminar espacios en blanco al inicio y final
        resultado = Trim(resultado)
    End If
    
    lineaError = 510
    
    fun804_ConvertirValorACadena = resultado
    
    Exit Function
    
ErrorHandler:
    fun804_ConvertirValorACadena = ""
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun804_ConvertirValorACadena" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun805_ValidarValoresOriginales() As Boolean

    ' =============================================================================
    ' FUNCI�N AUXILIAR 805: VALIDAR VALORES ORIGINALES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Valida que los valores originales le�dos sean v�lidos para restaurar
    ' Par�metros: Ninguno (usa variables globales)
    ' Retorna: Boolean (True si v�lidos, False si no v�lidos)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    Dim esValido As Boolean
    
    lineaError = 600
    esValido = True
    
    ' Validar Use System Separators (debe ser "True" o "False")
    If vExcel_UseSystemSeparators_ValorOriginal <> "True" And vExcel_UseSystemSeparators_ValorOriginal <> "False" Then
        If vExcel_UseSystemSeparators_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inv�lido para Use System Separators: '" & vExcel_UseSystemSeparators_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 610
    
    ' Validar Decimal Separator (debe ser un solo car�cter)
    If Len(vExcel_DecimalSeparator_ValorOriginal) <> 1 Then
        If vExcel_DecimalSeparator_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inv�lido para Decimal Separator: '" & vExcel_DecimalSeparator_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 620
    
    ' Validar Thousands Separator (debe ser un solo car�cter)
    If Len(vExcel_ThousandsSeparator_ValorOriginal) <> 1 Then
        If vExcel_ThousandsSeparator_ValorOriginal <> "" Then
            Debug.Print "ADVERTENCIA: Valor inv�lido para Thousands Separator: '" & vExcel_ThousandsSeparator_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
        End If
        esValido = False
    End If
    
    lineaError = 630
    
    fun805_ValidarValoresOriginales = esValido
    
    ' Log de valores validados
    If esValido Then
        Debug.Print "INFO: Valores v�lidos para restaurar - UseSystem:" & vExcel_UseSystemSeparators_ValorOriginal & " Decimal:'" & vExcel_DecimalSeparator_ValorOriginal & "' Thousands:'" & vExcel_ThousandsSeparator_ValorOriginal & "' - Funci�n: fun805_ValidarValoresOriginales - " & Now()
    End If
    
    Exit Function
    
ErrorHandler:
    fun805_ValidarValoresOriginales = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun805_ValidarValoresOriginales" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun806_RestaurarUseSystemSeparators(valorOriginal As String) As Boolean

    ' =============================================================================
    ' FUNCI�N AUXILIAR 806: RESTAURAR USE SYSTEM SEPARATORS
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura la configuraci�n de Use System Separators
    ' Par�metros: valorOriginal (String) - "True" o "False"
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    fun806_RestaurarUseSystemSeparators = True
    
    ' Verificar que el valor sea v�lido
    If valorOriginal <> "True" And valorOriginal <> "False" Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Use System Separators, valor inv�lido: '" & valorOriginal & "' - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        fun806_RestaurarUseSystemSeparators = False
        Exit Function
    End If
    
    lineaError = 710
    
    ' Usar compilaci�n condicional para compatibilidad con versiones
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lineaError = 720
        If valorOriginal = "True" Then
            Application.UseSystemSeparators = True
            Debug.Print "INFO: Use System Separators configurado a True - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        Else
            Application.UseSystemSeparators = False
            Debug.Print "INFO: Use System Separators configurado a False - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        End If
    #Else
        ' Excel 97, 2003 y anteriores
        lineaError = 730
        Debug.Print "ADVERTENCIA: Use System Separators no disponible en esta versi�n de Excel - Funci�n: fun806_RestaurarUseSystemSeparators - " & Now()
        ' En versiones antiguas, esta propiedad no existe, pero no es error
    #End If
    
    lineaError = 740
    
    Exit Function
    
ErrorHandler:
    fun806_RestaurarUseSystemSeparators = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun806_RestaurarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun807_RestaurarDecimalSeparator(valorOriginal As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 807: RESTAURAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura el separador decimal original
    ' Par�metros: valorOriginal (String) - car�cter del separador decimal
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    fun807_RestaurarDecimalSeparator = True
    
    ' Verificar que el valor sea v�lido (un solo car�cter)
    If Len(valorOriginal) <> 1 Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Decimal Separator, valor inv�lido: '" & valorOriginal & "' - Funci�n: fun807_RestaurarDecimalSeparator - " & Now()
        fun807_RestaurarDecimalSeparator = False
        Exit Function
    End If
    
    lineaError = 810
    
    ' Restaurar separador decimal (compatible con todas las versiones)
    Application.DecimalSeparator = valorOriginal
    Debug.Print "INFO: Decimal Separator restaurado a: '" & valorOriginal & "' - Funci�n: fun807_RestaurarDecimalSeparator - " & Now()
    
    lineaError = 820
    
    Exit Function
    
ErrorHandler:
    fun807_RestaurarDecimalSeparator = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun807_RestaurarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun808_RestaurarThousandsSeparator(valorOriginal As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 808: RESTAURAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura el separador de miles original
    ' Par�metros: valorOriginal (String) - car�cter del separador de miles
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 900
    fun808_RestaurarThousandsSeparator = True
    
    ' Verificar que el valor sea v�lido (un solo car�cter)
    If Len(valorOriginal) <> 1 Then
        Debug.Print "ADVERTENCIA: No se puede restaurar Thousands Separator, valor inv�lido: '" & valorOriginal & "' - Funci�n: fun808_RestaurarThousandsSeparator - " & Now()
        fun808_RestaurarThousandsSeparator = False
        Exit Function
    End If
    
    lineaError = 910
    
    ' Restaurar separador de miles (compatible con todas las versiones)
    Application.ThousandsSeparator = valorOriginal
    Debug.Print "INFO: Thousands Separator restaurado a: '" & valorOriginal & "' - Funci�n: fun808_RestaurarThousandsSeparator - " & Now()
    
    lineaError = 920
    
    Exit Function
    
ErrorHandler:
    fun808_RestaurarThousandsSeparator = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun808_RestaurarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO valorOriginal: " & valorOriginal & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function fun809_OcultarHojaDelimitadores(ws As Worksheet) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 809: OCULTAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Oculta la hoja de delimitadores si est� habilitada la opci�n
    ' Par�metros: ws (Worksheet)
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 1000
    fun809_OcultarHojaDelimitadores = True
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        fun809_OcultarHojaDelimitadores = False
        Exit Function
    End If
    
    lineaError = 1010
    
    ' Verificar que el libro permite ocultar hojas (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede ocultar hoja, libro protegido - Funci�n: fun809_OcultarHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 1020
    
    ' Ocultar la hoja (compatible con todas las versiones de Excel)
    ws.Visible = xlSheetHidden
    Debug.Print "INFO: Hoja " & ws.Name & " ocultada - Funci�n: fun809_OcultarHojaDelimitadores - " & Now()
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    fun809_OcultarHojaDelimitadores = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun809_OcultarHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun802_VerificarCompatibilidad() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun802_VerificarCompatibilidad
    ' PROP�SITO: Verifica compatibilidad con diferentes versiones de Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = compatible, False = no compatible)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun802
    
    Dim strVersionExcel As String
    Dim dblVersionNumero As Double
    
    ' Obtener versi�n de Excel
    strVersionExcel = Application.Version
    dblVersionNumero = CDbl(strVersionExcel)
    
    ' Verificar compatibilidad (Excel 97 = 8.0, 2003 = 11.0, 365 = 16.0+)
    If dblVersionNumero >= 8# Then
        fun802_VerificarCompatibilidad = True
    Else
        fun802_VerificarCompatibilidad = False
    End If
    
    Exit Function

ErrorHandler_fun802:
    ' En caso de error, asumir compatibilidad
    fun802_VerificarCompatibilidad = True
End Function

Public Sub fun803_ObtenerConfiguracionActual(ByRef strDecimalAnterior As String, ByRef strMilesAnterior As String)
    ' =============================================================================
    ' FUNCI�N: fun803_ObtenerConfiguracionActual
    ' PROP�SITO: Obtiene la configuraci�n actual de delimitadores
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error GoTo ErrorHandler_fun803
    
    ' Obtener delimitador decimal actual
    strDecimalAnterior = Application.International(xlDecimalSeparator)
    
    ' Obtener delimitador de miles actual
    strMilesAnterior = Application.International(xlThousandsSeparator)
    
    Exit Sub

ErrorHandler_fun803:
    ' En caso de error, usar valores por defecto
    strDecimalAnterior = "."
    strMilesAnterior = ","
End Sub

Public Function fun804_AplicarNuevosDelimitadores() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun804_AplicarNuevosDelimitadores
    ' PROP�SITO: Aplica los nuevos delimitadores al sistema
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = �xito, False = error)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun804
    
    ' Aplicar nuevo delimitador decimal
    Application.DecimalSeparator = vDelimitadorDecimal_HFM
    
    ' Aplicar nuevo delimitador de miles
    Application.ThousandsSeparator = vDelimitadorMiles_HFM
    
    ' Forzar que Excel use los delimitadores del sistema
    Application.UseSystemSeparators = False
    
    ' Actualizar la pantalla
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    fun804_AplicarNuevosDelimitadores = True
    Exit Function

ErrorHandler_fun804:
    fun804_AplicarNuevosDelimitadores = False
End Function

Public Function fun805_VerificarAplicacionDelimitadores() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun805_VerificarAplicacionDelimitadores
    ' PROP�SITO: Verifica que los delimitadores se aplicaron correctamente
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = aplicados correctamente, False = error)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun805
    
    Dim strDecimalActual As String
    Dim strMilesActual As String
    
    ' Obtener delimitadores actuales
    strDecimalActual = Application.DecimalSeparator
    strMilesActual = Application.ThousandsSeparator
    
    ' Verificar que coinciden con los deseados
    If strDecimalActual = vDelimitadorDecimal_HFM And strMilesActual = vDelimitadorMiles_HFM Then
        fun805_VerificarAplicacionDelimitadores = True
    Else
        fun805_VerificarAplicacionDelimitadores = False
    End If
    
    Exit Function

ErrorHandler_fun805:
    fun805_VerificarAplicacionDelimitadores = False
End Function

Public Sub fun806_RestaurarConfiguracion(ByVal strDecimalAnterior As String, ByVal strMilesAnterior As String)
    ' =============================================================================
    ' FUNCI�N: fun806_RestaurarConfiguracion
    ' PROP�SITO: Restaura la configuraci�n anterior en caso de error
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error Resume Next
    
    ' Restaurar delimitador decimal anterior
    Application.DecimalSeparator = strDecimalAnterior
    
    ' Restaurar delimitador de miles anterior
    Application.ThousandsSeparator = strMilesAnterior
    
    ' Restaurar uso de separadores del sistema
    Application.UseSystemSeparators = True
    
    On Error GoTo 0
End Sub

Public Sub fun807_MostrarErrorDetallado(ByVal strFuncion As String, ByVal strTipoError As String, _
                                        ByVal lngLinea As Long, ByVal lngNumeroError As Long, _
                                        ByVal strDescripcionError As String)
    
    ' =============================================================================
    ' FUNCI�N: fun807_MostrarErrorDetallado
    ' PROP�SITO: Muestra informaci�n detallada del error ocurrido
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    Dim strMensajeError As String
    
    ' Construir mensaje de error detallado
    strMensajeError = "ERROR EN FUNCI�N DE DELIMITADORES" & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Funci�n: " & strFuncion & vbCrLf
    strMensajeError = strMensajeError & "Tipo de Error: " & strTipoError & vbCrLf
    strMensajeError = strMensajeError & "L�nea Aproximada: " & CStr(lngLinea) & vbCrLf
    strMensajeError = strMensajeError & "N�mero de Error VBA: " & CStr(lngNumeroError) & vbCrLf
    strMensajeError = strMensajeError & "Descripci�n: " & strDescripcionError & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Fecha/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Mostrar mensaje de error
    MsgBox strMensajeError, vbCritical, "Error en F004_Forzar_Delimitadores_en_Excel"
    
End Sub



Public Function fun803_ObtenerCarpetaExcelActual() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta donde est� ubicado el archivo Excel actual
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener ruta completa del archivo actual
    If ThisWorkbook.Path <> "" Then
        strCarpeta = ThisWorkbook.Path
    ElseIf ActiveWorkbook.Path <> "" Then
        strCarpeta = ActiveWorkbook.Path
    Else
        strCarpeta = ""
    End If
    
    fun803_ObtenerCarpetaExcelActual = strCarpeta
    Exit Function
    
ErrorHandler:
    fun803_ObtenerCarpetaExcelActual = ""
End Function

Public Function fun804_ObtenerCarpetaTemp() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TEMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TEMP (compatible con Excel 97+)
    strCarpeta = Environ("TEMP")
    
    fun804_ObtenerCarpetaTemp = strCarpeta
    Exit Function
    
ErrorHandler:
    fun804_ObtenerCarpetaTemp = ""
End Function

Public Function fun805_ObtenerCarpetaTmp() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TMP (compatible con Excel 97+)
    strCarpeta = Environ("TMP")
    
    fun805_ObtenerCarpetaTmp = strCarpeta
    Exit Function
    
ErrorHandler:
    fun805_ObtenerCarpetaTmp = ""
End Function

Public Function fun806_ObtenerCarpetaUserProfile() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %USERPROFILE%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno USERPROFILE (compatible con Excel 97+)
    strCarpeta = Environ("USERPROFILE")
    
    fun806_ObtenerCarpetaUserProfile = strCarpeta
    Exit Function
    
ErrorHandler:
    fun806_ObtenerCarpetaUserProfile = ""
End Function

Public Function fun807_ValidarCarpeta(ByVal strCarpeta As String) As Boolean

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Valida si una carpeta existe y es accesible
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim objFSO As Object
    Dim blnResultado As Boolean
    
    blnResultado = False
    
    ' Verificar que la carpeta no est� vac�a
    If Len(Trim(strCarpeta)) = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Crear objeto FileSystemObject (compatible con Excel 97+)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si la carpeta existe y es accesible
    If objFSO.FolderExists(strCarpeta) Then
        blnResultado = True
    End If
    
    Set objFSO = Nothing
    fun807_ValidarCarpeta = blnResultado
    Exit Function
    
ErrorHandler:
    Set objFSO = Nothing
    fun807_ValidarCarpeta = False
End Function



Public Function F005_Proteger_Hoja_Contra_Escritura(ByVal vNombreHoja As String) As Boolean
    '******************************************************************************
    ' Funci�n: F005_Proteger_Hoja_Contra_Escritura
    ' Fecha y Hora de Creaci�n: 2025-06-09 12:53:08 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Protege una hoja espec�fica contra escritura, aplicando protecci�n est�ndar
    ' que impide modificaciones de contenido manteniendo la navegaci�n disponible.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables de control y optimizaci�n
    ' 2. Configurar optimizaciones de rendimiento (pantalla, c�lculos)
    ' 3. Validar existencia de la hoja especificada usando fun801_VerificarExistenciaHoja
    ' 4. Obtener referencia a la hoja de trabajo
    ' 5. Verificar estado actual de protecci�n
    ' 6. Aplicar protecci�n con configuraci�n est�ndar
    ' 7. Validar que la protecci�n se aplic� correctamente
    ' 8. Restaurar configuraciones de optimizaci�n
    ' 9. Registrar resultado en log del sistema
    ' 10. Manejo exhaustivo de errores con informaci�n detallada
    '
    ' Par�metros:
    ' - vNombreHoja (String): Nombre de la hoja a proteger
    '
    ' Valor de Retorno:
    ' - Boolean: True si la protecci�n se aplic� exitosamente, False si error
    '
    ' Compatibilidad: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    ' Versi�n: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimizaci�n
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim blnHojaExiste As Boolean
    Dim blnYaProtegida As Boolean
    
    ' Inicializaci�n
    strFuncion = "F005_Proteger_Hoja_Contra_Escritura"
    F005_Proteger_Hoja_Contra_Escritura = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializar variables de control y optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando protecci�n de hoja", False, "", vNombreHoja
    
    ' Almacenar configuraciones originales para restaurar despu�s
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Configurar optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 60
    ' Desactivar actualizaci�n de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar c�lculo autom�tico para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 3. Validar existencia de la hoja especificada
    '--------------------------------------------------------------------------
    lngLineaError = 70
    ' Obtener referencia al libro actual
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    ' Verificar que tenemos una referencia v�lida al libro
    If wb Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, _
            "No se pudo obtener referencia al libro de trabajo"
    End If
    
    ' Verificar existencia de la hoja usando funci�n auxiliar existente del proyecto
    blnHojaExiste = fun801_VerificarExistenciaHoja(wb, vNombreHoja)
    
    If Not blnHojaExiste Then
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, _
            "La hoja especificada no existe: " & vNombreHoja
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Obtener referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 80
    Set ws = wb.Worksheets(vNombreHoja)
    
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vNombreHoja
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Verificar estado actual de protecci�n
    '--------------------------------------------------------------------------
    lngLineaError = 90
    ' Verificar si la hoja ya est� protegida
    blnYaProtegida = ws.ProtectContents
    
    If blnYaProtegida Then
        ' La hoja ya est� protegida, registrar en log pero no es error
        fun801_LogMessage "La hoja ya estaba protegida", False, "", vNombreHoja
        F005_Proteger_Hoja_Contra_Escritura = True
        GoTo RestaurarConfiguracion
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Aplicar protecci�n con configuraci�n est�ndar
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Aplicando protecci�n a la hoja", False, "", vNombreHoja
    
    ' Aplicar protecci�n con configuraci�n compatible con todas las versiones de Excel
    ' Par�metros optimizados para compatibilidad Excel 97-365
    On Error Resume Next
    
    ' M�todo compatible con Excel 97+
    ws.Protect _
        Password:="", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        UserInterfaceOnly:=False, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, _
        AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, _
        AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, _
        AllowDeletingRows:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False, _
        AllowUsingPivotTables:=False
    
    ' Verificar si hubo error en la protecci�n
    If Err.Number <> 0 Then
        ' Si falla el m�todo avanzado, usar m�todo b�sico compatible
        Err.Clear
        ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
        
        ' Verificar nuevamente si hubo error
        If Err.Number <> 0 Then
            On Error GoTo GestorErrores
            Err.Raise ERROR_BASE_IMPORT + 4, strFuncion, _
                "Error al aplicar protecci�n a la hoja: " & Err.Description
        End If
    End If
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 7. Validar que la protecci�n se aplic� correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 110
    ' Verificar que la hoja ahora est� protegida
    If Not ws.ProtectContents Then
        Err.Raise ERROR_BASE_IMPORT + 5, strFuncion, _
            "La protecci�n no se aplic� correctamente a la hoja: " & vNombreHoja
    End If
    
    fun801_LogMessage "Protecci�n aplicada exitosamente", False, "", vNombreHoja
    
    '--------------------------------------------------------------------------
    ' 8. Proceso completado exitosamente
    '--------------------------------------------------------------------------
    lngLineaError = 120
    F005_Proteger_Hoja_Contra_Escritura = True
    
RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 9. Restaurar configuraciones de optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 130
    ' Restaurar configuraci�n original de actualizaci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Restaurar configuraci�n original de c�lculo
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    ' Restaurar configuraci�n original de eventos
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set ws = Nothing
    Set wb = Nothing
    
    fun801_LogMessage "Protecci�n de hoja completada", False, "", vNombreHoja
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 10. Manejo exhaustivo de errores con informaci�n detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja: " & vNombreHoja & vbCrLf & _
                      "Fecha y Hora: " & Now()
    
    ' Registrar error en log del sistema
    fun801_LogMessage strMensajeError, True, "", vNombreHoja
    
    ' Mostrar error al usuario (opcional, comentar si no se desea)
    MsgBox strMensajeError, vbCritical, "Error en Protecci�n de Hoja"
    
    ' Log del error para debugging
    Debug.Print strMensajeError
    
    ' Restaurar configuraciones en caso de error
    On Error Resume Next
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set ws = Nothing
    Set wb = Nothing
    
    ' Retornar False para indicar error
    F005_Proteger_Hoja_Contra_Escritura = False
End Function


