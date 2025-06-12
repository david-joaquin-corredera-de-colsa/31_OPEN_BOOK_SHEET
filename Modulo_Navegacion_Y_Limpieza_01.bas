Attribute VB_Name = "Modulo_Navegacion_Y_Limpieza_01"


Option Explicit

Option Explicit

' Variables globales del modulo
Public vHojaInicial As String

Public Function F011_Limpieza_Hojas_Historicas() As Boolean

    ' =============================================================================
    ' FUNCION: F011_Limpieza_Hojas_Historicas
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Funcion para limpiar hojas historicas segun criterios especificos
    ' PARAMETROS: Ninguno
    ' RETORNO: Integer (0=exito, >0=error)
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables de control
    ' 1.5. Configurar visibilidad de hojas especificas
    ' 2. Primera pasada - recopilar hojas Import_Envio_
    ' 3. Ordenar hojas Import_Envio_ lexicograficamente
    ' 4. Segunda pasada - aplicar reglas de limpieza especificas
    ' 5. Gestionar hojas Import_Envio_ con logica de ordenamiento lexicografico
    ' 6. Retornar codigo de resultado

    On Error GoTo ErrorHandler
    
    Dim vResultado As Integer
    Dim vContadorHojas As Integer
    Dim vNombreHoja As String
    Dim vLineaError As Integer
    Dim vTotalHojas As Integer
    Dim vHojasEnvio() As String
    Dim vContadorEnvio As Integer
    Dim vNumHojasEnvio As Integer
    Dim i As Integer, j As Integer
    Dim vTempNombre As String
    
    vLineaError = 10
    vContadorEnvio = 0
    vNumHojasEnvio = 0
    
    ' Paso 1: Inicializar variables de control
    vLineaError = 20
    vTotalHojas = ThisWorkbook.Worksheets.Count
    
    ' Paso 1.5: Configurar visibilidad de hojas especificas
    vLineaError = 25
    For vContadorHojas = 1 To vTotalHojas
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        Call fun801_Configurar_Visibilidad_Hojas_Especificas(vNombreHoja)
    Next vContadorHojas
    
    ' Redimensionar array para hojas Import_Envio_ (estimacion maxima)
    ReDim vHojasEnvio(1 To vTotalHojas)
    
    ' Paso 2: Primera pasada - recopilar hojas Import_Envio_
    vLineaError = 30
    For vContadorHojas = vTotalHojas To 1 Step -1
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        If Left(UCase(vNombreHoja), 13) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            vNumHojasEnvio = vNumHojasEnvio + 1
            vHojasEnvio(vNumHojasEnvio) = vNombreHoja
        End If
    Next vContadorHojas
    
    ' Paso 3: Ordenar hojas Import_Envio_ lexicograficamente (bubble sort compatible Excel 97)
    vLineaError = 40
    If vNumHojasEnvio > 1 Then
        For i = 1 To vNumHojasEnvio - 1
            For j = 1 To vNumHojasEnvio - i
                If StrComp(vHojasEnvio(j), vHojasEnvio(j + 1), vbTextCompare) < 0 Then
                    vTempNombre = vHojasEnvio(j)
                    vHojasEnvio(j) = vHojasEnvio(j + 1)
                    vHojasEnvio(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Paso 4: Segunda pasada - aplicar reglas de limpieza especificas
    vLineaError = 50
    For vContadorHojas = vTotalHojas To 1 Step -1
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        ' Para cada hoja aplicar reglas de limpieza especificas
        vLineaError = 60
        
        ' Regla: Hojas protegidas - no hacer nada
        If fun805_Es_Hoja_Protegida(vNombreHoja) Then
            ' No hacer nada con estas hojas
            
        ' Regla: Eliminar hojas Import_Working_
        ElseIf Left(UCase(vNombreHoja), 15) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_WORKING) Then
            vLineaError = 70
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Import_Comprob_
        ElseIf Left(UCase(vNombreHoja), 15) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_COMPROBACION) Then
            vLineaError = 80
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Import_ con longitud 22
        ElseIf Left(UCase(vNombreHoja), 7) = UCase(CONST_PREFIJO_HOJA_IMPORTACION) And Len(vNombreHoja) = 22 Then
            vLineaError = 90
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
            
        ' Regla: Eliminar hojas Del_Prev_Envio_
        ElseIf Left(vNombreHoja, 15) = "Del_Prev_Envio_" Then
            vLineaError = 100
            Call fun806_Eliminar_Hoja_Segura(vNombreHoja)
        
        ' Regla: Gestionar hojas Import_Envio_
        ElseIf Left(UCase(vNombreHoja), 13) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            vLineaError = 110
            Call fun807_Gestionar_Hoja_Envio(vNombreHoja, vHojasEnvio, vNumHojasEnvio)
        End If
    Next vContadorHojas
    
    ' Paso 6: Retornar codigo de resultado
    F011_Limpieza_Hojas_Historicas = True
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F011_Limpieza_Hojas_Historicas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Hoja procesando: " & vNombreHoja
    
    MsgBox vMensajeError, vbCritical, "Error F011_Limpieza_Hojas_Historicas"

    F011_Limpieza_Hojas_Historicas = False
    
End Function
Public Function Function_Return_Integer_to_Boolean(vInteger As Integer) As Boolean
    '******************************************************************************
    ' Módulo: Function_Return_Integer_to_Boolean
    ' Fecha y Hora de Creación: 2025-06-09 09:10:01 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    ' Descripción:
    ' Convierte un valor entero a un valor booleano siguiendo una lógica específica
    ' donde 0 se considera verdadero (True) y cualquier otro valor se considera
    ' falso (False). Esta función implementa una lógica inversa a la conversión
    ' booleana estándar de VBA.
    ' Parámetros:
    ' - vInteger (Integer): Valor entero a convertir a booleano
    ' Valor de Retorno:
    ' - Boolean: True si el valor de entrada es 0, False para cualquier otro valor
    ' Lógica de Conversión:
    ' - Input: 0 ? Output: True
    ' - Input: cualquier otro número ? Output: False
    ' Casos de Uso Típicos:
    ' - Validación de códigos de error (donde 0 indica éxito)
    ' - Conversión de flags numéricos a booleanos
    ' - Procesamiento de datos donde 0 representa un estado "activo" o "válido"
    '
    ' Ejemplos de Uso:
    ' Dim resultado As Boolean
    ' resultado = Function_Return_Integer_to_Boolean(0)    ' Devuelve True
    ' resultado = Function_Return_Integer_to_Boolean(1)    ' Devuelve False
    ' resultado = Function_Return_Integer_to_Boolean(-5)   ' Devuelve False
    ' Notas Importantes:
    ' ?? Esta función implementa una lógica inversa a la conversión booleana
    ' estándar de VBA, donde normalmente 0 equivale a False y cualquier valor
    ' diferente de 0 equivale a True.
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' Versión: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Inicialización
    strFuncion = "Function_Return_Integer_to_Boolean"
    Function_Return_Integer_to_Boolean = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' Lógica principal de conversión
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    If vInteger = 0 Then
        Function_Return_Integer_to_Boolean = True
    Else
        Function_Return_Integer_to_Boolean = False
    End If
    
    Exit Function

GestorErrores:
    ' Manejo de errores con información detallada
    Dim strMensajeError As String
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Valor de entrada: " & vInteger
    
    ' Log del error para debugging
    Debug.Print strMensajeError
    
    ' Retornar False en caso de error
    Function_Return_Integer_to_Boolean = False
End Function

Public Function F012_Inventariar_Hojas() As Integer
    ' =============================================================================
    ' FUNCION: F012_Inventariar_Hojas
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Funcion para crear inventario completo de todas las hojas del libro
    ' PARAMETROS: Ninguno
    ' RETORNO: Integer (0=exito, >0=error)
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Verificar existencia de hoja "01_Inventario"
    ' 2. Borrar contenido y formatos de la hoja "01_Inventario"
    ' 3. Crear encabezados en linea 2 con formato especifico
    ' 4. Recorrer todas las hojas y recopilar informacion completa
    ' 5. Crear enlaces (hyperlinks) para cada hoja
    ' 6. Aplicar formato segun visibilidad de cada hoja
    ' 7. Buscar fichero fuente en hoja "02_Log"
    ' 8. Ordenar el listado alfabeticamente
    ' 9. Asegurar visibilidad de hoja "01_Inventario"

    On Error GoTo ErrorHandler
    
    Dim vResultado As Integer
    Dim vLineaError As Integer
    Dim vHojaInventario As Worksheet
    Dim vContadorHojas As Integer
    Dim vFilaActual As Integer
    Dim vNombreHoja As String
    Dim vRangoOrdenar As Range
    Dim vEsVisible As Boolean
    Dim vFicheroFuente As String
    
    vResultado = 0
    vLineaError = 10
    vFilaActual = 3
    
    ' Paso 1: Verificar existencia de hoja "01_Inventario"
    vLineaError = 20
    Set vHojaInventario = fun808_Obtener_Hoja_Inventario()
    If vHojaInventario Is Nothing Then
        vResultado = 3001 ' Error: No se pudo acceder a hoja inventario
        GoTo ErrorHandler
    End If
    
    ' Paso 2: Borrar contenido y formatos de la hoja "01_Inventario"
    vLineaError = 30
    vHojaInventario.Cells.Clear
    vHojaInventario.Cells.ClearFormats
    
    ' Paso 3: Crear encabezados en linea 2
    vLineaError = 40
    vHojaInventario.Cells(2, 2).Value = "Nombre de la Hoja"
    vHojaInventario.Cells(2, 3).Value = "Link a la Hoja"
    vHojaInventario.Cells(2, 4).Value = "Visible/Oculta"
    vHojaInventario.Cells(2, 5).Value = "Fichero Fuente"
    
    ' Paso 3.1: Aplicar formato a encabezados
    vLineaError = 45
    Call fun803_Aplicar_Formato_Inventario_Encabezados(vHojaInventario)
    
    ' Paso 4: Recorrer todas las hojas y recopilar informacion completa
    vLineaError = 50
    For vContadorHojas = 1 To ThisWorkbook.Worksheets.Count
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        
        ' Paso 4.1: Escribir nombre de la hoja
        vLineaError = 60
        vHojaInventario.Cells(vFilaActual, 2).Value = vNombreHoja
        
        ' Paso 4.2: Crear enlaces (hyperlinks) para cada hoja
        vLineaError = 70
        Call fun809_Crear_Enlace_Hoja(vHojaInventario, vFilaActual, 3, vNombreHoja)
        
        ' Paso 4.3: Determinar si la hoja es visible
        vLineaError = 80
        vEsVisible = (ThisWorkbook.Worksheets(vNombreHoja).Visible = xlSheetVisible)
        
        ' Paso 4.4: Aplicar formato segun visibilidad
        vLineaError = 90
        Call fun804_Aplicar_Formato_Inventario_Fila(vHojaInventario, vFilaActual, vEsVisible)
        
        ' Paso 4.5: Buscar fichero fuente en hoja "02_Log"
        vLineaError = 100
        vFicheroFuente = fun802_Buscar_Fichero_Fuente_En_Log(vNombreHoja)
        vHojaInventario.Cells(vFilaActual, 5).Value = vFicheroFuente
        
        vFilaActual = vFilaActual + 1
    Next vContadorHojas
    
    ' Paso 5: Ordenar el listado alfabeticamente (compatible Excel 97)
    vLineaError = 110
    If vFilaActual > 3 Then
        Set vRangoOrdenar = vHojaInventario.Range("B3:E" & (vFilaActual - 1))
        vRangoOrdenar.Sort Key1:=vHojaInventario.Range("B3"), Order1:=xlAscending, Header:=xlNo
    End If
    
    ' Paso 6: Asegurar visibilidad de hoja "01_Inventario"
    vLineaError = 120
    vHojaInventario.Visible = xlSheetVisible
    
    ' Ajustar columnas automaticamente
    vLineaError = 125
    vHojaInventario.Columns("B:E").AutoFit
    
    F012_Inventariar_Hojas = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en F012_Inventariar_Hojas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description
    
    MsgBox vMensajeError, vbCritical, "Error F012_Inventariar_Hojas"
    
    If vResultado = 0 Then
        vResultado = 3999 ' Error no especificado en inventario
    End If
    
    F012_Inventariar_Hojas = vResultado
    
End Function

Public Sub fun801_Configurar_Visibilidad_Hojas_Especificas(vNombreHoja As String)
    
    ' =============================================================================
    ' FUNCION AUXILIAR: fun801_Configurar_Visibilidad_Hojas_Especificas
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Configura la visibilidad de hojas especificas segun criterios
    ' PARAMETROS: vNombreHoja (String)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim vHojasVisibles(1 To 4) As String
    Dim vHojasOcultas(1 To 3) As String
    Dim i As Integer
    Dim vHoja As Worksheet
    
    ' Definir hojas que deben estar visibles
    vHojasVisibles(1) = "00_Ejecutar_Procesos"
    vHojasVisibles(2) = "01_Inventario"
    vHojasVisibles(3) = "02_Log"
    vHojasVisibles(4) = "09_Report_PL"
    
    ' Definir hojas que deben estar ocultas
    vHojasOcultas(1) = "05_Username"
    vHojasOcultas(2) = "06_Delimitadores_Originales"
    vHojasOcultas(3) = "10_Report_PL_AH"
    
    Set vHoja = ThisWorkbook.Worksheets(vNombreHoja)
    
    ' Verificar si debe estar visible
    For i = 1 To 4
        If StrComp(vNombreHoja, vHojasVisibles(i), vbTextCompare) = 0 Then
            vHoja.Visible = xlSheetVisible
            Exit Sub
        End If
    Next i
    
    ' Verificar si debe estar oculta
    For i = 1 To 3
        If StrComp(vNombreHoja, vHojasOcultas(i), vbTextCompare) = 0 Then
            vHoja.Visible = xlSheetHidden
            Exit Sub
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

' =============================================================================
' FUNCION AUXILIAR: fun802_Buscar_Fichero_Fuente_En_Log
' FECHA: 2025-06-03 15:18:26 UTC
' DESCRIPCION: Busca el fichero fuente de una hoja en el log
' PARAMETROS: vNombreHoja (String)
' RETORNO: String (nombre del fichero fuente o "")
' =============================================================================
Public Function fun802_Buscar_Fichero_Fuente_En_Log(vNombreHoja As String) As String
    
    On Error GoTo ErrorHandler
    
    Dim vHojaLog As Worksheet
    Dim vUltimaFila As Long
    Dim i As Long
    Dim vValorColumnaD As String
    Dim vValorColumnaE As String
    Dim vBuscarTexto As String
    
    fun802_Buscar_Fichero_Fuente_En_Log = ""
    
    ' Obtener referencia a la hoja "02_Log"
    Set vHojaLog = Nothing
    Set vHojaLog = ThisWorkbook.Worksheets("02_Log")
    If vHojaLog Is Nothing Then Exit Function
    
    ' Determinar texto a buscar basado en el nombre de la hoja
    If Left(UCase(vNombreHoja), 13) = UCase(CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
        vBuscarTexto = CONST_PREFIJO_HOJA_IMPORTACION & Right(vNombreHoja, 15)
    Else
        vBuscarTexto = vNombreHoja
    End If
    
    ' Obtener ultima fila con datos
    vUltimaFila = vHojaLog.Cells(vHojaLog.Rows.Count, "D").End(xlUp).Row
    
    ' Recorrer las filas buscando la coincidencia
    For i = 1 To vUltimaFila
        vValorColumnaD = CStr(vHojaLog.Cells(i, 4).Value)
        vValorColumnaE = CStr(vHojaLog.Cells(i, 5).Value)
        
        ' Verificar condiciones: columna D diferente de "NA" y contiene "\"
        ' y columna E igual al texto buscado
        If StrComp(vValorColumnaD, "NA", vbTextCompare) <> 0 And _
           InStr(vValorColumnaD, "\") > 0 And _
           StrComp(vValorColumnaE, vBuscarTexto, vbTextCompare) = 0 Then
            fun802_Buscar_Fichero_Fuente_En_Log = vValorColumnaD
            Exit Function
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    fun802_Buscar_Fichero_Fuente_En_Log = ""
    
End Function

Public Sub fun803_Aplicar_Formato_Inventario_Encabezados(vHojaInventario As Worksheet)
    
    ' =============================================================================
    ' FUNCION AUXILIAR: fun803_Aplicar_Formato_Inventario_Encabezados
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Aplica formato a los encabezados del inventario
    ' PARAMETROS: vHojaInventario (Worksheet)
    ' =============================================================================
    On Error GoTo ErrorHandler
    
    Dim vRangoEncabezados As Range
    
    ' Definir rango de encabezados (fila 2, columnas 2 a 5)
    Set vRangoEncabezados = vHojaInventario.Range("B2:E2")
    
    ' Aplicar formato de fondo negro
    vRangoEncabezados.Interior.Color = RGB(0, 0, 0)
    
    ' Aplicar formato de fuente blanca y negrita
    With vRangoEncabezados.Font
        .Color = RGB(255, 255, 255)
        .Bold = True
    End With
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

Public Sub fun807_Gestionar_Hoja_Envio(vNombreHoja As String, vHojasEnvio() As String, vNumHojasEnvio As Integer)
    
    ' =============================================================================
    ' SUB AUXILIAR: fun807_Gestionar_Hoja_Envio
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Gestiona visibilidad de hojas Import_Envio_ segun antiguedad
    ' PARAMETROS: vNombreHoja (String), vHojasEnvio (Array), vNumHojasEnvio (Integer)
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim vPosicion As Integer
    Dim vHoja As Worksheet
    
    ' Buscar posicion de la hoja en el array ordenado
    vPosicion = 0
    For i = 1 To vNumHojasEnvio
        If StrComp(vHojasEnvio(i), vNombreHoja, vbTextCompare) = 0 Then
            vPosicion = i
            Exit For
        End If
    Next i
    
    Set vHoja = ThisWorkbook.Worksheets(vNombreHoja)
    
    ' Si hay mas hojas que el limite y esta fuera del rango visible
    If vNumHojasEnvio > CONS_NUM_HOJAS_HCAS_VISIBLES_ENVIO And vPosicion > CONS_NUM_HOJAS_HCAS_VISIBLES_ENVIO Then
        vHoja.Visible = xlSheetHidden
    Else
        vHoja.Visible = xlSheetVisible
    End If
    
    Exit Sub
    
ErrorHandler:
    ' No mostrar error, simplemente continuar
    
End Sub

Public Function Ejecutar_Procesos_Inicio_Libro() As Boolean

    ' =============================================================================
    ' SUB PRINCIPAL: Ejecutar_Procesos_Inicio_Libro
    ' FECHA: 2025-06-03 15:18:26 UTC
    ' DESCRIPCION: Sub principal que ejecuta todos los procesos al abrir el libro
    ' PARAMETROS: Ninguno
    ' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
    ' =============================================================================
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Ejecutar F010_Abrir_Hoja_Inicial
    ' 2. Ejecutar F011_Limpieza_Hojas_Historicas
    ' 3. Ejecutar F012_Inventariar_Hojas
    ' 4. Evaluar resultados y manejar errores

    On Error GoTo ErrorHandler
    
    '***************************
    Dim vAbrirHojaInicial As Boolean
    Dim vLimpiarHojasHistoricas As Boolean
    Dim vInventariarHojas As Boolean
    '***************************
    Dim vLineaError As Integer
    '***************************
    Dim vHMS_Inicial As String
    Dim vHMS_Final As String
    Dim vDuracion As String
    Dim vTextoPasoActual As String
    '***************************
    
    vLineaError = 10
    '---------------------------------------------
    ' Paso 1: Ejecutar F010_Abrir_Hoja_Inicial
    '---------------------------------------------
    vLineaError = 20
    
    vHMS_Inicial = fun830_ObtenerFechaHoraActual
    '----
    vAbrirHojaInicial = F010_Abrir_Hoja_Inicial()
    '----
    vHMS_Final = fun830_ObtenerFechaHoraActual
    
    vDuracion = fun831_CalcularDuracionTarea(vHMS_Inicial, vHMS_Final)
    vTextoPasoActual = " abrir hoja inicial "
    'MsgBox "Paso: " & vTextoPasoActual & vbCrLf & vDuracion
    
    '---------------------------------------------
    ' Paso 2: Ejecutar F011_Limpieza_Hojas_Historicas
    '---------------------------------------------
    vLineaError = 30
    
    vHMS_Inicial = fun830_ObtenerFechaHoraActual
    '----
    vLimpiarHojasHistoricas = F011_Limpieza_Hojas_Historicas()
    '----
    vHMS_Final = fun830_ObtenerFechaHoraActual
    vDuracion = fun831_CalcularDuracionTarea(vHMS_Inicial, vHMS_Final)
    vTextoPasoActual = " limpieza de hojas historicas "
    'MsgBox "Paso: " & vTextoPasoActual & vbCrLf & vDuracion
    
    '---------------------------------------------
    ' Paso 3: Ejecutar F012_Inventariar_Hojas
    '---------------------------------------------
    vLineaError = 40
    
    vHMS_Inicial = fun830_ObtenerFechaHoraActual
    '----
    vInventariarHojas = F012_Inventariar_Hojas()
    '----
    vHMS_Final = fun830_ObtenerFechaHoraActual
    vDuracion = fun831_CalcularDuracionTarea(vHMS_Inicial, vHMS_Final)
    vTextoPasoActual = " limpieza inventariar hojas"
    'MsgBox "Paso: " & vTextoPasoActual & vbCrLf & vDuracion
    
    
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en Ejecutar_Procesos_Inicio_Libro" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description
    
    MsgBox vMensajeError, vbCritical, "Error Principal"
    
End Function


Public Function fun808_ObtenerCarpetaSistema(ByVal strTipoOrigen As String) As String
    
    '******************************************************************************
    ' FUNCION CONSOLIDADA: fun808_ObtenerCarpetaSistema
    ' FECHA Y HORA DE CREACION: 2025-06-12 15:19:14 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPOSITO:
    ' Función consolidada que obtiene carpetas del sistema desde múltiples fuentes:
    ' variables de entorno del sistema, propiedades de Excel, y directorios especiales.
    ' Reemplaza las funciones individuales fun803, fun804, fun805, fun806.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validación y normalización de parámetros de entrada
    ' 2. Identificación del método de obtención según el tipo especificado
    ' 3. Ejecución del método específico para cada tipo de origen
    ' 4. Manejo de casos especiales y fallbacks automáticos
    ' 5. Normalización y limpieza de la ruta obtenida
    ' 6. Validación opcional de existencia de la carpeta
    ' 7. Logging detallado de la operación realizada
    ' 8. Retorno del resultado final normalizado
    '
    ' PARAMETROS:
    ' - strTipoOrigen (String): Tipo de origen de la carpeta a obtener
    '   Valores soportados exactos (case-insensitive):
    '   * "TEMP" - Variable de entorno %TEMP% del sistema
    '   * "TMP" - Variable de entorno %TMP% del sistema
    '   * "USERPROFILE" - Variable de entorno %USERPROFILE% del sistema
    '   * "CURRENT_DIR" - Directorio de trabajo actual
    '   * "EXCEL_PATH_CURRENT_BOOK" - Carpeta del libro Excel actual
    '   * "EXCEL_PATH_TEMP" - Carpeta temporal de Excel
    '
    ' RETORNO: String - Ruta de la carpeta obtenida o cadena vacía si error
    '
    ' EJEMPLOS DE USO:
    ' Dim strCarpetaTemp As String
    ' strCarpetaTemp = fun808_ObtenerCarpetaSistema("TEMP")
    ' strCarpetaTemp = fun808_ObtenerCarpetaSistema("TMP")
    ' strCarpetaTemp = fun808_ObtenerCarpetaSistema("USERPROFILE")
    ' strCarpetaTemp = fun808_ObtenerCarpetaSistema("CURRENT_DIR")
    ' strCarpetaTemp = fun808_ObtenerCarpetaSistema("EXCEL_PATH_CURRENT_BOOK")
    ' strCarpetaTemp = fun808_ObtenerCarpetaSistema("EXCEL_PATH_TEMP")
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento
    Dim strTipoNormalizado As String
    Dim strRutaObtenida As String
    Dim strRutaLimpia As String
    Dim strMetodoUsado As String
    
    ' Variables para validación y manejo de objetos
    Dim blnValidarCarpeta As Boolean
    Dim objFSO As Object
    Dim intUltimaBarraDiagonal As Integer
    
    ' Inicialización
    strFuncion = "fun808_ObtenerCarpetaSistema"
    fun808_ObtenerCarpetaSistema = ""
    lngLineaError = 0
    blnValidarCarpeta = True
    strMetodoUsado = ""
    strRutaObtenida = ""
    strRutaLimpia = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validación y normalización de parámetros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    ' Verificar que el parámetro no esté vacío
    If Len(Trim(strTipoOrigen)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 8101, strFuncion, _
            "Parámetro strTipoOrigen está vacío"
    End If
    
    ' Verificar longitud razonable del parámetro
    If Len(Trim(strTipoOrigen)) > 50 Then
        Err.Raise ERROR_BASE_IMPORT + 8102, strFuncion, _
            "Parámetro strTipoOrigen demasiado largo: " & Len(Trim(strTipoOrigen)) & " caracteres"
    End If
    
    ' Normalizar a mayúsculas y eliminar espacios
    strTipoNormalizado = UCase(Trim(strTipoOrigen))
    
    '--------------------------------------------------------------------------
    ' 2. Identificación del método de obtención según el tipo especificado
    '--------------------------------------------------------------------------
    lngLineaError = 40
    
    ' Validar tipos de origen soportados exactamente como especificado
    Select Case strTipoNormalizado
        Case "TEMP"
            strMetodoUsado = "Variable de entorno TEMP del sistema"
            
        Case "TMP"
            strMetodoUsado = "Variable de entorno TMP del sistema"
            
        Case "USERPROFILE"
            strMetodoUsado = "Variable de entorno USERPROFILE del sistema"
            
        Case "CURRENT_DIR"
            strMetodoUsado = "Directorio de trabajo actual"
            
        Case "EXCEL_PATH_CURRENT_BOOK"
            strMetodoUsado = "Carpeta del libro Excel actual"
            
        Case "EXCEL_PATH_TEMP"
            strMetodoUsado = "Carpeta temporal de Excel"
            
        Case Else
            Err.Raise ERROR_BASE_IMPORT + 8103, strFuncion, _
                "Tipo de origen no soportado: " & Chr(34) & strTipoNormalizado & Chr(34) & vbCrLf & _
                "Tipos válidos: TEMP, TMP, USERPROFILE, CURRENT_DIR, EXCEL_PATH_CURRENT_BOOK, EXCEL_PATH_TEMP"
    End Select
    
    '--------------------------------------------------------------------------
    ' 3. Ejecución del método específico para cada tipo de origen
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    Select Case strTipoNormalizado
        Case "TEMP"
            '----------------------------------------------------------------------
            ' 3A. Variable de entorno %TEMP% del sistema
            '----------------------------------------------------------------------
            lngLineaError = 51
            
            ' Obtener valor usando función Environ() compatible Excel 97-365
            strRutaObtenida = Environ("TEMP")
            
            If Len(strRutaObtenida) = 0 Then
                Call fun801_LogMessage("WARNING - Variable de entorno TEMP no encontrada o vacía", _
                    False, "", strFuncion)
                Exit Function
            End If
            
        Case "TMP"
            '----------------------------------------------------------------------
            ' 3B. Variable de entorno %TMP% del sistema
            '----------------------------------------------------------------------
            lngLineaError = 52
            
            ' Obtener valor usando función Environ() compatible Excel 97-365
            strRutaObtenida = Environ("TMP")
            
            If Len(strRutaObtenida) = 0 Then
                Call fun801_LogMessage("WARNING - Variable de entorno TMP no encontrada o vacía", _
                    False, "", strFuncion)
                Exit Function
            End If
            
        Case "USERPROFILE"
            '----------------------------------------------------------------------
            ' 3C. Variable de entorno %USERPROFILE% del sistema
            '----------------------------------------------------------------------
            lngLineaError = 53
            
            ' Obtener valor usando función Environ() compatible Excel 97-365
            strRutaObtenida = Environ("USERPROFILE")
            
            If Len(strRutaObtenida) = 0 Then
                Call fun801_LogMessage("WARNING - Variable de entorno USERPROFILE no encontrada o vacía", _
                    False, "", strFuncion)
                Exit Function
            End If
            
        Case "CURRENT_DIR"
            '----------------------------------------------------------------------
            ' 3D. Directorio de trabajo actual
            '----------------------------------------------------------------------
            lngLineaError = 54
            
            ' Usar CurDir() que es compatible con Excel 97-365
            strRutaObtenida = CurDir()
            
            If Len(strRutaObtenida) = 0 Then
                Call fun801_LogMessage("WARNING - No se pudo obtener directorio de trabajo actual", _
                    False, "", strFuncion)
                Exit Function
            End If
            
        Case "EXCEL_PATH_CURRENT_BOOK"
            '----------------------------------------------------------------------
            ' 3E. Carpeta del libro Excel actual
            '----------------------------------------------------------------------
            lngLineaError = 55
            
            ' Usar ThisWorkbook.Path que es compatible Excel 97-365
            strRutaObtenida = ThisWorkbook.Path
            
            ' Manejar casos especiales para OneDrive/SharePoint/Teams
            If Len(strRutaObtenida) = 0 Then
                ' El libro no está guardado o está en ubicación en línea
                Call fun801_LogMessage("WARNING - Libro no guardado o en ubicación en línea (OneDrive/SharePoint/Teams)", _
                    False, "", strFuncion)
                
                ' Intentar alternativa extrayendo ruta desde FullName
                If Len(ThisWorkbook.FullName) > 0 Then
                    intUltimaBarraDiagonal = InStrRev(ThisWorkbook.FullName, "\")
                    If intUltimaBarraDiagonal > 0 Then
                        strRutaObtenida = Left(ThisWorkbook.FullName, intUltimaBarraDiagonal - 1)
                        Call fun801_LogMessage("INFO - Ruta extraída desde FullName: " & strRutaObtenida, _
                            False, "", strFuncion)
                    End If
                End If
                
                ' Si aún no hay ruta válida, usar directorio actual como fallback
                If Len(strRutaObtenida) = 0 Then
                    strRutaObtenida = CurDir()
                    Call fun801_LogMessage("INFO - Usando directorio actual como fallback para Excel: " & _
                        strRutaObtenida, False, "", strFuncion)
                End If
            End If
            
        Case "EXCEL_PATH_TEMP"
            '----------------------------------------------------------------------
            ' 3F. Carpeta temporal de Excel
            '----------------------------------------------------------------------
            lngLineaError = 56
            
            ' Excel no tiene carpeta temporal específica, usar jerarquía de fallback:
            ' 1. Intentar TEMP del sistema
            ' 2. Si falla, intentar TMP del sistema
            ' 3. Si falla, usar directorio actual
            
            strRutaObtenida = Environ("TEMP")
            If Len(strRutaObtenida) = 0 Then
                strRutaObtenida = Environ("TMP")
                If Len(strRutaObtenida) > 0 Then
                    Call fun801_LogMessage("INFO - Usando TMP como fallback para Excel temp", _
                        False, "", strFuncion)
                End If
            End If
            
            If Len(strRutaObtenida) = 0 Then
                strRutaObtenida = CurDir()
                Call fun801_LogMessage("INFO - Usando directorio actual como fallback para Excel temp: " & _
                    strRutaObtenida, False, "", strFuncion)
            End If
            
            If Len(strRutaObtenida) = 0 Then
                Call fun801_LogMessage("ERROR - No se pudo obtener ninguna carpeta temporal válida", _
                    True, "", strFuncion)
                Exit Function
            End If
    End Select
    
    '--------------------------------------------------------------------------
    ' 4. Manejo de casos especiales y fallbacks automáticos
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Verificar que se obtuvo algún resultado válido
    If Len(strRutaObtenida) = 0 Then
        Call fun801_LogMessage("ERROR - No se obtuvo resultado para tipo: " & strTipoNormalizado, _
            True, "", strFuncion)
        Exit Function
    End If
    
    ' Verificar caracteres peligrosos o inválidos usando CHR() como solicitado
    If InStr(strRutaObtenida, Chr(34)) > 0 Or _   ' Comillas dobles "
       InStr(strRutaObtenida, Chr(60)) > 0 Or _   ' Menor que <
       InStr(strRutaObtenida, Chr(62)) > 0 Or _   ' Mayor que >
       InStr(strRutaObtenida, Chr(124)) > 0 Or _  ' Pipe |
       InStr(strRutaObtenida, Chr(42)) > 0 Or _   ' Asterisco *
       InStr(strRutaObtenida, Chr(63)) > 0 Then   ' Interrogación ?
        
        Err.Raise ERROR_BASE_IMPORT + 8104, strFuncion, _
            "Ruta contiene caracteres no válidos o peligrosos: " & strRutaObtenida
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Normalización y limpieza de la ruta obtenida
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Eliminar espacios al inicio y final
    strRutaLimpia = Trim(strRutaObtenida)
    
    ' Eliminar barra diagonal final si existe (normalización, excepto para raíces como C:\)
    If Right(strRutaLimpia, 1) = "\" And Len(strRutaLimpia) > 3 Then
        strRutaLimpia = Left(strRutaLimpia, Len(strRutaLimpia) - 1)
    End If
    
    ' Verificar que la ruta tiene una longitud mínima razonable
    If Len(strRutaLimpia) < 3 Then
        Err.Raise ERROR_BASE_IMPORT + 8105, strFuncion, _
            "Ruta demasiado corta para ser válida: " & Chr(34) & strRutaLimpia & Chr(34) & _
            " (Longitud: " & Len(strRutaLimpia) & " caracteres)"
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Validación opcional de existencia de la carpeta
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    If blnValidarCarpeta Then
        ' Crear objeto FileSystemObject compatible con Excel 97-365
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        ' Verificar si la carpeta existe realmente
        If Not objFSO.FolderExists(strRutaLimpia) Then
            Call fun801_LogMessage("WARNING - Carpeta obtenida no existe físicamente: " & strRutaLimpia & _
                " (Método: " & strMetodoUsado & ", Tipo: " & strTipoNormalizado & ")", _
                False, "", strFuncion)
            
            ' Para rutas de Excel en OneDrive/SharePoint, esto puede ser normal
            ' No fallar completamente, solo registrar advertencia y continuar
        Else
            Call fun801_LogMessage("VALIDACION - Carpeta existe y es accesible: " & strRutaLimpia, _
                False, "", strFuncion)
        End If
        
        ' Limpiar objeto
        Set objFSO = Nothing
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Logging detallado de la operación realizada
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    ' Registrar éxito completo con detalles para debugging y auditoría
    Call fun801_LogMessage("EXITO COMPLETO - " & strMetodoUsado & " obtenida exitosamente. " & _
        "Tipo solicitado: " & Chr(34) & strTipoOrigen & Chr(34) & ", " & _
        "Tipo normalizado: " & Chr(34) & strTipoNormalizado & Chr(34) & ", " & _
        "Ruta final: " & Chr(34) & strRutaLimpia & Chr(34), _
        False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 8. Retorno del resultado final normalizado
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun808_ObtenerCarpetaSistema = strRutaLimpia
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Limpiar objetos en caso de error
    On Error Resume Next
    Set objFSO = Nothing
    On Error GoTo 0
    
    ' Construir mensaje de error detallado y completo
    strMensajeError = "ERROR CRITICO en " & strFuncion & vbCrLf & _
                      "Fecha y hora: 2025-06-12 15:19:14 UTC" & vbCrLf & _
                      "Usuario: david-joaquin-corredera-de-colsa" & vbCrLf & _
                      "Línea aproximada: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción del Error: " & Err.Description & vbCrLf & _
                      "Parámetro de entrada original: " & Chr(34) & strTipoOrigen & Chr(34) & vbCrLf & _
                      "Tipo normalizado: " & Chr(34) & strTipoNormalizado & Chr(34) & vbCrLf & _
                      "Método utilizado: " & Chr(34) & strMetodoUsado & Chr(34) & vbCrLf & _
                      "Ruta obtenida (cruda): " & Chr(34) & strRutaObtenida & Chr(34) & vbCrLf & _
                      "Ruta limpia (procesada): " & Chr(34) & strRutaLimpia & Chr(34) & vbCrLf & _
                      "Validar carpeta habilitado: " & blnValidarCarpeta & vbCrLf & _
                      "Compatibilidad: Excel 97/2003/2007/365, OneDrive/SharePoint/Teams"
    
    ' Registrar error completo en el log del sistema
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Para debugging en desarrollo (solo visible en VBA Editor)
    Debug.Print strMensajeError
    
    ' Retornar cadena vacía para indicar fallo
    fun808_ObtenerCarpetaSistema = ""
End Function
