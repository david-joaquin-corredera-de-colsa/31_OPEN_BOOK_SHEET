Attribute VB_Name = "Modulo_000_Variables_Globales"
Option Explicit

'******************************************************************************
' M�dulo: Global_Variables
' Fecha y Hora de Creaci�n: 2025-05-26 10:04:46 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripci�n:
' Este m�dulo contiene todas las variables globales utilizadas en el sistema
'******************************************************************************

' Constante para version de la Macro
Public Const CONST_MACRO_VERSION As String = "Macro - Version 20250510 - 065125"

' Constante para Scenario Admitido
Public Const CONST_ESCENARIO_ADMITIDO As String = "BUDGET_OS"

' Constante para Ultimo Mes De Carga
Public Const CONST_ULTIMO_MES_DE_CARGA As String = "M12"

' Constante numero de hojas historicas visibles
Public Const CONS_NUM_HOJAS_HCAS_VISIBLES_ENVIO As Integer = 5


' Constante para columna de la Enity
Public Const CONST_COLUMNA_ENTITY As Integer = 4

' Constantes para mostrar o no mensajes durante la ejecuci�n
Public Const CONST_MOSTRAR_MENSAJES_HOJAS_CREADAS As Boolean = False

' Constantes para control de errores
Public Const ERROR_BASE_IMPORT As Long = vbObjectError + 1000

' Constantes para nombres de hojas
Public Const CONST_HOJA_EJECUTAR_PROCESOS As String = "00_Ejecutar_Procesos"
Public Const CONST_HOJA_INVENTARIO As String = "01_Inventario"
Public Const CONST_HOJA_LOG As String = "02_Log"
Public Const CONST_HOJA_USERNAME As String = "05_Username"
Public Const CONST_HOJA_DELIMITADORES_ORIGINALES As String = "06_Delimitadores_Originales"
Public Const CONST_HOJA_REPORT_PL As String = "09_Report_PL"
Public Const CONST_HOJA_REPORT_PL_AH As String = "10_Report_PL_AH"

' Constantes para posicion dimensiones a "modificar" en hoja Informe PL AdHoc
' Constantes para filas:
Public Const CONST_FILA_SCENARIO As Integer = 1
Public Const CONST_FILA_YEAR As Integer = 2
Public Const CONST_FILA_PERIOD As Integer = 3
Public Const CONST_FILA_VIEW As Integer = 4
Public Const CONST_FILA_ENTITY As Integer = 5
Public Const CONST_FILA_VALUE As Integer = 6
Public Const CONST_FILA_ICP As Integer = 7
Public Const CONST_FILA_C4 As Integer = 8
Public Const CONST_FILA_C2_ACTIVITY As Integer = 9
Public Const CONST_FILA_C3_BUSINESS As Integer = 10
' Constantes para COLUMNAS:
Public Const CONST_COLUMNA_INICIAL_HEADERS As Integer = 3
Public Const CONST_COLUMNA_FINAL_HEADERS As Integer = 15
Public Const CONST_COLUMNA_ADICIONAL_HEADERS As Integer = 16

' Constante para hoja de backup
Public Const CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO As String = "BK_"


' Constantes para celdas clave de esas hojas
Public Const CONST_CELDA_USERNAME As String = "C2"
Public Const CONST_CELDA_HEADER_USERNAME As String = "B2" 'Contendra la string "Username:"
Public Const CONST_VALOR_HEADER_USERNAME As String = "Username:" 'Valor que se asigna a la celda anterior


' Constantes para etiquetas de procesamiento de l�neas (NUEVAS)
Public Const CONST_TAG_LINEA_TRATADA As String = "Linea_Tratada"
Public Const CONST_TAG_LINEA_SUMA As String = "Linea_Suma"
Public Const CONST_TAG_LINEA_REPETIDA As String = "Linea_Repetida"

' Constantes para prefijos enlos nombres de las hojas
Public Const CONST_PREFIJO_HOJA_IMPORTACION As String = "Import_"
Public Const CONST_PREFIJO_HOJA_IMPORTACION_WORKING As String = "Import_Working_"
Public Const CONST_PREFIJO_HOJA_IMPORTACION_ENVIO As String = "Import_Envio_"
Public Const CONST_PREFIJO_HOJA_IMPORTACION_COMPROBACION As String = "Import_Comprob_"


' Variable para hoja de env�o anterior
Public gstrPreviaHojaImportacion_Envio As String
' Variable para nombre de copia de hoja de env�o anterior
Public gstrPrevDelHojaImportacion_Envio As String

' Variables para hojas base del sistema
Public gstrHoja_EjecutarProcesos As String
Public gstrHoja_Inventario As String
Public gstrHoja_Log As String
Public gstrHoja_DelimitadoresOriginales As String
Public gstrHoja_UserName As String

' Variables para configuraci�n de importaci�n
Public gstrColumnaInicial_Importacion As String
Public glngFilaInicial_Importacion As Long
Public gstrDelimitador_Importacion As String
Public glngLineaInicial_HojaImportacion As Long
Public glngLineaFinal_HojaImportacion As Long

' Constantes para nombres de hojas


' Variables para nombres de hojas
Public gstrNuevaHojaImportacion As String
Public gstrNuevaHojaImportacion_Working As String
Public gstrNuevaHojaImportacion_Envio As String
Public gstrNuevaHojaImportacion_Comprobacion As String

' Variables para configuraci�n de importaci�n (adicional)
Public vColumnaInicial_Importacion As String
Public vFilaInicial_Importacion As Long
Public vDelimitador_Importacion As String

' Variables para detecci�n de rango
Public vLineaInicial_HojaImportacion As Long
Public vLineaFinal_HojaImportacion As Long

' =============================================================================
' VARIABLES GLOBALES PARA DELIMITADORES DE EXCEL
' =============================================================================
' Fecha y hora de creaci�n: 2025-05-26 17:43:59 UTC
' Autor: david-joaquin-corredera-de-colsa
' Descripci�n: Variables globales para el manejo de delimitadores de Excel
' =============================================================================

Public vHojaDelimitadoresExcelOriginales As String
Public vCelda_Header_Excel_UseSystemSeparators As String
Public vCelda_Header_Excel_DecimalSeparator As String
Public vCelda_Header_Excel_ThousandsSeparator As String
Public vCelda_Valor_Excel_UseSystemSeparators As String
Public vCelda_Valor_Excel_DecimalSeparator As String
Public vCelda_Valor_Excel_ThousandsSeparator As String
Public vExcel_UseSystemSeparators As String
Public vExcel_DecimalSeparator As String
Public vExcel_ThousandsSeparator As String


' =============================================================================
' VARIABLES GLOBALES ADICIONALES PARA RESTAURACI�N DE DELIMITADORES
' =============================================================================
' Fecha y hora de creaci�n: 2025-05-26 18:41:20 UTC
' Usuario: david-joaquin-corredera-de-colsa
' Descripci�n: Variables globales adicionales para restaurar delimitadores originales
' =============================================================================

Public Const CONST_OCULTAR_REPOSITORIO_DELIMITADORES As Boolean = True 'Poner como True si se desea ocultar la hoja
Public Const CONST_OCULTAR_HOJA_USERNAME As Boolean = True 'Poner como True si se desea ocultar la hoja


' Variables para celdas que contienen valores originales
Public vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal As String
Public vCelda_Valor_Excel_DecimalSeparator_ValorOriginal As String
Public vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal As String

' Variables para almacenar valores originales le�dos
Public vExcel_UseSystemSeparators_ValorOriginal As String
Public vExcel_DecimalSeparator_ValorOriginal As String
Public vExcel_ThousandsSeparator_ValorOriginal As String

' AUTOR: Sistema Automatizado
' VERSI�N: 1.0
' COMPATIBILIDAD: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
' =============================================================================

' Variables globales para delimitadores
Public vDelimitadorDecimal_HFM As String
Public vDelimitadorMiles_HFM As String


' =============================================================================
' VARIABLES GLOBALES PARA DETECCI�N DE RANGOS POR PALABRAS CLAVE
' =============================================================================
' Fecha y hora de creaci�n: 2025-06-03 03:19:45 UTC
' Autor: david-joaquin-corredera-de-colsa
' Descripci�n: Variables para detectar rangos basados en palabras clave espec�ficas
' =============================================================================

Public vPalabraClave_PrimeraFila As String
Public vPalabraClave_PrimeraColumna As String
Public vPalabraClave_UltimaFila As String
Public vPalabraClave_UltimaColumna As String


Public Sub fun801_InicializarVariablesGlobales()

    ' =============================================================================
    ' FUNCI�N: fun801_InicializarVariablesGlobales
    ' PROP�SITO: Inicializa las variables globales con valores por defecto
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error GoTo ErrorHandler_fun801
    
    ' Inicializar delimitador decimal si est� vac�o
    If vDelimitadorDecimal_HFM = "" Or vDelimitadorDecimal_HFM = vbNullString Then
        vDelimitadorDecimal_HFM = "."
    End If
    
    ' Inicializar delimitador de miles si est� vac�o
    If vDelimitadorMiles_HFM = "" Or vDelimitadorMiles_HFM = vbNullString Then
        vDelimitadorMiles_HFM = ","
    End If
    
    Exit Sub

ErrorHandler_fun801:
    ' En caso de error, usar valores por defecto
    vDelimitadorDecimal_HFM = "."
    vDelimitadorMiles_HFM = ","
End Sub

'------------------------------------------------------------------------------
' Procedimiento: InitializeGlobalVariables
' Descripci�n: Inicializa todas las variables globales con valores por defecto
'------------------------------------------------------------------------------
Public Sub InitializeGlobalVariables()
    
    'Inicializar hoja de envio anterior
    gstrPreviaHojaImportacion_Envio = ""
    'Inicializar nombre de copia de hoja anterior
    gstrPrevDelHojaImportacion_Envio = ""
    
    'Nombres de hojas base
    gstrHoja_EjecutarProcesos = CONST_HOJA_EJECUTAR_PROCESOS
    gstrHoja_Inventario = CONST_HOJA_INVENTARIO
    gstrHoja_Log = CONST_HOJA_LOG
    gstrHoja_DelimitadoresOriginales = CONST_HOJA_DELIMITADORES_ORIGINALES
    gstrHoja_UserName = CONST_HOJA_USERNAME
        
    'Inicializar variables de l�neas
    glngLineaInicial_HojaImportacion = 0
    glngLineaFinal_HojaImportacion = 0
    
    'Inicializar nombres de hojas
    gstrNuevaHojaImportacion = ""
    gstrNuevaHojaImportacion_Working = ""
    gstrNuevaHojaImportacion_Envio = ""
    gstrNuevaHojaImportacion_Comprobacion = ""
    
    'Adicional
    'Configuraci�n de importaci�n
    vColumnaInicial_Importacion = "B"        ' Columna B (2)
    vFilaInicial_Importacion = 2             ' Fila 2
    vDelimitador_Importacion = ";"           ' Delimitador por defecto
    
    'Inicializar variables de rango
    vLineaInicial_HojaImportacion = 0
    vLineaFinal_HojaImportacion = 0
    
    'Inicializar palabras clave para detecci�n de rangos
    vPalabraClave_PrimeraFila = "BUDGET_OS"      ' Palabra clave para primera fila
    vPalabraClave_PrimeraColumna = "BUDGET_OS"   ' Palabra clave para primera columna
    vPalabraClave_UltimaFila = "BUDGET_OS"       ' Palabra clave para �ltima fila
    vPalabraClave_UltimaColumna = "M12"          ' Palabra clave para �ltima columna
    
    'Configurar nombre de backup autom�tico
    If gstrPreviaHojaImportacion_Envio <> "" Then
        gstrPrevDelHojaImportacion_Envio = CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO & gstrPreviaHojaImportacion_Envio 'Normalmente CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO="BK_"
    End If
    
    
End Sub
