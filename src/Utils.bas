Option Explicit
' =====================================================================
' Author:  Domingo Herrera
' Email:   crowslayer@gmail.com
' License: GNU AGPLv3 <https://spdx.org/licenses/AGPL-3.0-or-later.html>
'
' Funciones de apoyo, reutilizables en todo el proyecto.
' =====================================================================

' Convierte string "a|b|c" en Collection
Public Function StringToCollection(ByVal s As String) As Collection
    Dim col As New Collection
    Dim part As Variant

    For Each part In Split(s, "|")
        If Trim$(part) <> "" Then col.Add Trim$(part)
    Next part

    Set StringToCollection = col
End Function

' Verifica si un valor está en la Collection (case-insensitive)
Public Function Contains(col As Collection, key As String) As Boolean
    Dim item As Variant
    For Each item In col
        If StrComp(item, key, vbTextCompare) = 0 Then
            Contains = True
            Exit Function
        End If
    Next item
End Function

' Obtiene columnas dinámicas para el marcado de estatus
Public Sub GetStatusColumns(ws As Worksheet, ByRef colProcessed As Long, ByRef colDate As Long)
    Dim lastHeaderCol As Long
    lastHeaderCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    colProcessed = lastHeaderCol + 1
    colDate = lastHeaderCol + 2
End Sub

'---------------------------------------------------------------------
' Valida existencia de hoja BASE y columnas mínimas requeridas
'---------------------------------------------------------------------
Public Function ValidateSheet() As Boolean
    Dim ws As Worksheet
    Dim reqHeaders As Variant
    Dim foundHeaders As Object
    Dim lastCol As Long 
    Dim i As Long
    Dim header as String
    
    ' Lista de columnas mínimas requeridas
    reqHeaders = Array("ADJUNTOS", "TO", "CC", "BCC", "ASUNTO", "PLANTILLA")
    
    ' Verificar que existe la hoja BASE
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DEFAULT_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "ERROR: No existe una hoja llamada" & DEFAULT_SHEET_NAME &"en este archivo.", vbCritical
        Exit Function
    End If
    
    ' Determinar última columna con datos en la FILA 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Si fila 1 está vacía, error inmediato
    If lastCol < 1 Then
        MsgBox "ERROR: La hoja BASE no contiene encabezados.", vbCritical
        Exit Function
    End If

    ' Crear diccionario para marcas
    Set foundHeaders = CreateObject("Scripting.Dictionary")
    
    ' Leer encabezados de la fila 1
    For i = 1 To lastCol
        header = Trim(UCase(ws.Cells(1, i).Value))
        If header <> "" Then foundHeaders(header) = True
    Next i
    
    ' Confirmar que todos los encabezados requeridos existen
    For i = LBound(reqHeaders) To UBound(reqHeaders)
        If Not foundHeaders.Exists(UCase(reqHeaders(i))) Then
            MsgBox "ERROR: Falta la columna obligatoria '" & reqHeaders(i) & _
                   "' en la hoja BASE.", vbCritical
            Exit Function
        End If
    Next i
    
    ' Si llegó aquí, todo está bien
    ValidateSheet = True
End Function
