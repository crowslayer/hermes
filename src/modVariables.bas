Attribute VB_Name = "modVariables"
Option Explicit

' =====================================================================
' Manejo de variables dinámicas (diccionario de claves → valores).
' =====================================================================

' Carga todas las variables dinámicas desde la fila.
Public Function DictVariablesFromRow(ws As Worksheet, rowIndex As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long, col As Long
    Dim key As String, value As Variant

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 7 To lastCol
        key = Trim$(CStr(ws.Cells(1, col).Value))
        value = ws.Cells(rowIndex, col).Value

        If key <> "" Then
            dict.Add "[" & UCase$(key) & "]", CStr(value)
        End If
    Next col

    Set DictVariablesFromRow = dict
End Function

' Reemplaza todas las variables del diccionario dentro del HTML.
Public Function ApplyVariablesToHTMLBody(bodyHTML As String, dict As Object) As String
    Dim k As Variant
    For Each k In dict.Keys
        bodyHTML = Replace(bodyHTML, k, dict(k))
    Next k
    ApplyVariablesToHTMLBody = bodyHTML
End Function
