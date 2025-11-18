Attribute VB_Name = "Utils"
Option Explicit

' =====================================================================
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
