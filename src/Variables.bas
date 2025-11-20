Option Explicit

' =====================================================================
' Author:  Domingo Herrera
' Email:   crowslayer@gmail.com
' License: GNU AGPLv3 <https://spdx.org/licenses/AGPL-3.0-or-later.html>
'
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
        value = ws.Cells(rowIndex, col).Text

        If Not dict.Exists("[" & UCase$(key) & "]") Then
            dict.Add "[" & UCase$(key) & "]", value
        End If
    Next col

    Set DictVariablesFromRow = dict
End Function

' Reemplazo que respetando imágenes incrustadas / CIDs
Public Function ApplyVariablesToHTMLProtected(bodyHTML As String, dict As Object) As String
    Dim key As Variant
    Dim sorted() As Variant
    Dim i As Long, j As Long

    sorted = dict.Keys
    ' Orden descendente por longitud → evita colisiones
    For i = LBound(sorted) To UBound(sorted) - 1
        For j = i + 1 To UBound(sorted)
            If Len(sorted(j)) > Len(sorted(i)) Then
                Dim tmp As Variant
                tmp = sorted(i)
                sorted(i) = sorted(j)
                sorted(j) = tmp
            End If
        Next j
    Next i

    For Each key In sorted
        bodyHTML = Replace(bodyHTML, key, dict(key))
    Next key

    ApplyVariablesToHTMLProtected = bodyHTML
End Function