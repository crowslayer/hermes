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
Public Sub GetStatusColumns(ByVal ws As Worksheet, ByRef colStatus As Long, ByRef colDate As Long)
    
    Dim i As Long
    Dim lastCol As Long
    Dim header As String
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
      
    For i = 1 To lastCol
        header = UCase$(Trim$(CStr(ws.Cells(1, i).value)))
        
        Select Case header

            Case "STATUS"
                colStatus = i

            Case "STATUS_DATE"
                colDate = i

        End Select
    Next i
    
    If colStatus = 0 Then

        colStatus = lastCol + 1
        ws.Cells(1, colStatus).value = "STATUS"

    End If

    If colDate = 0 Then

        colDate = colStatus + 1
        ws.Cells(1, colDate).value = "STATUS_DATE"

    End If
    
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
    Dim header As String
    Dim colStatus As Long
    Dim colDate As Long
    
    ' Lista de columnas mínimas requeridas
    reqHeaders = Array("ADJUNTOS", "TO", "CC", "BCC", "ASUNTO", "PLANTILLA")
    
    ' Verificar que existe la hoja BASE
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(DEFAULT_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "ERROR: No existe una hoja llamada " & DEFAULT_SHEET_NAME & " en este archivo.", vbCritical
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
        header = Trim(UCase(ws.Cells(1, i).value))
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
    
    ' agregando campos de validacion
    
    GetStatusColumns ws, colStatus, colDate
    
    ' Si llegó aquí, todo está bien
    ValidateSheet = True
End Function


Public Function BuildFileIndex(folderObj As Object) As Object

    Dim dict As Object
    Dim fileObj As Object
    Dim fso As Object
    Dim fileKey As String
    Dim filesCol As Collection

    Set dict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each fileObj In folderObj.Files
    
        fileKey = UCase$(fso.GetBaseName(fileObj.Path))
        
        If Not dict.Exists(fileKey) Then

            Set filesCol = New Collection
            dict.Add fileKey, filesCol

        End If
        dict(fileKey).Add fileObj.Path
        'dict(UCase$(fso.GetBaseName(fileObj.Path))) = fileObj.Path

    Next fileObj

    Set BuildFileIndex = dict
    Set fso = Nothing

End Function

