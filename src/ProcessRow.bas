Option Explicit

' =====================================================================
' Procesamiento de una fila individual de la hoja BASE.
' =====================================================================

Public Sub RowProcessing( _
    ByVal outlookApp As Object, _
    ByVal fso As Object, _
    ByVal fileIndex As Object, _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal modo As String, _
    ByRef processedCount As Long _
)

    On Error GoTo ERR_FILA

    Dim templatePath As String, fullTemplate As String
    Dim mail As Object
    Dim expectedFiles As Collection
    Dim fileObj As Object
    Dim fileBase As Variant
    Dim dictVars As Object
    Dim colProc As Long, colDate As Long
    Dim result As String
        
    ' Leer plantilla
    templatePath = Trim$(CStr(ws.Cells(rowIndex, 6).value))
    If templatePath = "" Then
        MsgBox "Fila " & rowIndex & ": No se especificó plantilla.", vbExclamation
        Exit Sub
    End If

    ' Extensión .oft
    If LCase$(Right$(templatePath, 4)) <> ".oft" Then
        templatePath = templatePath & ".oft"
    End If

    ' Resolver ruta
    If fso.FileExists(templatePath) Then
        fullTemplate = templatePath
    Else
        fullTemplate = DEFAULT_TEMPLATE_DIR & templatePath
    End If

    If Not fso.FileExists(fullTemplate) Then
        MsgBox "Fila " & rowIndex & ": Plantilla no encontrada: " & vbCrLf & fullTemplate, vbExclamation
        Exit Sub
    End If

    ' Crear correo
    Set mail = outlookApp.CreateItemFromTemplate(fullTemplate)

    ' Archivos esperados
    Set expectedFiles = StringToCollection(CStr(ws.Cells(rowIndex, 1).value))

    ' Adjuntar archivos
    Dim attachmentPath As Variant
    Dim attachmentsList As Collection
    
    For Each fileBase In expectedFiles
        fileBase = UCase$(Trim$(CStr(fileBase)))
        
        If fileIndex.Exists(fileBase) Then
            Set attachmentsList = fileIndex(fileBase)
            For Each attachmentPath In attachmentsList
                mail.Attachments.Add attachmentPath
            Next attachmentPath
        End If
    Next fileBase
    
    ' If mail.Attachments.Count = 0 Then
    '     ws.Cells(rowIndex, colProc).value = "Error"
    '     ws.Cells(rowIndex, colDate).value = Now
    '    MsgBox "Fila " & rowIndex & ": No se encontraron adjuntos.", vbInformation
    '    Exit Sub
    '    GoTo CLEANUP
    ' End If

    ' Validación de destinatario
    If Trim$(ws.Cells(rowIndex, 2).value) = "" Then
        MsgBox "Fila " & rowIndex & ": Falta destinatario.", vbExclamation
        Exit Sub
    End If

    ' Completar campos básicos
    mail.To = ws.Cells(rowIndex, 2).value
    mail.CC = ws.Cells(rowIndex, 3).value
    mail.BCC = ws.Cells(rowIndex, 4).value
    mail.Subject = ws.Cells(rowIndex, 5).value

    ' Variables dinámicas
    Set dictVars = DictVariablesFromRow(ws, rowIndex)
    mail.HTMLBody = ApplyVariablesToHTMLProtected(mail.HTMLBody, dictVars)

    ' Enviar o mostrar
    If UCase$(modo) = MODE_SEND Then
        mail.Send
        result = "Sent"
    Else
        mail.Display
        result = "Display"
    End If
    
    
    ' Columnas dinámicas para marcas
    GetStatusColumns ws, colProc, colDate
    ws.Cells(rowIndex, colProc).value = result
    ws.Cells(rowIndex, colDate).value = Now

    processedCount = processedCount + 1

CLEANUP:

Set dictVars = Nothing
Set expectedFiles = Nothing
Set mail = Nothing

    Exit Sub

ERR_FILA:
    ws.Cells(rowIndex, colProc).value = "Error"
    ws.Cells(rowIndex, colDate).value = Now
    MsgBox "Error en fila " & rowIndex & ": " & Err.Description, vbExclamation
    Err.Clear
    
    Resume CLEANUP
End Sub
