Attribute VB_Name = "modProcesarFila"
Option Explicit

' =====================================================================
' Procesamiento de una fila individual de la hoja BASE.
' =====================================================================

Public Sub ProcesarFila( _
    ByVal outlookApp As Object, _
    ByVal fso As Object, _
    ByVal carpetaArchivos As Object, _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal modo As String, _
    ByRef processedCount As Long _
)

    On Error GoTo ERR_FILA

    Dim templatePath As String, completeTemplate As String
    Dim mail As Object
    Dim expectedFiles As Collection
    Dim fileObj As Object
    Dim fileBase As String
    Dim dictVars As Object
    Dim colProc As Long, colDate As Long

    ' Leer plantilla
    templatePath = Trim$(CStr(ws.Cells(rowIndex, 6).Value))
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
        completeTemplate = templatePath
    Else
        completeTemplate = DEFAULT_TEMPLATE_DIR & templatePath
    End If

    If Not fso.FileExists(completeTemplate) Then
        MsgBox "Fila " & rowIndex & ": Plantilla no encontrada: " & vbCrLf & completeTemplate, vbExclamation
        Exit Sub
    End If

    ' Crear correo
    Set mail = outlookApp.CreateItemFromTemplate(completeTemplate)

    ' Archivos esperados
    Set expectedFiles = StringToCollection(CStr(ws.Cells(rowIndex, 1).Value))

    ' Adjuntar archivos
    For Each fileObj In carpetaArchivos.Files
        fileBase = fso.GetBaseName(fileObj.Path)
        If Contains(expectedFiles, fileBase) Then
            mail.Attachments.Add fileObj.Path
        End If
    Next fileObj

    If mail.Attachments.Count = 0 Then
        MsgBox "Fila " & rowIndex & ": No se encontraron adjuntos.", vbInformation
        Exit Sub
    End If

    ' Completar campos básicos
    mail.To = ws.Cells(rowIndex, 2).Value
    mail.CC = ws.Cells(rowIndex, 3).Value
    mail.BCC = ws.Cells(rowIndex, 4).Value
    mail.Subject = ws.Cells(rowIndex, 5).Value

    ' Variables dinámicas
    Set dictVars = DictVariablesFromRow(ws, rowIndex)
    mail.HTMLBody = ApplyVariablesToHTMLBody(mail.HTMLBody, dictVars)

    ' Enviar o mostrar
    If UCase$(modo) = MODE_ENVIAR Then
        mail.Send
    Else
        mail.Display
    End If

    ' Columnas dinámicas para marcas
    GetStatusColumns ws, colProc, colDate
    ws.Cells(rowIndex, colProc).Value = "Procesado"
    ws.Cells(rowIndex, colDate).Value = Now

    processedCount = processedCount + 1
    Exit Sub

ERR_FILA:
    MsgBox "Error en fila " & rowIndex & ": " & Err.Description, vbExclamation
    Err.Clear
End Sub
