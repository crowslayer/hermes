Option Explicit

' =====================================================================
' Author:  Domingo Herrera
' Email:   crowslayer@gmail.com
' License: GNU AGPLv3 <https://spdx.org/licenses/AGPL-3.0-or-later.html>
'
' Módulo Principal:
'   Punto de entrada de la automatización.
'   Control de flujo, selección de carpeta, control de retraso.
' =====================================================================

Public Const MODE_SEND As String = "ENVIAR"
Public Const MODE_DISPLAY As String = "MOSTRAR"
Public Const DEFAULT_DELAY As Double = 5
Public Const DEFAULT_TEMPLATE_DIR As String = "C:\Templates\"
Public Const DEFAULT_SHEET_NAME As String = "BASE"

' Entrada: Vista previa
Public Sub PROCESAR_PREVISUALIZAR()
    ProcessEmail MODE_DISPLAY
End Sub

' Entrada: Envío automático
Public Sub PROCESAR_ENVIAR()
    ProcessEmail MODE_SEND
End Sub

' ---------------------------------------------------------------------
' Control general del proceso
' ---------------------------------------------------------------------
Public Sub ProcessEmail(ByVal modo As String)

    Dim fso As Object
    Dim dlgFolder As FileDialog
    Dim folderPath As String
    Dim outlookApp As Object
    Dim pathFiles As Object

    Dim delaySeconds As Double
    Dim i As Long, lastRow As Long
    Dim processed As Long
    Dim startTimer As Double
    Dim ws As Worksheet
    Dim fileIndex As Object
   
   ' fix
    Dim colStatus As Long
    Dim colDate As Long
    Dim startRow As Long
    Dim statusValue As String
   
    On Error GoTo ERR_HANDLER
    
    'Validando hoja base
    If Not ValidateSheet() Then Exit Sub
    Set ws = ThisWorkbook.Sheets(DEFAULT_SHEET_NAME)
    
    'Determindo longitud de columnas
    GetStatusColumns ws, colStatus, colDate
    
    ' Selección de carpeta
    Set dlgFolder = Application.FileDialog(msoFileDialogFolderPicker)
    If dlgFolder.Show <> -1 Then Exit Sub
    folderPath = dlgFolder.SelectedItems(1)
    
    ' Crear objetos principales
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set outlookApp = CreateObject("Outlook.Application")
    Set pathFiles = fso.GetFolder(folderPath)
    
    'Listado de archivos
    Set fileIndex = BuildFileIndex(pathFiles)
    
    ' Retraso si es ENVIAR
    If UCase$(modo) = MODE_SEND Then
        delaySeconds = GetDelay()
    Else
        delaySeconds = 0
    End If

    ' Procesamiento por fila
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    processed = 0
    startRow = 0
    
    For i = 2 To lastRow
        
        statusValue = Trim$(UCase$(CStr(ws.Cells(i, colStatus).value)))
    
        If statusValue = "" _
           Or statusValue = "PENDING" _
           Or statusValue = "ERROR" _
           Or statusValue = "DISPLAY" Then
    
            startRow = i
            Exit For
    
        End If
    
    Next i
    
    If startRow = 0 Then
    
        MsgBox "No existen registros pendientes de procesar.", vbInformation
    
        GoTo CLEANUP
    
    End If
        
    Dim pageSize As Long
    Dim currentPageLimit As Long
    
    If modo = MODE_DISPLAY Then
        pageSize = InputBox("¿Cuántos correos deseas mostrar por página?", "Vista previa", 10)
        If pageSize < 1 Then pageSize = 10
        currentPageLimit = pageSize
    End If

    For i = startRow To lastRow
        'Saltando registros procesados
        statusValue = Trim$(UCase$(CStr(ws.Cells(i, colStatus).value)))
        If statusValue = "SENT" Then
            GoTo NEXT_ROW
        End If
        
        RowProcessing _
            outlookApp, _
            fso, _
            fileIndex, _
            ws, _
            i, _
            modo, _
            processed

        ' Control de paginación
        If modo = MODE_DISPLAY Then
            If processed >= currentPageLimit Then
                MsgBox "Mostrados " & processed & " correos. Pulsa Aceptar para continuar.", vbInformation
                currentPageLimit = currentPageLimit + pageSize
            End If
        End If
        
        ' Retardo entre envíos
        If modo = MODE_SEND And delaySeconds > 0 Then
            Application.Wait Now + TimeSerial(0, 0, delaySeconds)
        End If
NEXT_ROW:

    Next i
   
' -------------------------
' Limpieza
' -------------------------
CLEANUP:

Set fileIndex = Nothing
Set pathFiles = Nothing
Set outlookApp = Nothing
Set fso = Nothing
Set dlgFolder = Nothing

    MsgBox "Proceso completado: " & processed & " correos procesados.", vbInformation
    Exit Sub

' -------------------------
' Manejo global de errores
' -------------------------
ERR_HANDLER:
    MsgBox "Error inesperado: " & Err.Description, vbCritical
    Err.Clear

End Sub

' ---------------------------------------------------------------------
' Pregunta el retraso entre envíos
' ---------------------------------------------------------------------
Public Function GetDelay() As Double
    Dim inputValue As Variant
    Dim delaySeconds As Double
    
    If MsgBox("¿Deseas enviar los correos automáticamente?", vbYesNo + vbQuestion) = vbNo Then
        GetDelay = 0
        Exit Function
    End If

    inputValue = InputBox("¿Cuántos segundos deseas esperar entre cada envío?", "Retardo", DEFAULT_DELAY)

    If IsNumeric(inputValue) Then
        delaySeconds = CDbl(inputValue)
    Else
        delaySeconds = DEFAULT_DELAY
    End If

    MsgBox "Se aplicará un retraso de " & delaySeconds & " segundos.", vbInformation
    
    GetDelay = delaySeconds
End Function

