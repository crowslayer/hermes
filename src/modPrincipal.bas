Attribute VB_Name = "modPrincipal"
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

Public Const MODE_ENVIAR As String = "ENVIAR"
Public Const MODE_MOSTRAR As String = "MOSTRAR"
Public Const DEFAULT_DELAY As Double = 5
Public Const DEFAULT_TEMPLATE_DIR As String = "C:\Templates\"   ' Ajustable

' Entrada: Vista previa
Public Sub PROCESAR_PREVISUALIZAR()
    ProcesarCorreos MODE_MOSTRAR
End Sub

' Entrada: Envío automático
Public Sub PROCESAR_ENVIAR()
    ProcesarCorreos MODE_ENVIAR
End Sub

' ---------------------------------------------------------------------
' Control general del proceso
' ---------------------------------------------------------------------
Public Sub ProcesarCorreos(ByVal modo As String)

    Dim fso As Object
    Dim dlgFolder As FileDialog
    Dim folderPath As String
    Dim outlookApp As Object
    Dim carpetaArchivos As Object

    Dim delaySeconds As Double
    Dim i As Long, lastRow As Long
    Dim processed As Long
    Dim startTimer As Double
    Dim ws As Worksheet

    On Error GoTo ERR_HANDLER

    ' Selección de carpeta
    Set dlgFolder = Application.FileDialog(msoFileDialogFolderPicker)
    If dlgFolder.Show <> -1 Then Exit Sub
    folderPath = dlgFolder.SelectedItems(1)

    ' Crear objetos principales
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set outlookApp = CreateObject("Outlook.Application")
    Set carpetaArchivos = fso.GetFolder(folderPath)
    Set ws = ThisWorkbook.Sheets("BASE")

    ' Retraso si es ENVIAR
    If UCase$(modo) = MODE_ENVIAR Then
        delaySeconds = ObtenerDelay()
    Else
        delaySeconds = 0
    End If

    ' Procesamiento por fila
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    processed = 0

    For i = 2 To lastRow
        ProcesarFila outlookApp, fso, carpetaArchivos, ws, i, modo, processed

        ' Retardo entre envíos
        If modo = MODE_ENVIAR And delaySeconds > 0 Then
            startTimer = Timer
            Do While Timer < startTimer + delaySeconds
                DoEvents
            Loop
        End If
    Next i

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
Public Function ObtenerDelay() As Double
    Dim inputValue As Variant

    If MsgBox("¿Deseas enviar los correos automáticamente?", vbYesNo + vbQuestion) = vbNo Then
        ObtenerDelay = 0
        Exit Function
    End If

    inputValue = InputBox("¿Cuántos segundos deseas esperar entre cada envío?", "Retardo", DEFAULT_DELAY)

    If IsNumeric(inputValue) Then
        ObtenerDelay = CDbl(inputValue)
    Else
        ObtenerDelay = DEFAULT_DELAY
    End If

    MsgBox "Se aplicará un retraso de " & ObtenerDelay & " segundos.", vbInformation
End Function
