Attribute VB_Name = "main"
Option Explicit

'@EntryPoint "Main program structure"
Public Sub main()

    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "main"

    Dim codigo As String

    Dim hojPed As Range
    Dim hojProno As Range
    Dim hojStock As Range

    Dim xOffset As Long
    Dim yOffset As Long

    Set hojPed = sheets("Seleccionados").range("A3")
    Set hojStock = sheets("Stock").range("A2")
    Set hojProno = sheets("Pronostico").range("A3")

    xOffset = 0
    codigo = hojPed.offset(xOffset,0)
    Do While codigo <> vbNullString
        xOffset = xOffset + 1
        codigo = hojPed.offset(xOffset,0)
        stockAlcacne = FinalAlcance(codigo)
        If Suficiente(stockAlcacne) Then
            pedido = 0
            GoTo NextIteration
        End If
        NextIteration:
    Loop

End Sub
