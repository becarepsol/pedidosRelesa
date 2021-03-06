Attribute VB_Name = "main"
Option Explicit

'@EntryPoint "Main program structure"
Public Sub main()

    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "main"

    Dim codigo As String
    Dim xCount As Variant
    Dim hojPed As Object, hojProno As Object, hojStock As Object
    Dim stockTrans As Long, stockGeneral As Long, stockProvisional As Long
    Dim promVentMes As Long
    Dim xOffset As Variant, yOffset As Variant

    Set hojPed = sheets("Seleccionados").range("A3")
    Set hojStock = sheets("Stock").range("A2")
    Set hojProno = sheets("Pronostico").range("A3")

    xCount = 1
    Do While hojProno.offset(xCount, 0) <> vbNullString
        xCount = xCount + 1
        stockGeneral = hojStock.offset(xOffset, 4)
        stockTrans = hojStock.offset(xOffset, 5)
        promVentMes = PromVentasMes(codigo, 3)
        stockProvisional = stockProvisional(stockGeneral, stockTrans, pomVentMes)
        Alcance = Alcance(stockProvisional, promVentMes)
    Loop

End Sub
