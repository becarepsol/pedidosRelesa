Attribute VB_Name = "moduleMain"
Option Explicit

Public Sub main()

    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source="main"

    Dim intStock As Integer
    Dim codigo As String
    Dim xCount As Variant
    Dim hojPed As Object, hojProno As Object, hojStock As Object
    Dim stockTrans as Integer, stockGeneral As Integer, stockProvisional as Integer
    Dim promVentMes As Integer
    Dim xOffset as Variant, yOffset as Variant

    set hojPed = Sheets("Seleccionados").Range("A3")
    set hojStock = Sheets("Stock").Range("A2")
    set hojProno = Sheets("Pronostico").Range("A3")

    xCount = 1
    Do While hojProno.offset(xCount,0) <> ""
        xCount = xCount + 1
        stockGeneral = hojStock.offset(xOffset,4)
        stockTrans = hojStock.offset(xOffset,5)
        promVentMes = PromVentasMes(codigo,3)
        stockProvisional = StockProvisional(stockGeneral, stockTrans, pomVentMes)
        alcance = Alcance(stockProvisional, promVentMes)
    Loop

End Sub