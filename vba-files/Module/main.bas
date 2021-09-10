Attribute VB_Name = "main"
Option Explicit

Public hojPed As Range
Public hojStock As Range
Public hojProno As Range

'@EntryPoint "Main program structure"
Public Sub main()

    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "main"

    Dim codigo As String
    Dim pedido As Long

    Dim stockAlcacne As Long
    Dim pronosAjustado As Long
    Dim stckProv As Long

    Dim stockGeneral As Long
    Dim stockTrans As Long
    Dim promVentMes As Long

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
        pronosAjustado = Pronostico(codigo, hojProno)
        stockAlcance = FinalAlcance(codigo)

        stockGeneral = GeneralStock(codigo)
        stockTrans = TransStock(codigo, 3)
        promVentMes = PromVentasMes(codigo, 1)

        stckProv = ProvisionalStock(stockGeneral, stockTrans, promVentMes)

        If Suficiente(stockAlcacne) Then

            pedido = 0
            PrintValue(pedido, codigo, stckProv, stockAlcance)

        Else

            pedido = pronosAjustado
            stckProv = stckProv + pedido
            PrintValue(pedido, codigo, stckProv, stockAlcance)

        End If

    Loop

End Sub
