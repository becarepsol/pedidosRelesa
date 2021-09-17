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
    Dim promVentMes as Long

    Dim stckProv As Long
    Dim stockAlcance As Long
    Dim pronos as Long
    Dim pronosAjustado As Long

    Dim xOffset As Long

    Set hojPed = sheets("Pedidos").range("A3")
    Set hojStock = sheets("Stock").range("A2")
    Set hojProno = sheets("Pronostico").range("A3")

    xOffset = 0
    codigo = hojPed.offset(xOffset,0)
    Do While codigo <> vbNullString

        xOffset = xOffset + 1
        codigo = hojPed.offset(xOffset,0)

        stockAlcance = AlcanceFinal(codigo)
        stckProv = ProvisionFinal(codigo)

        promVentMes = PromVentasMes(codigo, 1)

        If (promVentMes <= 0) Then

            pedido = 0
            stckProv = 0
            stockAlcance = 0
            Call PrintValues(pedido, codigo, stckProv, stockAlcance)

        Else

            If Suficiente(stockAlcance) Then

                pedido = 0
                Call PrintValues(pedido, codigo, stckProv, stockAlcance)

                Else

                pronos = Pronostico(codigo, hojProno)
                pronosAjustado = AjustePronos(codigo, pronos, stckProv)
                pedido = pronosAjustado
                stckProv = stckProv + pedido
                stockAlcance = stckProv / promVentMes
                Call PrintValues(pedido, codigo, stckProv, stockAlcance)

            End If

        End If

    Loop

End Sub
