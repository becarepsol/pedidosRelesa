Attribute VB_Name = "routines"

Public Function AlcanceFinal(ByVal codigo as String) as Long

    Dim finalAlcance as Long
    Dim promVentMes As Long

    Dim stockGeneral As Long
    Dim stockTrans As Long
    Dim stockProvisional As Long

    Dim periodo as Long

    For periodo = 1 To 3

        stockTrans = TransStock(codigo, periodo)

        If (periodo = 1) Then
            stockGeneral = GeneralStock(codigo)
            promVentMes = PromVentasMes(codigo, periodo)
            else
            stockGeneral = stockProvisional
        End If

        stockProvisional = ProvisionalStock(stockGeneral,stockTrans, promVentMes)
        finalAlcance = Alcance(stockProvisional, promVentMes)

    Next periodo

    AlcanceFinal = finalAlcance

End Function

Public Function ProvisionFinal(ByVal codigo as String) as Long

    Dim promVentMes As Long

    Dim stockGeneral As Long
    Dim stockTrans As Long
    Dim stockProvisional As Long

    Dim periodo as Long

    For periodo = 1 To 3

        stockTrans = TransStock(codigo, periodo)

        If (periodo = 1) Then
            stockGeneral = GeneralStock(codigo)
            promVentMes = PromVentasMes(codigo, periodo)
            else
            stockGeneral = stockProvisional
        End If

        stockProvisional = ProvisionalStock(stockGeneral,stockTrans, promVentMes)

    Next periodo

    ProvisionFinal = stockProvisional

End Function

Public Sub PrintValues(  _
    ByVal pedido As Long, _
    ByVal codigo As String, _
    ByVal stckProv As Long, _
    ByVal stockAlcance As Long)

    Dim undxPalet As Long
    Dim numPalets As Long
    Dim pedLitros As Long

    Dim stckPosition As Variant
    Dim stckLitros As Long
    Dim pedPosition As Variant

    pedPosition = Xposition(codigo, hojPed)
    stckPosition = Xposition(codigo, hojStock)
    undxPalet = hojStock.offset(stckPosition, 8)

    If pedido > 0 Then

        stckLitros = hojStock.offset(stckPosition, 3)
        numPalets = Round(pedido/undxPalet)
        pedLitros = Round(pedido * stckLitros)

        Else
        numPalets = 0
        pedLitros = 0

    End If

    ' Print one by one
    hojPed.offset(pedPosition, 4) = undxPalet
    hojPed.offset(pedPosition, 5) = numPalets
    hojPed.offset(pedPosition, 6) = pedido
    hojPed.offset(pedPosition, 7) = pedLitros
    hojPed.offset(pedPosition, 8) = stckProv ' Falta este
    hojPed.offset(pedPosition, 9) = stockAlcance

End Sub
