Attribute VB_Name = "routines"

Public Function FinalAlcance(ByVal codigo as String)

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
            promVentMes = finalAlcance
        End If

        stockProvisional = ProvisionalStock(stockGeneral,stockTrans, promVentMes)
        finalAlcance = Alcance(stockProvisional, promVentMes)

    Next periodo

End Sub

public Sub PrintValues(  _
    ByVal pedido as Long _
    ByVal codigo as String)



End Sub
