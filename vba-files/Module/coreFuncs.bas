Attribute VB_Name = "moduleFunctions"

Option Explicit

Public Function StockProvisional( _
    stockGeneral As Integer, _
    stockTrans As Integer, _
    promVentMes As Integer) As Integer

    On Error Resume Next
    'Set Error Source Macro/Function name
    Err.Source="StockProvisional"

    if (stockGeneral + stockTrans - promVentMes) < 0 then
        StockProvisional = 0
        else
        StockProvisional = Round((stockGeneral + stockTrans - promVentMes), 1)
    end if
End Function

Public Function Alcance( _
    stckProv As Integer, _
    promVentaMes As Integer) As Integer

    On Error Resume Next
    'Set Error Source Macro/Function name
    Err.Source="Alcance"

    If (stckProv <> 0 and promVentMes <> 0) Then
        Alcance = Round(stckProv/promVentMes,1)
        Else
        Alcance = 0
    End If
End Function

Public Function Pronostico( _
    codigo As Variant, _
    pivot as Object)  As Integer

    Dim xCount as Variant
    Dim yOffset As Integer

    xCount = Xposition(codigo, pivot)
    if xCount = False then
        Pronostico = 0
        Exit Function
    end if

    Pronostico = 0
    For yOffset = 0 To 2
        Pronostico = pivot.offset(xCount,4 + yOffset) + Pronostico
    Next yOffset

    Pronostico = round(Pronostico, 1)

End Function

Public Function PromVentasMes( _
    codigo as Variant, _
    period as Integer) as Integer
End Function