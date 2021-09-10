Attribute VB_Name = "functions"

Option Explicit

Public Function stockProvisional( _
    ByVal stockGeneral As Long, _
    ByVal stockTrans As Long, _
    ByVal promVentMes As Long) As Long

    On Error Resume Next
    'Set Error Source Macro/Function name
    Err.Source = "StockProvisional"

    If (stockGeneral + stockTrans - promVentMes) < 0 Then
        stockProvisional = 0
        Else
        stockProvisional = round((stockGeneral + stockTrans - promVentMes), 1)
    End If
End Function

Public Function Alcance( _
    ByVal stckProv As Long, _
    ByVal promVentMes As Long) As Long

    On Error Resume Next
    'Set Error Source Macro/Function name
    Err.Source = "Alcance"

    If (stckProv <> 0 And promVentMes <> 0) Then
        Alcance = round(stckProv / promVentMes, 1)
        Else
        Alcance = 0
    End If
End Function

Public Function Pronostico( _
    ByRef codigo As Variant, _
    ByRef pivot As Object) As Long

    Dim xCount As Variant
    Dim yOffset As Long

    xCount = Xposition(codigo, pivot)
    If xCount = False Then
        Pronostico = 0
        Exit Function
    End If

    Pronostico = 0
    For yOffset = 0 To 2
        Pronostico = pivot.offset(xCount, 4 + yOffset) + Pronostico
    Next yOffset

    Pronostico = round(Pronostico, 1)

End Function

Public Function PromVentasMes( _
    ByVal codigo As String, _
    ByVal period As Long) As Long

    Dim pivot As Range
    Dim xCount As Long

    set pivot = sheets("VentasxMes2021").range("A2")
    xCount = Xposition(codigo, pivot)
    PromVentasMes = pivot.offset(xCount,15 + period)
    If (PromVentasMes < 0) Then
       PromVentasMes = 0 
    End If

End Function
