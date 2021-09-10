Attribute VB_Name = "functions"

Option Explicit

Public Function ProvisionalStock( _
    ByVal stockGeneral As Long, _
    ByVal stockTrans As Long, _
    ByVal promVentMes As Long) As Long

    On Error Resume Next
    'Set Error Source Macro/Function name
    Err.Source = "ProvisionalStock"

    If (stockGeneral + stockTrans - promVentMes) < 0 Then
        ProvisionalStock = 0
        Else
        ProvisionalStock = round((stockGeneral + stockTrans - promVentMes), 1)
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
    ByRef codigo As String, _
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
    If xCount = False Then
        PromVentasMes = 0
        Exit Function
    End If
    PromVentasMes = pivot.offset(xCount,15 + period)
    If (PromVentasMes < 0) Then
        PromVentasMes = 0
        MsgBox "El promedio de ventas del codigo " & codigo & " es negativo"
    End If

End Function

Public Function GeneralStock(ByVal codigo as String) as Long

    Dim pivot as Range
    Dim xCount as Long

    Set pivot = sheets("Stock").range("A2")
    xCount = Xposition(codigo, pivot)
    If xCount = False Then
        GeneralStock = 0
        Exit Function
    End If
    GeneralStock = pivot.offset(xCount, 4)

End Function

Public Function TransStock( _
    ByVal codigo As String, _
    ByVal period As Long) as Long

    Dim pivot As Range
    Dim xCount as Long

    Set pivot = sheets("Stock").range("A2")
    xCount = Xposition(codigo, pivot)
    If xCount = False Then
        TransStock = 0
        Exit Function
    End If
    TransStock = pivot.offset(xCount, 4 + period)

End Function
