Attribute VB_Name = "test"

Option Explicit

Public Sub test_xPosition()
    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "test_xPosition"

    Dim codigo As String
    Dim pivot As Object

    codigo = "RP165M51"
    Set pivot = sheets("Pronostico").range("A3")

    sheets("Pronostico").range("J11") = Xposition(codigo, pivot)

End Sub

Public Sub test_Pronostico()
    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "test_Pronostico"

    Dim codigo As String
    Dim pivot As Object

    codigo = "RP167N51"
    Set pivot = sheets("Pronostico").range("A3")

    sheets("Pronostico").range("J11") = Pronostico(codigo, pivot)

End Sub

Public Sub test_PromVentasMes()
    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "test_PromVentasMes"

    Dim codigo As String
    Dim period As Long

    codigo = "RP167N51"
    sheets("Pronostico").range("J11") = PromVentasMes(codigo,2)

End Sub

Public Sub test_GeneralStock()

    Dim codigo As String

    codigo = "RP167N51"
    sheets("Pronostico").range("J11") = GeneralStock(codigo)
End Sub

Public Sub test_stockProvisional()
    ' If an error occurs, pass error to VSCode
    On Error Resume Next    ' Defer error handling.
    'Set Error Source
    Err.Source = "test_stockProvisional"

    Dim stockGeneral As Long
    Dim stockTrans As Long
    Dim promVentMes As Long
    Dim codigo As String

    codigo = "RP167N51"
    promVentMes = PromVentasMes(codigo, 1)
    stockTrans = TransStock(codigo, 1)
    stockGeneral = GeneralStock(codigo)

    sheets("Pronostico").range("J11") = stockProvisional(stockGeneral,stockTrans, promVentMes)

End Sub

Public Sub test_TransStock()

    Dim codigo As String

    codigo = "RP167N51"
    sheets("Pronostico").range("J11") = TransStock(codigo, 1)
    sheets("Pronostico").range("J12") = TransStock(codigo, 2)
    sheets("Pronostico").range("J13") = TransStock(codigo, 3)

End Sub

Public Sub test_Alcance()

    Dim stckProv As Long
    Dim promVentMes As Long

    Dim stockGeneral As Long
    Dim stockTrans As Long
    Dim promVentMes As Long

    Dim codigo As String

    codigo = "RP167N51"
    promVentMes = PromVentasMes(codigo, 1)
    stockTrans = TransStock(codigo, 1)
    stockGeneral = GeneralStock(codigo)
    stockProvisional(stockGeneral,stockTrans, promVentMes)

    sheets("Pronostico").range("J11") = Alcance(stockProvisional, promVentMes)

End Sub
