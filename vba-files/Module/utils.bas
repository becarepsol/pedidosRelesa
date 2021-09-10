Attribute VB_Name = "utils"
Option Explicit

Public Function Xposition( _
    ByVal target As String, _
    ByRef pivot As Range) As Variant

    Dim xCount As Integer
    Dim tempValue As String

    xCount = 0
    Do
        xCount = xCount + 1
        tempValue = pivot.offset(xCount, 0)
        If (tempValue = vbNullString) Then
            Xposition = False
            MsgBox target & " no esta en la base de datos"
            Exit Function
        End If
    Loop Until tempValue = target

    Xposition = xCount

End Function
