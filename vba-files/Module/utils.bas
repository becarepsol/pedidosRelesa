Attribute VB_Name = "utils"
Option Explicit

Public Function Xposition( _
    ByVal target As String, _
    ByRef pivot As Range) As Variant

    Dim xCount As Variant
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

Public Function Suficiente(ByVal stockAlcance as Long) As Boolean

    If stockAlcance > 3 Then
        Suficiente = true
        Else
        Suficiente = false
    End If

End Function
