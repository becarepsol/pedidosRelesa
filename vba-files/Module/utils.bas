Attribute VB_Name = "moduleUtils"
Option Explicit

Public Function Xposition( _
    target As Variant, _
    pivot as Object) As Variant

    Dim xCount as integer
    dim tempValue As Variant

    xCount = 0
    Do
        xCount = xCount + 1
        tempValue = pivot.offset(xCount,0)
        If (tempValue = "") Then
            xPosition = false
            MsgBox target & " no esta en la base de datos"
            Exit Function
        End If
    Loop until tempValue = target

    xPosition = xCount

End Function