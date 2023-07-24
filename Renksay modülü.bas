Attribute VB_Name = "Module2"
Public Function RENKSAY(pRange1 As Range, pRange2 As Range) As Double

Application.Volatile

Dim rng As Range
For Each rng In pRange1
    If rng.Interior.Color = pRange2.Interior.Color Then
        RENKSAY = RENKSAY + 1
    End If
Next
End Function

