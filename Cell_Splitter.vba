Sub VunjaMifupa()
' Kizito | kizomanizo@gmail.com
Dim LR As Long, i As Long
Dim X As Variant
Application.ScreenUpdating = False
LR = Range("C" & Rows.Count).End(xlUp).Row
Columns("C").Insert
For i = LR To 1 Step -1
    With Range("D" & i)
        If InStr(.Value, "/") = 0 Then
            .Offset(, -1).Value = .Value
        Else
            X = Split(.Value, "/")
            .Offset(1).Resize(UBound(X)).EntireRow.Insert
            .Offset(, -1).Resize(UBound(X) - LBound(X) + 1).Value = Application.Transpose(X)
        End If
    End With
Next i
Columns("D").Delete
LR = Range("C" & Rows.Count).End(xlUp).Row
With Range("A1:E" & LR)
    On Error Resume Next
    .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    On Error GoTo 0
    .Value = .Value
End With
Application.ScreenUpdating = True
End Sub