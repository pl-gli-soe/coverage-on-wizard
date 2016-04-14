Attribute VB_Name = "CommonFunctionsModule"
Public Function special_offset(r As Range) As Range
    
    If r.Offset(1, 0).Value <> "" Then
        Set special_offset = r.Offset(1, 0)
    Else
        Set special_offset = r.End(xlDown)
    End If
End Function
