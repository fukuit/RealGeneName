
Function RGN(obj As Object) As String
If IsDate(obj) Then
    If DateValue(obj) = CDate("9/1") Then
        GN = "SEPT1"
    End If
Else
    If CStr(obj) = "41518" Then
        GN = "SEPT1"
    Else
    GN = "ERROR"
End If
End If
End Function
