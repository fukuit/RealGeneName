Function RGN(obj As Object) As String
If IsDate(obj) Then
    If DateValue(obj) = CDate("3/1") Then
        GN = "MARCH1"
    ELSEIf DateValue(obj) = CDate("9/1") Then
        GN = "SEPT1"
    ElseIf DateValue(obj) = CDate("9/2") Then
        GN = "SEPT2"
    ElseIf DateValue(obj) = CDate("9/3") Then
        GN = "SEPT3"
    ElseIf DateValue(obj) = CDate("9/4") Then
        GN = "SEPT4"
    ElseIf DateValue(obj) = CDate("9/5") Then
        GN = "SEPT5"
    ElseIf DateValue(obj) = CDate("9/6") Then
        GN = "SEPT6"
    ElseIf DateValue(obj) = CDate("9/7") Then
        GN = "SEPT7"
    ElseIf DateValue(obj) = CDate("9/8") Then
        GN = "SEPT8"
    ElseIf DateValue(obj) = CDate("9/9") Then
        GN = "SEPT9"
    ElseIf DateValue(obj) = CDate("9/10") Then
        GN = "SEPT10"
    ElseIf DateValue(obj) = CDate("9/11") Then
        GN = "SEPT11"
    ElseIf DateValue(obj) = CDate("9/12") Then
        GN = "SEPT12"
    ElseIf DateValue(obj) = CDate("9/13") Then
        GN = "SEPT13"
    ElseIf DateValue(obj) = CDate("9/14") Then
        GN = "SEPT14"
    ElseIf DateValue(obj) = CDate("9/15") Then
        GN = "SEPT15"
    ElseIf DateValue(obj) = CDate("10/1") Then
        GN = "OCT1"
    ElseIf DateValue(obj) = CDate("10/2") Then
        GN = "OCT2"
    ElseIf DateValue(obj) = CDate("10/3") Then
        GN = "OCT3"
    ElseIf DateValue(obj) = CDate("10/4") Then
        GN = "OCT4"
    ElseIf DateValue(obj) = CDate("12/1") Then
        GN = "DEC1"
    ELSE
        GN = CStr(obj)
    End If
Else
    If CStr(obj) = "41518" Then
        GN = "SEPT1"
    ElseIf CStr(obj) = "41519" Then
        GN = "SEPT2"
    ElseIf CStr(obj) = "41520" Then
        GN = "SEPT3"
    ElseIf CStr(obj) = "41521" Then
        GN = "SEPT4"
    ElseIf CStr(obj) = "41522" Then
        GN = "SEPT5"
    ElseIf CStr(obj) = "41523" Then
        GN = "SEPT6"
    ElseIf CStr(obj) = "41524" Then
        GN = "SEPT7"
    ElseIf CStr(obj) = "41525" Then
        GN = "SEPT8"
    ElseIf CStr(obj) = "41526" Then
        GN = "SEPT9"
    ElseIf CStr(obj) = "41527" Then
        GN = "SEPT10"
    ElseIf CStr(obj) = "41528" Then
        GN = "SEPT11"
    ElseIf CStr(obj) = "41529" Then
        GN = "SEPT12"
    ElseIf CStr(obj) = "41530" Then
        GN = "SEPT13"
    ElseIf CStr(obj) = "41531" Then
        GN = "SEPT14"
    ElseIf CStr(obj) = "41532" Then
        GN = "SEPT15"
    Else
        GN = CStr(obj)
    End If
End If
End Function
