Attribute VB_Name = "formatter"

Public  Function Formatter(ByVal str, form As String)        ' OK
    Formatter = format(str, form)
End Function

Public  Sub AccountingInRange(ByRef rng As Range, form As String)      ' OK
    Dim formatter As String
    Dim zeroFormat As String
    Dim splitForm As Variant

    ' default zero = -
    ' this function zero = 0
    zeroFormat = "0"
    
    If (InStr(form, ".") <> 0) Then
        splitForm = Split(form, ".")

        If (IsNumeric(splitForm(1))) Then
            zeroFormat = "0." & splitForm(1)
        End If
    End If

    ' _ : 빈칸
    ' ; : format text 끝
    formatter = "_-* " _
                & form _
                & "_-;-* " _
                & form _
                & "_-;_-* " _
                & zeroFormat _
                & "_-;_-@_-"   
    
    rng.NumberFormatLocal = formatter

End Sub

Public  Sub PercentInRange(ByRef rng As Range)
    rng.NumberFormatLocal = "0.0%_;"
End Sub