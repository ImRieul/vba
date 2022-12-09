Public Function RangeMin(list as Range) As Long
    RangeMin = Application.Min(list)
End Function
Public Function RangeMax(list as Range) As Long
    RangeMax = Application.Max(list)
End Function

Public Function MaxAddress(ByVal ran as Range) as Range

    dim result as Range

    num = 0
    count = 0

    For Each r In ran

        If (count = 0) Then
            
            count = 1
            num = r.value
            Set Result = r

        End If

        If (r.value > num) Then
            
            num = r.value
            Set Result = r

        End If
    
    Next r 

    Set MaxAddress = result

End Function
    
Public Function MinAddress(ByVal ran as Range) as Range

    dim result as Range

    num = 0
    count = 0

    For Each r In ran

        If (count = 0) Then
            
            count = 1
            num = r.value
            Set Result = r

        End If

        If (r.value < num) Then
            
            num = r.value
            Set Result = r

        End If
    
    Next r 

    Set MinAddress = result

End Function

Public Function ColumnValueMax(ByVal ran, columns as Range) As String

    ColumnValueMax = cells(columns.row, MaxAddress(ran).column).value

End Function

Public Function RowValueMax(Byval ran, rows as Range) As String

    RowValueMax = cells(MaxAddress(ran).row, rows.column).value

End Function

Public Function ColumnValueMin(ByVal ran, columns as Range) As String

    ColumnValueMin = cells(columns.row, MinAddress(ran).column).value

End Function

Public Function RowValueMin(Byval ran, rows as Range) As String

    RowValueMin = cells(MinAddress(ran).row, rows.column).value

End Function

Public  Function Formatter(ByVal str, form As String)

    Formatter = format(str, form)

End Function