Attribute VB_Name = "tools"

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

Public  Sub UnMergePull(ByRef rng as Range)     ' OK
    If (rng.MergeCells) Then
        rng.UnMerge
        
        For Each cell In rng
            If cell.value <> rng.item(1, 1).value Then
                cell.value = rng.item(1, 1).value
            End if
        Next cell 
    End If
End Sub


Public  Sub Merge(ByRef rng as Range)       ' OK
    If Not (rng.MergeCells) Then
        Application.DisplayAlerts = False
        rng.Merge False
        rng.HorizontalAlignment = xlCenter
        Application.DisplayAlerts = True
    End If
End Sub


Public  Sub MergeEqualValue(ByRef rng as Range)     ' OK
    Dim startCell As Range, endCell as Range

    Set startCell = rng.item(1, 1)

    For Each cell In rng
        If Not (startCell.Value = cell.value) Then
            Call Merge(Range(startCell, endCell))
            Set startCell = cell
        End If
        
        Set endCell = cell
    Next cell 

    if Not (Range(startCell, endCell).MergeCells) Then
        Call Merge(Range(startCell, endCell))
    End if

End Sub

Public  Sub FunctionToString(ByRef rng as Range, Optional func as String)
    Dim cell As Range

    For Each cell In rng
        If func = "" Or InStr(cell.Formula, func) <> 0 Then
            cell.value = cell.value
        End If
    Next cell
End Sub


Public Function RankIndex(ByVal rng As Range, rank As Integer, Optional sort As Integer) As Range

    Dim cell As Range

    If (sort <> 0) Then
        sort = 1
    End If

    For Each cell In rng
        If (Application.rank(cell, rng, sort) = rank) Then
            Set RankIndex = cell
            Exit Function
        End If
    Next cell 

End Function


Public  Function RankIndexColumn(ByVal rng As Range, columnAddress As Range, rank As Integer, Optional sort As Integer) As String

    RankIndexColumn = Cells(RankIndex(rng, rank, sort).row, columnAddress.column).value

End Function

Public  Function RankIndexRow(ByVal rng As Range, rowAddress As Range, rank As Integer, Optional sort As Integer) As String

    RankIndexRow = Cells(rowAddress.row, RankIndex(rng, rank, sort).column).value

End Function

Public  Function indexColumn(ByVal rng As Range, columnAddress As Range, index As Integer) As String

    indexColumn = Cells(rng.item(index, 1).row, columnAddress.column).value

End Function

Public  Function indexRow(ByVal rng As Range, rowAddress As Range, index As Integer) As String

    indexRow = Cells(rowAddress.row, rng.item(1, index).column).value

End Function

public sub divideRange(ByVal rng as Range, ByVal Optional div as Long)
    dim cell as Range
    dim i as Integer
    
    if div = 0 then
        div = 1
    end if
    
    for each cell in rng
        if Not isFormulaInCell(cell) then

            if (cell.value = 0 or not IsNumeric(cell.value)) then
                cell.value = 0
            else
                cell.value = cell.value / div
            end if

        end if
    next cell
end sub


public function isFormulaInCell(rng as Range) as Boolean
    if (InStr(rng.Formula, "=") <> 0) then
        isFormulaInCell = True
    else
        isFormulaInCell = False
    end if
end function