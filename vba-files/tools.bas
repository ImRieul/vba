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


public function isFormulaInCell(rng as Range) as Boolean
    if (InStr(rng.Formula, "=") <> 0) then
        isFormulaInCell = True
    else
        isFormulaInCell = False
    end if
end function


Public sub paintBetweenRow(ByVal rng As Range, between As Integer, Optional rowEnd as Long)

    dim rowStart as Long
    dim row as Long

    if (rowEnd = 0) then
        rowEnd = cells(rows.count, rng.column).end(xlup).row
        rowStart = rng.row
    elseif (rng.row > rowEnd) then
        rowStart = rowEnd
        rowEnd = rng.row
    else
        rowStart = rng.row
    end if
        
    for row = rowStart to rowEnd step between
        cells(row, rng.column).interior.color = vbyellow
    next row

End sub


public function compareThreeValues(val1, val2, val3 as Long) as String

    upup = "매년 상승 추세"
    upgo = ""
    updown = ""

    goup = ""
    gogo = ""
    godown = ""

    downup = ""
    downgo = ""
    downdown = ""

    If (val1 > val2) Then 

        If (val2 > val3) Then
            compareThreeValues = downdown
        Elseif (val2 = val3) Then
            compareThreeValues = downgo
        else    ' val2 < val3
            compareThreeValues = downup
        End If

    Elseif (val1 = val2) Then

        If (val2 > val3) Then
            compareThreeValues = godown
        Elseif (val2 = val3) Then
            compareThreeValues = gogo
        else    ' val2 < val3
            compareThreeValues = goup
        End If

    else    ' val1 < val2

        If (val2 > val3) Then
            compareThreeValues = updown
        Elseif (val2 = val3) Then
            compareThreeValues = upgo
        else    ' val2 < val3
            compareThreeValues = upup
        End If
    
    End If

end function

Public  Function compareTwoValues(val1, val2 as Long) as String

    up = "많음"
    go = ""
    down = "적음"

    If (val1 > val2) Then 
        compareTwoValues = down
    Elseif (val1 = val2) Then
        compareTwoValues = go
    else    ' val1 < val2
        compareTwoValues = up
    End If

End Function

Public  Sub resetFormula(rng as Range)

    For Each r In rng

        if r.hasformula then
            r.value = r.formula
        end if
    
    Next r 

End Sub