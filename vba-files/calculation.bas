Attribute VB_Name = "calculation"

Public  Sub DivideRange (ByRef ran As Range, ByVal divisor As Double)

    Dim cell As Range

    For Each cell In ran
        
        If (cell.value = "") Then
            cell.value = 0
        ElseIf IsNumeric(cell.Value) and cell.value <> 0 Then
            cell.Value = cell.Value / divisor
        End If
    
    Next cell 

End Sub