Attribute VB_Name = "main"

Public Sub SetChartXY()

    Dim patient as ChartObject
    Dim cost as ChartObject

    Dim costMax As Integer
    Dim costMin As Integer

    Dim patientMin As Integer
    Dim patientMax As Integer

    Dim endRow As Integer
    
    endRow = cells(rows.count, 3).end(xlup).row

    set patient = activesheet.chartObjects("patientChart")
    set cost = ActiveSheet.chartObjects("costChart")

    costMin = Application.Round(RangeMin(Range(cells(3, 10), cells(endRow, 10))), -2)
    costMax = Application.Round(RangeMax(Range(cells(3, 10), cells(endRow, 10))), -2)

    patientMin = Application.Round(RangeMin(Range(cells(3, 6), cells(endRow, 6))), -2)
    patientMax = Application.Round(RangeMax(Range(cells(3, 6), cells(endRow, 6))), -2)
    
    patient.Activate
    call ChangeChartRange(ActiveChart, activeSheet.name)
    call chartTools.SetChartScaleHigh(ActiveChart, patientMax)
    call chartTools.SetChartScaleLow(ActiveChart, patientMin)

    cost.Activate
    call chartTools.SetChartScaleHigh(ActiveChart, costMax)
    call chartTools.SetChartScaleLow(ActiveChart, costMin)

    'patient.fullSeriesCollection(1).xvalues = _

End Sub

Public Sub ChangeChartRange(ch as chart, sheetName as String)
    
    Dim xStr As String
    Dim yStr2019 As String
    Dim yStr2020 As String
    Dim yStr2021 As String
    
    xStr = "='" + sheetName + "'!" + range(cells(3, 3), cells(14, 3)).address
    yStr2019 = "='" + sheetName + "'!" + range("F3", "F14").address
    ystr2020 = "='" + sheetName + "'!" + range("F15", "F26").address
    yStr2021 = "='" + sheetName + "'!" + range("F27", "F38").address

    ch.FullSeriesCollection(1).XValues = xStr
    ch.FullSeriesCollection(1).Values = yStr2019
    ch.FullSeriesCollection(2).Values = ystr2020
    ch.FullSeriesCollection(3).Values = yStr2021


End Sub


Public Sub PasteChart() 

    Dim patientTag As String
    Dim costTag As String
    Dim sheetName As String
    

    patientTag = worksheets(1).cells(5, "L")
    costTag = worksheets(1).cells(20, "L")

    sheetName = replace(worksheets(1).name, "-", " ")

    for sheetCount = 2 to sheets.count


        With worksheets(sheetCount)


            .chartObjects("patientChart").delete
            .chartObjects("costChart").delete

            worksheets(1).range("L5", "S35").copy
            .Paste Destination:=.range("L5", "S35")
            
            .cells(5, "L") = replace(patientTag, sheetName, _
                .cells(3, 1) + " " + .cells(3, 5) _
            )            

            .cells(20, "L") = replace(patientTag, sheetName, _
                .cells(3, 1) + " " + .cells(3, 5) _
            )

        End With

    next
    

End Sub

Public  Sub UseAfterDelete()


    Dim defaultChart As Chart
    Dim h, w As Double

    With Worksheets(4).Shapes("patientChart")

        h = .height
        w = .width
        
    End With

    With ActiveChart

        .parent.height = h
        .parent.width = w

    End With

End Sub


sub editrange()

    Dim nowRange as Range

    Set nowRange = Selection

    call calculation.DivideRange(nowRange, 1000000)
    nowRange.NumberFormatLocal = "#,##0.0"

end sub

public sub mergeWithShortcut()

    if (Selection.MergeCells) then
        call tools.UnMergePull(Selection)
    else
        call tools.Merge(Selection)
    end if

End Sub

Public  Sub mergeGroup()
    call tools.MergeEqualValue(Selection)
End Sub

Public  Sub costFormat()
    call formatter.AccountingInRange(Selection, "#,##0.0")
End Sub

Public  Sub parientFormat()
    call formatter.AccountingInRange(Selection, "#,##0")
End Sub

Public  Sub percentFormat()
    call formatter.PercentInRange(Selection)
End Sub

Public  Sub dataToSrting()
    call tools.FunctionToString(Selection, "SUMIFS")
End Sub

public sub tableDivision()
    Dim THOUSAND, MILLION, HUNDRED_MILLION as Long
    dim box As String

    THOUSAND = 1000
    MILLION = 1000000
    HUNDRED_MILLION = 100000000

    box = UCase(inputbox("please input a divisible number" & chr(13) & "thousand or million or hundred million"))

    if (box = "THOUSAND") then
        call calculation.divideRange(Selection, THOUSAND)
    elseif (box = "MILLION") then
        call calculation.divideRange(Selection, MILLION)
    elseif (box = "HUNDRED MILLION") then
        call calculation.divideRange(Selection, HUNDRED_MILLION)
    else
        msgbox box & " is not exist option."
        exit sub
    end if
end sub

public sub tablePullEmptyCell()
    call calculation.DivideRange(Selection, 1)
end sub

public sub checkCellUseColor()

    dim WHITE as Long

    WHITE = 16777215

    if (selection.interior.color = WHITE) then
        selection.interior.color = vbyellow
    else
        selection.interior.color = xlNone
    end if
end sub

public sub selectFirstCell()

    dim thisSheet as worksheet
    dim i as long

    set thisSheet = activeSheet

    for i = 1 to sheets.count
        worksheets(i).select
        cells(1, 1).select
    next i

    thisSheet.select

end sub

public sub paintBetweenColorYellow()
    
    dim between as integer
    dim rowEnd as Long

    dim betweenInput, rowEndInput as String

    betweenInput = inputbox("pleesae between number")
    rowEndInput = inputbox("please row end")

    if not (isNumeric(betweenInput)) then
        between = 0
    else
        between = Val(betweenInput)
    end if

    if not (isNumeric(rowEndInput)) then
        rowEnd = 0
    else
        rowEnd = Val(rowEndInput)
    end if

    call tools.paintBetweenRow(Selection, between)
end sub

public sub resetFunction()
    call tools.resetFormula(Selection)
end sub