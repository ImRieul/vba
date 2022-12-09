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