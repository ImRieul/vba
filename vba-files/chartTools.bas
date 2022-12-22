Attribute VB_Name = "chartTools"

Public Sub SetChartScaleHigh(ByRef request as chart, ByVal max As Integer)       ' OK
    request.Axes(xlValue).MaximumScale = max
End Sub

Public Sub SetChartScaleLow(ByRef request as chart, ByVal min As Integer)        ' OK
    request.Axes(xlValue).MinimumScale = min
End Sub

Public  Sub SetChartScaleAuto(ByRef request as chart)      ' OK
    request.Axes(xlValue).MinimumScaleIsAuto = True
    request.Axes(xlValue).MaximumScaleIsAuto = True 
End Sub

' not test down ~
Public Sub SetChartScaleInterval(ch as chart, size as Integer)
    ch.Axes(xlValue).majorunit = size
End Sub

'----------------------------------------------------------------------------------------------------------------


Public Sub newChartCustom(ByVal ran As range, Optional chartName As String, Optional ByVal h As Double, Optional ByVal w As Double, Optional ByVal pointRange As range, Optional ByVal sh As Worksheet)

    dim userName, year2019, year2020, year2021 As String

    If (sh is nothing) Then
        set sh = ActiveSheet
    End If

    sh.Select

    userName = ""
    year2019 = "=""value"""
    year2020 = "=""value"""
    year2021 = "=""value"""

    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select

    With activechart

        ' Get Chart Template
        .ApplyChartTemplate ( _
            "C:\Users\" & userName & "\AppData\Roaming\Microsoft\Templates\Charts\chart_2019~2021.crtx")
        .SetSourceData Source:=range("'" & sh.name & "'!" & ran.Address)


        ' set X Title Rename
        
        ' SetChartXAxisName request:=ActiveChart index:=1 name:=year2019
        ' SetChartXAxisName request:=ActiveChart index:=2 name:=year2020
        ' SetChartXAxisName request:=ActiveChart index:=3 name:=year2021

        .FullSeriesCollection(1).Name = year2019
        .FullSeriesCollection(2).Name = year2020
        .FullSeriesCollection(3).Name = year2021

        ' chart move
        if not pointRange is Nothing then
            .Parent.Cut
            sh.Paste Destination:=sh.Range(pointRange.Address)
        end if

        call SetChartScaleAuto(ActiveChart)

        ' set height
        If h <> 0 Then
            '.Parent.height = h
        End If

        ' set width
        If w <> 0 Then
            '.Parent.width = w
        End If

        ' set Name
        if chartName <> "" then
            'sh.Shapes(.Name).Name = chartName
        end if

    End With


End Sub

Public  Function getXAxisCount(ByVal request as chart) As Long      ' OK
    getXAxisCount = request.SeriesCollection.Count
End Function

Public  Function getXAxisName(ByVal request as chart, ByVal index As Long) As String        ' OK
    ' Get X axis name from chart.
    ' index : XAxis index.
    ' return : XAxis name. If index is bigger then chart axist count, return 'Error!

    ' Check index
    if index > getXAxisCount(request) then
        Dim errorMessage As String

        errorMessage = "Index is bigger then chart axist count." _
            & Chr(13) & "ChartName : " & getChartName request:=request _
            & Chr(13) & "Index : " & index

        msgbox(errorMessage)
        getXAxisName = "Error!"
    Else 

        getXAxisName = request.SeriesCollection(index).Name
    end if

End Function

public function getTemplatePath(ByVal userName, templateName as String) as string
    getTemplatePath = "C:\Users\" & userName & "\AppData\Roaming\Microsoft\Templates\Charts\" & templateName
end function


public function createChart() as CustomChart
    Dim customCh as CustomChart

    set customCh = new CustomChart
    customCh.setChart(ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Chart)

    Set createChart = customCh
end function

Public  Function CustomChartConstructor(ByVal name As String, Optional ByVal sheet As Worksheet) as CustomChart         ' OK
    ' Constructor of CustomChart
    ' name : Chart Name
    ' sheet : Chart in Sheet
    ' Return CustomChart

    If Not ExistChart(name, sheet) Then
        MsgBox "Not Exist Chart" & _
                Chr(13) & "name : " & name & _
                Chr(13) & "sheet : " & sheet
        
        CustomChartConstructor = NoThing
    End If

    Dim result As new CustomChart
    Dim chart As Chart

    Set chart = getChart(name, sheet)

    result.setChart chart:=chart

    Set CustomChartConstructor = result

End Function    

Public  Function getChart(ByVal name As String, Optional ByVal sheet As Worksheet) as Chart         ' OK
    ' Get Chart by Name
    ' name : Chart Name
    ' sheet : Chart in Sheet
    ' Return Chart

    if sheet Is Nothing Then
        Set sheet = ActiveSheet
    End If

    If Not (ExistChart(name, sheet)) Then
        getChart = NoThing
    End If

    Set getChart = sheet.ChartObjects(name).Chart

End Function

Public  Function ExistChart(ByVal name As String, Optional ByVal sheet As Worksheet) As Boolean     ' OK
    ' Check Exist Chart
    ' name : Chart Name
    ' sheet : Chart in Sheet
    ' Return True or False

    Dim result As Boolean

    if sheet Is Nothing Then
        Set sheet = ActiveSheet
    End If

    result = False

    For Each chart In sheet.ChartObjects

        If (chart.Chart.Parent.Name = name) Then
            result = True
        End If
    
    Next chart 
    
    ExistChart = result

End Function