Attribute VB_Name = "chart_main"

public sub charttest()

    Dim ch As CustomChart
    Dim templatePath As String

    templatePath = chartTools.getTemplatePath("chart_2019~2021_age.crtx")

    set ch = chartTools.CustomChartConstructor("parientChart")

    ch.setTemplate(templatePath)
    
end sub

public sub testCreateChart() 

    Dim ch As customChart
    Dim range2019, range2020, range2021 as Range
    Dim templatePath As String
    Dim height, width As Double
    Dim userName, age As String

    userName = ""
    age = ""

    templatePath = chartTools.getTemplatePath(userName, "chart_2019~2021_age.crtx")

    height = 200.591888427734
    width = 511.01708984375

    Set ch = chartTools.createChart

    set range2019 = Union(Range("E9:P9"), Range("E13:P13"), Range("E17:P17"))
    set range2020 = Union(Range("E10:P10"), Range("E14:P14"), Range("E18:P18"))
    set range2021 = Union(Range("E11:P11"), Range("E15:P15"), Range("E19:P19"))

    ch.getChart.SeriesCollection.NewSeries
    ch.getChart.SeriesCollection.NewSeries
    ch.getChart.SeriesCollection.NewSeries

    ch.chartTitleVisible(False)

    call ch.setName("parientChart")
    call ch.setGroup(1, Range("A73:AJ74"))

    call ch.setValues(1, range2019)
    call ch.setValues(2, range2020)
    call ch.setValues(3, range2021)

    ch.setTemplate(templatePath)
    ch.setHeight(height)
    ch.setWidth(width)

    call ch.setValuesName(1, "20-29" & age)
    call ch.setvaluesName(2, "30-39" & age)
    call ch.setvaluesName(3, "40-49" & age)

    call ch.move(Range("A58"))
    Range("A58").Select

end sub


public sub testCreateCostChart() 

    Dim ch As customChart
    Dim range2019, range2020, range2021 as Range
    Dim templatePath As String
    Dim height, width As Double
    Dim userName, age As String

    userName = ""
    age = ""

    templatePath = chartTools.getTemplatePath(userName, "chart_2019~2021_age.crtx")

    height = 200.591888427734
    width = 511.01708984375

    Set ch = chartTools.createChart

    set range2019 = Union(Range("E37:P37"), Range("E41:P41"), Range("E45:P45"))
    set range2020 = Union(Range("E38:P38"), Range("E42:P42"), Range("E46:P46"))
    set range2021 = Union(Range("E39:P39"), Range("E43:P43"), Range("E47:P47"))

    ch.getChart.SeriesCollection.NewSeries
    ch.getChart.SeriesCollection.NewSeries
    ch.getChart.SeriesCollection.NewSeries

    ch.chartTitleVisible(False)

    call ch.setName("costChart")
    call ch.setGroup(1, Range("A73:AJ74"))

    call ch.setValues(1, range2019)
    call ch.setValues(2, range2020)
    call ch.setValues(3, range2021)

    ch.setTemplate(templatePath)
    ch.setHeight(height)
    ch.setWidth(width)

    call ch.setValuesName(1, "20-29" & age)
    call ch.setvaluesName(2, "30-39" & age)
    call ch.setvaluesName(3, "40-49" & age)

    call ch.move(Range("K58"))
    Range("K58").Select

end sub
