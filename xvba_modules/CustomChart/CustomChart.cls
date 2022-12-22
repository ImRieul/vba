
VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "CustomChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private self as Chart

' XAxis : Group


Public  Sub setChart(ByVal chart As Chart)      ' OK

    If Not (self Is Nothing) Then
        Exit Sub
    End If

    Set self = chart

End Sub

public function getName() as String
    getChartName = self.Parent.Name
End Function

public sub setName (ByVal name as String)
    self.Parent.Name = name
End Sub

Public  Function getChart() As Chart        ' OK
    Set getChart = self
End Function

public function getHeight() as Double
    getHeight = self.Parent.Height
End Function

public function getWidth() as Double
    getWidth = self.Parent.Width
End Function

public sub setHeight(ByVal height as Double)
    self.Parent.Height = height
End Sub

public sub setWidth(ByVal width as Double)
    self.Parent.Width = width
End Sub

public sub setTemplate(ByVal TemplateName as String)
    self.ApplyChartTemplate TemplateName
End Sub

public function getXAxisCount() as Long
    getXAxisCount = self.SeriesCollection.Count
End Function

public function setValuesName(index As long, name as String)
    self.FullSeriesCollection(index).Name = name
End Function

public function getValuesName(index As long) as String
    if index > self.SeriesCollection.Count then
        getXAxisName = ""
        Exit Function
    End If

    getXAxisName = self.FullSeriesCollection(index).Name
End Function

public sub setGroup(ByVal index as Long, ByVal rng as Range)
    ' set Chart Group
    ' index : Group Index
    ' rng : Range

    self.FullSeriesCollection(index).XValues = rng
End Sub

public sub setValues(ByVal index as Long, ByVal rng as Range)
    ' set Chart Value
    ' index : Value Index
    ' rng : Range

    self.FullSeriesCollection(index).Values = rng
End Sub

public sub chartTitleVisible(ByVal visible as Boolean)

    if (visible) then
        self.SetElement (mosElementChartTitleNone)
    else
        self.SetElement (mosElementChartTitleAboveChart)
    End If

End Sub

private sub createSeriesCollection(Optional count As Integer) 

    if count = 0 then
        count = 1
    End If

    for i = 1 to count
        self.SeriesCollection.NewSeries
    Next i

End Sub

public sub move(ByVal rng As Range, Optional sh As Worksheet)

    if sh Is Nothing then
        Set sh = self.Application.ActiveSheet
    End If

    self.Parent.Cut
    sh.Paste Destination:=sh.Range(rng.Address)

end sub