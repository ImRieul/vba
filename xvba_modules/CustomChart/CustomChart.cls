
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

Public  Sub setChart(ByVal chart As Chart)      ' OK

    If Not (self Is Nothing) Then
        Exit Sub
    End If

    Set self = chart

End Sub

Public  Function getChart() As Chart        ' OK
    Set getChart = self
End Function