Attribute VB_Name = "PlotStyling"
Option Explicit

Sub StylePlot()
'

'
    Dim i As Integer
    Dim n As Integer
    Dim lineStyle As Variant
    Dim val As Double
    Dim obj As Object
    On Error Resume Next

    Selection.Format.Line.Visible = msoFalse
    With ActiveChart.Parent
         .Height = 210 ' resize
         .Width = 230  ' resize
         '.Top = 150    ' reposition
         '.Left = 400   ' reposition
    End With
    
    lineStyle = Array(msoLineSolid, msoLineSysDash, msoLineSysDashDot, msoLineSysDot, msoLineDash, msoLineLongDashDot, msoLineDashDotDot, msoLineDashDot, msoLineLongDashDotDot)
    Application.CommandBars("Format Object").Visible = False
    For i = 1 To ActiveChart.FullSeriesCollection.Count
        ActiveChart.FullSeriesCollection(i).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .DashStyle = lineStyle(i)
        End With
    Next i
    
    'Formatting Axes
    With ActiveChart
        .HasTitle = False
        'X axis name
        If .Axes(xlCategory, xlPrimary).HasTitle = False Then
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "X-Axis"
        End If
        If .Axes(xlValue, xlPrimary).HasTitle = False Then
        'y-axis name
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Y-Axis"
        End If
    End With
    
    'Formatting x-axis
    With ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Font
        .Name = "Times New Roman"
        .Size = 8
        .Color = msoThemeColorText2
    End With
    
    With ActiveChart.Axes(xlCategory).TickLabels.Font
        .Name = "Times New Roman"
        .Size = 8
        .Color = msoThemeColorText2
    End With
    
    'formatting y-axis

    With ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Font
        .Name = "Times New Roman"
        .Size = 8
        .Color = msoThemeColorText2
    End With
    
    With ActiveChart.Axes(xlValue, xlPrimary).TickLabels.Font
        .Name = "Times New Roman"
        .Size = 8
        .Color = msoThemeColorText2
    End With
    
    With ActiveChart.ChartArea.Format.TextFrame2.TextRange.Font
        .Name = "Times New Roman"
        .Size = 8
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
    End With
    
    
    ActiveChart.Legend.Position = xlBottom
    
    
    'Remove Gridlines
    ActiveChart.Axes(xlValue).MajorGridlines.Delete
    ActiveChart.Axes(xlCategory).MajorGridlines.Delete
    
    'Add tickmarks inside
    ActiveChart.Axes(xlValue).MajorTickMark = xlInside
    ActiveChart.Axes(xlCategory).MajorTickMark = xlInside
    
    'plotarea border
    With ActiveChart.PlotArea.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    With ActiveChart.Axes(xlValue).Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    With ActiveChart.Axes(xlCategory).Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
    
End Sub

Sub SetColorSeries()
    Dim val As Double
    Dim obj As Object
    val = 1#
    For Each obj In ActiveChart.FullSeriesCollection
        val = val / 2
        With obj.Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = val
            .Solid
        End With
    Next
End Sub

