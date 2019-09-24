Option Explicit
Public worksheetName As String

Sub drawChart()
    Dim worksheetCount As Integer
    Dim item As Integer

    With Application
        .ScreenUpdating = False
        .DisplayFullScreen = True
        .DisplayFormulaBar = False
    End With

    worksheetName = ActiveSheet.Name
    sheets("Plán_kod_jmeno").Visible = True
    sheets("Plán_kod_jmeno").Activate
    ActiveSheet.Name = "Graf_" & worksheetName
    worksheetCount = ActiveWorkbook.Worksheets.Count

    For item = 1 To worksheetCount
       If Worksheets(item).Name = "Plán_" & worksheetName Then
       Else
       Worksheets(item).Visible = False
       End If
    Next item
    
    With ActiveChart
        .ChartTitle.Text = "Plán a skutečnost k aktuálnímu dni pro MOS " & worksheetName
        .SetSourceData Source:=sheets(worksheetName).Range( _
            "B9:D29")
        .SeriesCollection(1).Name = "='" & worksheetName & "'!$C$8"
        .SeriesCollection(1).Values = "='" & worksheetName & "'!$C$9:$C$29"
        .SeriesCollection(1).XValues = "='" & worksheetName & "'!$B$9:$B$29"
        .SeriesCollection.NewSeries
        .SeriesCollection(2).ChartType = xlColumnClustered
        .SeriesCollection(2).Name = "='" & worksheetName & "'!$D$8"
        .SeriesCollection(2).Values = "='" & worksheetName & "'!$D$9:$D$29"
        .ChartGroups(1).GapWidth = 0
        .SeriesCollection(2).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .SeriesCollection(2).ApplyDataLabels
    End With

    ActiveChart.SeriesCollection(2).DataLabels.Select
    Selection.Position = xlLabelPositionOutsideEnd
    Selection.NumberFormat = "# ##0"

    With Selection.Format.TextFrame2.TextRange.Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorDark1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
    Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    ActiveChart.Activate
    ActiveSheet.Protect "pass"
    Worksheets("Pořadí").Activate
    
    Application.ScreenUpdating = True
End Sub

Sub hideChart()
    Dim worksheetCount As Integer
    Dim item As Integer

    Application.ScreenUpdating = False
    ActiveSheet.Name = "Plán_kod_jmeno"
    ActiveSheet.Unprotect "pass"
    worksheetCount = ActiveWorkbook.Worksheets.Count

    For item = 7 To worksheetCount
       If Worksheets(item).Name = "Plán_kod_jmeno" Then
           Worksheets(item).Visible = False
       Else
           Worksheets(item).Visible = True
       End If
    Next item

    sheets("Plán_kod_jmeno").Visible = False
    sheets(worksheetName).Activate

    With Application
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .ScreenUpdating = True
    End With

End Sub
