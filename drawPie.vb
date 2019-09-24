Option Explicit
Sub drawPie()
    With Application
        .ScreenUpdating = False
        .DisplayFullScreen = True
        .DisplayFormulaBar = False
    End With

    Dim offsetData As Integer
    Dim dataArray() As Variant
    Dim worksheetCount As Integer
    Dim item As Integer

    worksheetName = ActiveSheet.Name
    sheets("Produkty_kod_jmeno").Visible = True
    sheets("Produkty_kod_jmeno").Activate
    ActiveSheet.Name = "Produkty_" & worksheetName
    ActiveSheet.Tab.Color = 255
    worksheetCount = ActiveWorkbook.Worksheets.Count

    For item = 1 To worksheetCount
       If Worksheets(item).Name = "Produkty_" & worksheetName Then
       Else
       Worksheets(item).Visible = False
       End If
    Next item

    ActiveWorkbook.Worksheets(worksheetName).Activate
    ActiveSheet.Range("I36", "I62").End(xlDown).Activate

    ReDim dataArray(0 To 3)
    For offsetData = 0 To 3
        If ActiveCell.Value = Empty Then
            dataArray(offsetData) = 0
            ActiveCell.Offset(28, 0).Activate
        Else
            dataArray(offsetData) = ActiveCell.Value
            ActiveCell.Offset(28, 0).Activate
        End If
    Next offsetData
    Range("A1").Activate
    Charts("Produkty_" & worksheetName).Activate
    ActiveChart.ChartArea.Select
    
    With ActiveChart
        .SeriesCollection(1).XValues = Array("A", "E", "F", "G")
        .SeriesCollection(1).Values = Array(CLng(dataArray(0)), CLng(dataArray(1)), CLng(dataArray(2)), CLng(dataArray(3)))
        .SeriesCollection(1).Name = "Podíl produktů k aktuálnímu dni pro MOS " & worksheetName
    End With
    
    ActiveSheet.Protect "pass"
    Worksheets("Pořadí").Activate
    Application.ScreenUpdating = True
End Sub

Sub hidePie()
    Dim worksheetCount As Integer
    Dim item As Integer
    Application.ScreenUpdating = False
    ActiveSheet.Name = "Produkty_kod_jmeno"
    ActiveSheet.Unprotect "pass"
    worksheetCount = ActiveWorkbook.Worksheets.Count

    For item = 7 To worksheetCount
       If Worksheets(item).Name = "Produkty_kod_jmeno" Then
           Worksheets(item).Visible = False
       Else
           Worksheets(item).Visible = True
       End If
    Next item

    sheets("Produkty_kod_jmeno").Visible = False
    sheets(worksheetName).Activate

    With Application
        .DisplayFullScreen = False
        .DisplayFormulaBar = True
        .ScreenUpdating = True
    End With
End Sub

