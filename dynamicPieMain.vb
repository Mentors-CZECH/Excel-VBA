'skript upravit tak, aby se listy načítaly ořezané do pole pouze jedno, při stuštění souboru

'Option Explicit

Sub Worksheet_SelectionChange(ByVal Target2 As Excel.Range)
    Application.ScreenUpdating = False
    Dim VRange2 As Range
    Dim salesChart As Chart
    Dim wks As Worksheet
    Dim WS_Count As Integer
    Dim sheets, Characters, offsetData As Integer
    Dim userName As String
    Dim sheetID, targetID, result As String
    Dim LastLine As Long
    Dim dataArray() As Variant
    Dim myChartObj As Object
    Dim row, col As Variant
    Dim pieChart As Chart
    Dim seriesData As Series
    Dim test As String

    If ActiveSheet.Name = "Pořadí" Then
        For Each wks In Worksheets
            If wks.ChartObjects.Count > 0 Then
               wks.ChartObjects.Delete
            End If
        Next wks
        test = ActiveSheet.Range("C20").End(xlDown).Address
        Set VRange2 = Range("C3:" & test)
        If Union(Target2, VRange2).Address = VRange2.Address Then
            If ActiveCell = Empty Then
                For Each wks In Worksheets
                    If wks.ChartObjects.Count > 0 Then
                        wks.ChartObjects.Delete
                    End If
                Next wks
			Else
                For Each wks In Worksheets
                    If wks.ChartObjects.Count > 0 Then
                        wks.ChartObjects.Delete
                    End If
                Next wks
                userName = ActiveCell.Offset(0, 2).Value

                ActiveCell.Offset(0, 0).Select
            End If

            WS_Count = ActiveWorkbook.Worksheets.Count
            sheetID = ActiveCell.Value

            For sheets = 8 To WS_Count
                targetID = ActiveWorkbook.Worksheets(sheets).Name
                Characters = InStr(1, targetID, "-") - 2
                If Characters < 0 Then
                Else
                    result = Left(targetID, Characters)
                End If
                If result = sheetID Then
                    'pokud je nalezen list, načti data do pole, vybreslit graf a ukonči smyčku
                    ActiveWorkbook.Worksheets(sheets).Activate
                    ActiveSheet.Range("I36", "I62").End(xlDown).Activate

                    ReDim dataArray(0 To 3)
                    'Načte data produktů daného ÚP do pole myData
                    For offsetData = 0 To 3
                        If ActiveCell.Value = Empty Then
                            dataArray(offsetData) = 0
                            ActiveCell.Offset(28, 0).Activate
                        Else
                            dataArray(offsetData) = ActiveCell.Value
                            ActiveCell.Offset(28, 0).Activate
                        End If
                    Next offsetData

                    ActiveWorkbook.Worksheets("Pořadí").Activate
                    row = ActiveCell.row
                    col = ActiveCell.Column
                    Set myChartObj = ActiveSheet.ChartObjects.Add _
                            (Left:=100, Width:=375, Top:=75, Height:=225)
                        myChartObj.Name = "Chart 1"
                    Set pieChart = ActiveSheet.ChartObjects("Chart 1").Chart
                    Set chartPos = ActiveSheet.Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(20, 5))

                    With pieChart
                    .ChartType = xlPie
                    .SeriesCollection.NewSeries
                    .SeriesCollection(1).XValues = Array("A", "E", "F", "G")
                    .SeriesCollection(1).Values = Array(CLng(dataArray(0)), CLng(dataArray(1)), CLng(dataArray(2)), CLng(dataArray(3)))
                    .SeriesCollection(1).ApplyDataLabels
                    .SeriesCollection(1).DataLabels.Select
                        Selection.Position = xlLabelPositionOutsideEnd
                        Selection.ShowPercentage = True
                        Selection.ShowValue = False
                    .SeriesCollection(1).Name = "Podíl produktů MOS " & userName
                    .SeriesCollection(1).DataLabels.Select
                        Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue

                    End With
                    myChartObj.Left = chartPos.Left
                    myChartObj.Width = chartPos.Width
                    myChartObj.Top = chartPos.Top
                    myChartObj.Height = chartPos.Height
                    ActiveChart.ChartArea.Select
					ActiveChart.ChartArea.Select
					ActiveChart.SeriesCollection(1).Select
					ActiveChart.ClearToMatchStyle
					ActiveChart.ChartStyle = 18
					ActiveChart.ClearToMatchStyle
					With Selection.Format.Line
						.Visible = msoTrue
						.ForeColor.ObjectThemeColor = msoThemeColorAccent1
						.ForeColor.TintAndShade = 0
						.ForeColor.Brightness = 0
					End With
					With Selection.Format.Line
						.Visible = msoTrue
						.ForeColor.ObjectThemeColor = msoThemeColorBackground1
						.ForeColor.TintAndShade = 0
						.ForeColor.Brightness = 0
						.Transparency = 0
					End With
					With Selection.Format.Line
						.Visible = msoTrue
						.Weight = 1.25
					End With
					Worksheets("Pořadí").Activate
                    Exit For
                End If
             Next sheets
        End If
    End If
    Application.ScreenUpdating = True
End Sub

