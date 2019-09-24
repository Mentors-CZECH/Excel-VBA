'Option Explicit

'Private Sub Worksheet_Activate()
'        Call Rebuild_Months
''        With Sheets("Dashboard")
''                .EnableSelection = xlNoSelection
''                .Protect , , , , True
''        End With
'End Sub

Private Sub monthComboBox_Change()
        Range("Y18").Value = monthComboBox
        Call monthProductComboBox_Change
        Call drawMonthProductDistribution
        Call drawProductResults
        
End Sub

Private Sub drawProductResults()

End Sub

Private Sub drawMonthProductDistribution()

        yearStartRow = Sheets("RawData").Range("D:D"). _
        Find(What:=monthComboBox.Value, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Row
        
        Sheets("Dashboard").ChartObjects("monthProdDistr").Chart.FullSeriesCollection(1).Name = "=RawData!$D$" & yearStartRow
        Sheets("Dashboard").ChartObjects("monthProdDistr").Chart.FullSeriesCollection(1).Values = "=RawData!$F$" & yearStartRow & ":$J$" & yearStartRow

End Sub

Private Sub monthProductComboBox_Change()
        'rebuild graph
        
        Dim yearStartRow As Integer
        Dim numberOfItems As Integer
        Dim columnArray As Variant
        
        columnArray = Array("E", "F", "G", "H", "I", "J")
        
        'najít první záznam roku startYearCombo
        Sheets("RawData").Activate

        yearStartRow = Sheets("RawData").Range("D:D"). _
        Find(What:=monthComboBox.Value, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Row
        
        If monthProductComboBox = "A" Then
                startColumn = "K"
                endColumn = "AO"
        ElseIf monthProductComboBox = "B" Then
                startColumn = "AP"
                endColumn = "BT"
        ElseIf monthProductComboBox = "C" Then
                startColumn = "BU"
                endColumn = "CY"
        ElseIf monthProductComboBox = "D" Then
                startColumn = "CZ"
                endColumn = "ED"
        ElseIf monthProductComboBox = "E" Then
                startColumn = "EE"
                endColumn = "FI"
        End If
        
        productType = monthProductComboBox.Value
        test = "=RawData!$" & startColumn & "$" & yearStartRow & ":$" & endColumn & "$" & yearStartRow
        
        Sheets("Dashboard").ChartObjects("monthDevGraph").Chart.FullSeriesCollection(1).Name = monthComboBox.Value
        Sheets("Dashboard").ChartObjects("monthDevGraph").Chart.SeriesCollection(1).Values = test
        Sheets("Dashboard").Activate
        
End Sub

Private Sub startYearCombo_Change()
        If startYearCombo > endYearCombo Then
                MsgBox ("špatně datum")
                If endYearCombo = "" Then
                        endYearCombo = endYearCombo.ListIndex = 0
                Else
                        startYearCombo = endYearCombo
                End If
        Else
                Sheets("Dashboard").Range("W3").Value = startYearCombo
                Call Rebuild_Months
        End If
        Call rebuildYearChart(startYearCombo, endYearCombo)
End Sub

Private Sub endYearCombo_Change()
        If startYearCombo > endYearCombo Then
                MsgBox ("špatně datum")
                endYearCombo = startYearCombo
        Else
                Sheets("Dashboard").Range("Y3").Value = endYearCombo
                Call Rebuild_Months
        End If
        Call rebuildYearChart(startYearCombo, endYearCombo)
End Sub

Sub rebuildYearChart(startYearCombo, endYearCombo)
        Dim yearStartRow As Integer
        Dim numberOfItems As Integer
        Dim columnArray As Variant
        
        columnArray = Array("E", "F", "G", "H", "I", "J")
        
        'najít první záznam roku startYearCombo
        Sheets("RawData").Activate

        yearStartRow = Sheets("RawData").Range("C:C"). _
        Find(What:=startYearCombo, _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Row
        
        'vypočítat rozdíl (ten určuje počet řádků, které budou načteny)
        numberOfItems = (endYearCombo - startYearCombo) + 1
        numberOfItems = numberOfItems * 12
        Sheets("RawData").Cells.EntireColumn.Hidden = False
        
        For i = 0 To UBound(columnArray)
                Sheets("Dashboard").ChartObjects("yearGraph").Chart.SeriesCollection((i + 1)).Values = "=RawData!$" & columnArray(i) & "$" & yearStartRow & ":$" & columnArray(i) & "$" & yearStartRow + numberOfItems - 1
                Sheets("Dashboard").ChartObjects("yearGraph").Chart.FullSeriesCollection(1).XValues = "=RawData!$D$" & yearStartRow & ":$D$" & yearStartRow + numberOfItems - 1
        Next i
        yearProductA.Value = True
        yearProductB.Value = True
        yearProductC.Value = True
        yearProductD.Value = True
        yearProductE.Value = True
        yearProductAll.Value = True
        Sheets("Dashboard").Activate

End Sub

Sub Rebuild_Months()
        Dim resultYears, buildNewArray As Integer
        Dim monthsArray As Variant
        Dim arrayItem As Variant
        Dim monthNameConc  As String

        resultYears = endYearCombo - startYearCombo
        monthsArray = Array("Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec")
        monthComboBox.Clear
    
        For buildNewArray = 0 To resultYears
                For Each arrayItem In monthsArray
                        If resultYears = 0 Then
                                monthNameConc = arrayItem & " - " & endYearCombo
                                Sheets("Dashboard").monthComboBox.AddItem monthNameConc
                        Else
                                monthNameConc = arrayItem & " - " & endYearCombo - buildNewArray
                                Sheets("Dashboard").monthComboBox.AddItem monthNameConc
                        End If
                Next arrayItem
        Next buildNewArray
End Sub

Sub rebuildChart()

End Sub

Private Sub yearProductAll_Click()
        With yearProductAll
                .BackColor = RGB(64, 64, 64)
                If yearProductAll.Value = True Then
                        .ForeColor = RGB(0, 112, 192)
                        Sheets("RawData").Range("E:E").EntireColumn.Hidden = False
                Else
                        .ForeColor = RGB(125, 125, 125)
                        Sheets("RawData").Range("E:E").EntireColumn.Hidden = True
                End If
        End With
End Sub

Private Sub yearProductA_Click()
        With yearProductA
                .BackColor = RGB(64, 64, 64)
                If yearProductA.Value = True Then
                        .ForeColor = RGB(0, 112, 192)
                        Sheets("RawData").Range("F:F").EntireColumn.Hidden = False
                Else
                        .ForeColor = RGB(125, 125, 125)
                        Sheets("RawData").Range("F:F").EntireColumn.Hidden = True
                End If
        End With
End Sub

Private Sub yearProductB_Click()
        With yearProductB
                .BackColor = RGB(64, 64, 64)
                If yearProductB.Value = True Then
                        .ForeColor = RGB(0, 112, 192)
                        Sheets("RawData").Range("G:G").EntireColumn.Hidden = False
                Else
                        .ForeColor = RGB(125, 125, 125)
                        Sheets("RawData").Range("G:G").EntireColumn.Hidden = True
                End If
        End With
End Sub

Private Sub yearProductC_Click()
        With yearProductC
                .BackColor = RGB(64, 64, 64)
                If yearProductC.Value = True Then
                        .ForeColor = RGB(0, 112, 192)
                        Sheets("RawData").Range("H:H").EntireColumn.Hidden = False
                Else
                        .ForeColor = RGB(125, 125, 125)
                        Sheets("RawData").Range("H:H").EntireColumn.Hidden = True
                End If
        End With
End Sub

Private Sub yearProductD_Click()
        With yearProductD
                .BackColor = RGB(64, 64, 64)
                If yearProductD.Value = True Then
                        .ForeColor = RGB(0, 112, 192)
                        Sheets("RawData").Range("I:I").EntireColumn.Hidden = False
                Else
                        .ForeColor = RGB(125, 125, 125)
                        Sheets("RawData").Range("I:I").EntireColumn.Hidden = True
                End If
        End With
End Sub

Private Sub yearProductE_Click()
        With yearProductE
                .BackColor = RGB(64, 64, 64)
                If yearProductE.Value = True Then
                        .ForeColor = RGB(0, 112, 192)
                        Sheets("RawData").Range("J:J").EntireColumn.Hidden = False
                Else
                        .ForeColor = RGB(125, 125, 125)
                        Sheets("RawData").Range("J:J").EntireColumn.Hidden = True
                End If
        End With
End Sub
