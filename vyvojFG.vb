

Sub fetchData()
	Dim lastrowcolb, worksheetNum, rowNum, baseNumF, baseNumG, cPointer, fPointer, gPointer, sPointer, test As Integer
	Dim planData, realData, requestPlan, requestReal, countryPlaned, countryReal As Double
	Dim setMonth, setYear, thisMonth, thisYear, mPointer, newSheetName, datum, S, targetWb As String
	Dim productsArray(1 To 4), Arr As Variant
	Dim sorceFileDialog As Object
	Dim sourceFilePath As String

	Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

	With Application
        .DisplayAlerts = False
        .PrintCommunication = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With

	thisWbName = ActiveWorkbook.Name
	
	With sourceFileDialog
		.Title = "Zadejte cestu k souboru: Plán a skutečnost v regionech (MMMM YYYY)"
		.Filters.Clear
		.Filters.Add "Soubory MS Excel", "*.xl*"
		.ButtonName = "Načíst data"
		.AllowMultiSelect = False
		'Pokud uživatel klikne na storno, zobrzit promt - nebyly vyplněny listy
		If .Show <> -1 Then
			msgVal = MsgBox("Storno", vbCritical)
			Exit Sub
		End If
		'Přiřazení cesty k souboru reports SD salesreport do proměnné
		sourceFilePath = .SelectedItems(1)
	End With

	'Nastavení jména uživatele
	UserName = Environ("username")

    'Vytvoření jména souboru z jeho celého jména (ořezání stringů)
    test = InStr(sourceFilePath, "(")
    test2 = Mid(sourceFilePath, test + 1, 20)
    test3 = InStr(test2, " ")
    setMonth = Left(test2, test3 - 1)
    test4 = Left(test2, test3 + 4)
    setYear = Right(test4, 4)

    Worksheets("template").Visible = False
 
    'template visibility checker
	If Worksheets("template").Visible = False Then
		Worksheets("template").Visible = True
		Else
    End If
    
    'Vytvoření cesty k šablonám použitým pro tvorbu grafů
    UserName = Environ("username")
    templatePath = "C:\Users\" & UserName & "\AppData\Roaming\Microsoft\Šablony\Charts\"
    Workbooks.Open (sourceFilePath)
    Workbooks(thisWbName).Activate

    'Výběr měsíce (pro vytvoření jména listu)
    setMonth = LCase(setMonth)
    Select Case setMonth
        Case "leden"
            mPointer = "01"
        Case "únor"
            mPointer = "02"
        Case "březen"
            mPointer = "03"
        Case "duben"
            mPointer = "04"
        Case "květen"
            mPointer = "05"
        Case "červen"
            mPointer = "06"
        Case "červenec"
            mPointer = "07"
        Case "srpen"
            mPointer = "08"
        Case "září"
            mPointer = "09"
        Case "říjen"
            mPointer = "10"
        Case "listopad"
            mPointer = "11"
        Case "prosinec"
            mPointer = "12"
        Case Else
            mPointer = "ERR_"
    End Select
	
    'ukazatel prvního řádku produktu F (95)
    baseNumF = 95
    'ukazatel prvního řádku produktu G (122)
    baseNumG = 122
    'Vytvoření názvu listu
    newSheetName = "FG_" & mPointer & setYear & "_reg"
    'Kontrola, zda list s tímto jménem existuje, pokud ano, bude smazán (bez upozornění)
    For Each sh In Worksheets
        If sh.Name Like newSheetName Then flag = True: Exit For
        Next
        If flag = True Then
        Sheets(newSheetName).Delete
    End If
	
    'Vytvoření nového listu
    Workbooks(thisWbName).Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = newSheetName
    sPointer = Workbooks(thisWbName).Worksheets.Count
    'Vytvoření souhrného data ve formátu MM_YYYY pro pisování grafů
    setMonth = setMonth & " " & setYear
    'worksheetNum se chová jako ukazatel listu regionu v Plánu
    worksheetNum = 1
    'Ukazatel sloupce
    cPointer = 2
    'Nastavéní názvu souboru, ze kterého se budou brát data
    DataSourceWb = "Plán a skutečnost v regionech produkce a recidiva (" & setMonth & ").xls"
    'DoEvents pro zajištění synchronizace MS EXCEL a WIN
    DoEvents
    'krokuje všechny listy (1-8) v "Plán a skutečnost v regionech produkce a recidiva"
	
    For i = 1 To 8
        Workbooks(DataSourceWb).Activate
        Worksheets(worksheetNum).Activate
        'get product F data
        Range("D" & baseNumF).Activate
        rowNum = 0
        'hledá poslední řádek obsahující data
        Do Until IsEmpty(ActiveCell)
            If ActiveCell = 0 Then
                rowNum = rowNum + 1
                ActiveCell.Offset(1, 0).Select
            ElseIf ActiveCell > 0 Then
                rowNum = rowNum + 1
                ActiveCell.Offset(1, 0).Select
            Else
                rowNum = rowNum + 1
                ActiveCell.Offset(1, 0).Select
            End If
        Loop
		
        'Naplnění proměnných daty z příslušných sloupců
        productPointer = "F"
        datum = Range("B" & baseNumF + rowNum - 1).Value
        planData = Range("C" & baseNumF + rowNum - 1).Value
        realData = Range("D" & baseNumF + rowNum - 1).Value
        requestPlan = Range("AH" & baseNumF + rowNum - 1).Value
        requestReal = Range("AI" & baseNumF + rowNum - 1).Value
        
        'Naplnění pole
        productsArray(1) = planData
        productsArray(2) = realData
        productsArray(3) = requestPlan
        productsArray(4) = requestReal
        Workbooks(thisWbName).Activate
        Worksheets(sPointer).Activate
        fPointer = 6
        
        'vytisknutí dat z pole
        For j = 1 To 4
            Cells(fPointer, cPointer).Value = productsArray(j)
            fPointer = fPointer + 1
        Next
        
        'Produkt G
        Workbooks(DataSourceWb).Activate
        Worksheets(worksheetNum).Activate
        
        Erase productsArray
        productPointer = "G"
        Range("D" & baseNumG).Activate
        rowNum = 0
        
        Do Until IsEmpty(ActiveCell)
            rowNum = rowNum + 1
            ActiveCell.Offset(1, 0).Select
        Loop
        planData = Range("C" & baseNumG + rowNum - 1).Value
        realData = Range("D" & baseNumG + rowNum - 1).Value
        requestPlan = Range("AH" & (baseNumG + rowNum - 1)).Value
        requestReal = Range("AI" & baseNumG + rowNum - 1).Value
        
        productsArray(1) = planData
        productsArray(2) = realData
        productsArray(3) = requestPlan
        productsArray(4) = requestReal
        Workbooks(thisWbName).Activate
        Worksheets(sPointer).Activate
        Range("B16").Activate
        gPointer = 16
        
        For j = 1 To 4
            Cells(gPointer, cPointer).Value = productsArray(j)
            gPointer = gPointer + 1
        Next

        Workbooks(DataSourceWb).Activate
        Worksheets(worksheetNum).Activate
        worksheetNum = worksheetNum + 1
        cPointer = cPointer + 1
        Erase productsArray
    Next
	
    'zavře plán a skutečnost v regionech bez uložení změn (dělá se proto aby nedošlo ke kumulování souborů a zpomalování skriptu)
    Workbooks(DataSourceWb).Close Savechanges:=False

    
    'Tady začíná prasárna s macro recorderem
    'názvy je potřeba přepsat do funkce (array smyčky)
    'pro vytváření grafů se používají grafy z listu Template (xlVeryHiden)
    'a kopírují se na nový list, potom měním zdroje dat a další parametry (pozice a typ)
    'nejspíš by to šlo i bez AppData a kopírovat grafy i s jejich vzhledem, ale nechce se mi to předělávat
    'forátování tabulek je opět autoMakro, takže někdy je to potřeba předělat
    
    
    'MsgBox "Vytvoření datasetu - OK"
    DoEvents
    Workbooks(thisWbName).Activate
    Worksheets(sPointer).Activate
    Cells(1, 1).Value = setMonth
    Cells(5, 1).Value = "Produkt F"
    
    For i = 1 To 7
		Cells(5, i + 1).Value = "REGION " & i
    Next

    Cells(5, 9).Value = "CELKEM"
    Cells(6, 1).Value = "Plán"
    Cells(7, 1).Value = "Skutečnost"
    Cells(8, 1).Value = "Žádosti - plán"
    Cells(9, 1).Value = "Žádosti - skutečnost"
    Cells(10, 1).Value = "Plnění plánu"
    Cells(11, 1).Value = "100%"
    Cells(12, 1).Value = "Rozdíl"
    Range("B11:I11").Value = 1

    Cells(15, 1).Value = "Produkt G"
    For i = 1 To 7
		Cells(15, i + 1).Value = "REGION " & i
    Next

    Cells(15, 9).Value = "CELKEM"
    Cells(16, 1).Value = "Plán"
    Cells(17, 1).Value = "Skutečnost"
    Cells(18, 1).Value = "Žádosti - plán"
    Cells(19, 1).Value = "Žádosti - skutečnost"
    Cells(20, 1).Value = "Plnění plánu"
    Cells(21, 1).Value = "100%"
    Cells(22, 1).Value = "Rozdíl"
    Range("B21:I21").Value = 1

    Range("B10").Select
    ActiveCell.FormulaR1C1 = "=R[-3]C/R[-4]C"
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "=R[-5]C-R[-6]C"
    Range("B10").Select
    Selection.AutoFill Destination:=Range("B10:I10"), Type:=xlFillDefault
    Range("B10:I10").Select
    Range("B12").Select
    Selection.AutoFill Destination:=Range("B12:I12"), Type:=xlFillDefault
    Range("B12:I12").Select
    Range("B20").Select
    ActiveCell.FormulaR1C1 = "=R[-3]C/R[-4]C"
    Range("B22").Select
    ActiveCell.FormulaR1C1 = "=R[-5]C-R[-6]C"
    Range("B20").Select
    Selection.AutoFill Destination:=Range("B20:I20"), Type:=xlFillDefault
    Range("B20:I20").Select
    Range("B22").Select
    Selection.AutoFill Destination:=Range("B22:I22"), Type:=xlFillDefault
    Range("B22:I22").Select
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
    Range("I6").Select
    Selection.AutoFill Destination:=Range("I6:I9"), Type:=xlFillDefault
    Range("I6:I9").Select
    Range("I16").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
    Range("I16").Select
    Selection.AutoFill Destination:=Range("I16:I19"), Type:=xlFillDefault
    Range("I16:I19").Select
    DoEvents
    Range("B6:I7").Select
    Selection.NumberFormat = "#,##0 Kč"
    Range("B8:I9").Select
    Selection.NumberFormat = "0"
    Range("B10:I10").Select
    Selection.Style = "Percent"
    Range("B11:I11").Select
    Selection.Style = "Percent"
    Range("B12:I12").Select
    Selection.NumberFormat = "#,##0 Kč"
    Range("B16:I17").Select
    Selection.NumberFormat = "#,##0 Kč"
    Range("B18:I19").Select
    Selection.NumberFormat = "0"
    Range("B20:I20").Select
    Selection.Style = "Percent"
    Range("B21:I21").Select
    Selection.Style = "Percent"
    Range("B22:I22").Select
    Selection.NumberFormat = "#,##0 Kč"

    Range("A5:I5").Select
    Selection.Font.Bold = True
    Range("A6:A12").Select
    Selection.Font.Bold = True
    Range("A15:I15").Select
    Selection.Font.Bold = True
    Range("A16:A22").Select
    Selection.Font.Bold = True
    Range("A5:A12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("A15:A22").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("B15:I15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("B5:I5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("B6:I12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Range("B16:I22").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Columns("H:H").Select
    Range("H4").Activate
    Selection.EntireColumn.Hidden = True
    Range("A5").Select
    Worksheets("template").Visible = True
    setMonth = StrConv(setMonth, vbProperCase)
    
    DoEvents
    Sheets("template").Select
    ActiveSheet.ChartObjects("Graf 2").Activate
    ActiveChart.ChartArea.Copy
    Sheets(newSheetName).Select
    Range("K5").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Graf 1").Activate
    ActiveChart.SetSourceData Source:=Range("A5:G7")
    ActiveChart.SetSourceData Source:=Range("A5:G7,A12:G12")
    ActiveChart.SeriesCollection(1).Name = "=""Plán"""
    ActiveChart.SeriesCollection(2).Name = "=""Skutečnost"""
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Produkt F plán a skutečnost (" & LCase(setMonth) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Produkt F plán a skutečnost (" & LCase(setMonth) & ")"
    ActiveChart.ChartType = xlBarClustered
    ActiveChart.ApplyChartTemplate ( _
        templatePath & "FG_1.crtx")
    ActiveSheet.Shapes("Graf 1").ScaleWidth 1.0732009926, msoFalse, _
        msoScaleFromTopLeft
       
    Sheets("template").Select
    Range("U2").Select
    ActiveSheet.ChartObjects("Graf 3").Activate
    ActiveChart.ChartArea.Copy
    Sheets(newSheetName).Select
    Range("Y5").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Graf 2").Activate
    ActiveChart.SetSourceData Source:=Range("B5:I5")
    ActiveChart.SetSourceData Source:=Range("B5:I5,B10:I11")
    ActiveChart.SeriesCollection(1).Name = "=""Plnění plánu NH"""
    ActiveChart.SeriesCollection(2).Name = "=""100%"""
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Produkt F plnění (" & LCase(setMonth) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
	"Produkt F plnění (" & LCase(setMonth) & ")"
    ActiveChart.ChartType = xlBarClustered
    ActiveChart.ApplyChartTemplate ( _
	templatePath & "FG_2.crtx")
    Sheets("template").Select


	Sheets("template").Select
    ActiveSheet.ChartObjects("Graf 4").Activate
    ActiveChart.ChartArea.Copy
    Sheets(newSheetName).Select
    Range("AK5").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Graf 3").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Produkt F počty žádostí a plnění plánu (" & LCase(setMonth) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
	"Produkt F počty žádostí a plnění plánu (" & LCase(setMonth) & ")"
    ActiveChart.SeriesCollection(1).Name = "=" & newSheetName & "!$A$10"
    ActiveChart.SeriesCollection(1).Values = "=" & newSheetName & "!$B$10:$G$10"
    ActiveChart.SeriesCollection(2).Name = "=" & newSheetName & "!$A$8"
    ActiveChart.SeriesCollection(2).Values = "=" & newSheetName & "!$B$8:$G$8"
    ActiveChart.SeriesCollection(3).Name = "=" & newSheetName & "!$A$9"
    ActiveChart.SeriesCollection(3).Values = "=" & newSheetName & "!$B$9:$G$9"

    DoEvents
    'GRAF Produkt G Plán a skutečnost
    Sheets("template").Select
    ActiveSheet.ChartObjects("Graf 5").Activate
    ActiveChart.ChartArea.Copy
    Sheets(newSheetName).Select
    Range("K30").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Graf 4").Activate
    ActiveChart.SetSourceData Source:=Range("A15:G17")
    ActiveChart.SetSourceData Source:=Range("A15:G17,A22:G22")
    ActiveChart.SeriesCollection(1).Name = "=""Plán"""
    ActiveChart.SeriesCollection(2).Name = "=""Skutečnost"""
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Produkt G plán a skutečnost (" & LCase(setMonth) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
	"Produkt G plán a skutečnost (" & LCase(setMonth) & ")"
    ActiveChart.ChartType = xlBarClustered
    ActiveChart.ApplyChartTemplate ( _
	templatePath & "FG_1.crtx")
    ActiveSheet.Shapes("Graf 4").ScaleWidth 1.0732009926, msoFalse, _
	msoScaleFromTopLeft

    Sheets("template").Select
    ActiveSheet.ChartObjects("Graf 7").Activate
    ActiveChart.ChartArea.Copy
    Sheets(newSheetName).Select
    Range("Y30").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Graf 5").Activate
    ActiveChart.SetSourceData Source:=Range("B15:I15")
    ActiveChart.SetSourceData Source:=Range("B15:I15,B20:I21")
    ActiveChart.SeriesCollection(1).Name = "=""Plnění plánu NH"""
    ActiveChart.SeriesCollection(2).Name = "=""100%"""
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Produkt G plnění (" & LCase(setMonth) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
	"Produkt G plnění (" & LCase(setMonth) & ")"
    ActiveChart.ChartType = xlBarClustered
    ActiveChart.ApplyChartTemplate ( _
	templatePath & "FG_2.crtx")

    Sheets("template").Select
    ActiveSheet.ChartObjects("Graf 10").Activate
    ActiveChart.ChartArea.Copy
    Sheets(newSheetName).Select
    Range("AK30").Select
    ActiveSheet.Paste
    Range("AK28").Select
    ActiveSheet.ChartObjects("Graf 6").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Produkt G počty žádostí a plnění plánu (" & LCase(setMonth) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
	"Produkt G počty žádostí a plnění plánu (" & LCase(setMonth) & ")"
    ActiveChart.SeriesCollection(1).Name = "=" & newSheetName & "!$A$20"
    ActiveChart.SeriesCollection(1).Values = "=" & newSheetName & "!$B$20:$G$20"
    ActiveChart.SeriesCollection(2).Name = "=" & newSheetName & "!$A$18"
    ActiveChart.SeriesCollection(2).Values = "=" & newSheetName & "!$B$18:$G$18"
    ActiveChart.SeriesCollection(3).Name = "=" & newSheetName & "!$A$19"
    ActiveChart.SeriesCollection(3).Values = "=" & newSheetName & "!$B$19:$G$19"

    DoEvents
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
	With Selection
		.HorizontalAlignment = xlRight
		.VerticalAlignment = xlBottom
		.WrapText = False
		.Orientation = 0
		.AddIndent = False
		.IndentLevel = 0
		.ShrinkToFit = False
		.ReadingOrder = xlContext
		.MergeCells = False
	End With
    Range("A1").Select
	With Selection
		.HorizontalAlignment = xlCenter
		.VerticalAlignment = xlBottom
		.WrapText = False
		.Orientation = 0
		.AddIndent = False
		.IndentLevel = 0
		.ShrinkToFit = False
		.ReadingOrder = xlContext
		.MergeCells = False
	End With
    Selection.Font.Size = 14
    Selection.Font.Bold = True
	Sheets(newSheetName).Activate
	ActiveWindow.DisplayGridlines = False

	Sheets(newSheetName).Select
	With ActiveWorkbook.Sheets(newSheetName).Tab
		.Color = 255
		.TintAndShade = 0
	End With
    DoEvents
 
    Range("A1").Value = UCase(setMonth)
    Range("A2").Value = UCase(datum)

    Call SortWorksheets
    Worksheets("template").Visible = False
    Worksheets("template").Visible = xlSheetVeryHidden

    With Application
		.ScreenUpdating = True
		.EnableEvents = True
		.PrintCommunication = True
		.DisplayAlerts = True
    End With

    'aktivuje nově vytvořený list
    Workbooks(thisWbName).Activate
    Worksheets(newSheetName).Activate
	
Exit Sub