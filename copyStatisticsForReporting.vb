Sub copyStatToWorkbook()

    Dim sourceFileDialog As Object
    Dim fileName() As Variant
    Dim filePath() As Variant
    Dim statType() As Variant
    Dim zadostName As String
    Dim upperArraySize, lastRow As Integer
    Dim USERNAME, sourceFilePath, thisWbName, msgVal As String
    Dim stat03, stat05, stat06, stat13, stat12 As String
    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    Dim pivot As pivotTable
    Dim worksheets As Worksheet
    


    ReDim filePath(0 To 6)
    ReDim fileName(0 To 6)

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    'Nastavení jména uživatele (systémového jména) pro vytváření například ceskty k desktop
    USERNAME = Environ("username")
    
    'Jméno sešitu, který volá dialog
    thisWbName = ActiveWorkbook.Name

    'Je možné přidávat další atributy, a měnit jejich parametry
    With sourceFileDialog
        .Title = "Zadejte cestu k souboru: STAT03,STAT05,STAT06,STAT12,STAT13 a STAT09"
        .Filters.Clear
        .Filters.Add "Soubory MS Excel", "*.xl*"
        .ButtonName = "Načíst data"
        .AllowMultiSelect = True
        'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou
        If .Show <> -1 Then
            msgVal = MsgBox("Storno - chybova hlaska", vbCritical)
            Exit Sub
        End If
        'Přiřadí soubor (jeho cestu), který jste vybraly pomocí dialogu do proměné sourceFilePath
            For Each vrtSelectedItem In .SelectedItems
                filePath(i) = vrtSelectedItem
                fileName(i) = Dir(filePath(i))
                i = i + 1
            Next
    End With

'stat05 A6-AE(xlEnd)
For y = 0 To 6
    If fileName(y) = "stat5on99.xls" Then
        Workbooks.Open (filePath(y))
        lastRow = Cells(6, "F").End(xlDown).row
        SourceRange = Range("A6:AE" & lastRow).Copy
        Workbooks(thisWbName).worksheets("stat05on99").Range("A2").PasteSpecial
        Workbooks(fileName(y)).Close saveChanges:=False
        Exit For
    End If
Next y

'stat12 A6-AV(xlEnd)
For y = 0 To 6
    If fileName(y) = "nstat12.xls" Then
        Workbooks.Open (filePath(y))
        lastRow = Cells(6, "F").End(xlDown).row
        SourceRange = Range("A6:AV" & lastRow).Copy
        Workbooks(thisWbName).worksheets("nstat12").Range("A2").PasteSpecial
        Workbooks(fileName(y)).Close saveChanges:=False
        Exit For
        
    End If
Next y

'stat06 A6-W(xlEnd)
For y = 0 To 6
    If fileName(y) = "stat6on.xls" Then
        Workbooks.Open (filePath(y))
        lastRow = Cells(6, "F").End(xlDown).row
        SourceRange = Range("A6:W" & lastRow).Copy
        Workbooks(thisWbName).worksheets("stat06on").Range("A2").PasteSpecial
        Workbooks(fileName(y)).Close saveChanges:=False
        Exit For
        
    End If
Next y

'stat03 A6-AL(xlEnd)
For y = 0 To 6
    If fileName(y) = "stat3on99.xls" Then
        Workbooks.Open (filePath(y))
        lastRow = Cells(6, "F").End(xlDown).row
        
'rewrite zadosti PS
    With sourceFileDialog
        .Title = "Zadejte cestu k souboru: Žádosti PS"
        .Filters.Clear
        .Filters.Add "Soubory MS Excel", "*.xlsm"
        .ButtonName = "Načíst data"
        .AllowMultiSelect = False
        'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou
        If .Show <> -1 Then
            msgVal = MsgBox("Storno - chybova hlaska", vbCritical)
            Exit Sub
        End If
        'Přiřadí soubor (jeho cestu), který jste vybraly pomocí dialogu do proměné sourceFilePath
            zadostFile = .SelectedItems(1)
    End With
        
        zadostName = Dir(zadostFile)
        Workbooks.Open (zadostFile)
		'Vymyzat stara data
        lastRow = Cells(5, "AL").End(xlDown).row
        Range("A5:AL" & lastRow).ClearContents
        Workbooks(fileName(y)).Activate
        SourceRange = Range("A6:AL" & lastRow).Copy
        Workbooks(zadostName).worksheets("Nstat3").Range("A5").PasteSpecial
        ActiveWorkbook.Save
        

        Workbooks(fileName(y)).Close saveChanges:=False
        
        Workbooks(thisWbName).Activate
        For Each Worksheet In ThisWorkbook.worksheets
        
            For Each pivot In Worksheet.PivotTables
             pivot.RefreshTable
            Next pivot
        
        Next Worksheet
        
        Workbooks(zadostName).Close saveChanges:=True
        
        Exit For
    End If
Next y

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
End Sub

