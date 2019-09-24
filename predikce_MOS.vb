'Option Explicit

Sub Makro2()

'zadej cestu k žádosti ps
'zadej cestu k predikce MOS
'vyber zdrojová data
'for each ktb in workbook
'přemapovat ktb revolvingy
'přemapovat ktb nové produkce


    Dim sourceFileDialog As Object
    Dim fileName() As Variant
    Dim filePath() As Variant
    Dim upperArraySize, lastRow, i, y As Integer
    Dim zadostName, zadostFile, zadostiMsg As String
    Dim sourceFilePath, thisWbName, msgVal As String
    Dim stat03, stat05, stat06, stat13, stat12 As String
    Dim sourceRange As Range
    Dim vrtSelectedItem As Variant
    Dim worksheets As Worksheet
    
    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)


    With Application
        .ScreenUpdating = False
        .Calculation = xlManual
        .DisplayAlerts = False
    End With

    UserForm1.Show
    'Jméno sešitu, který volá dialog
    thisWbName = ActiveWorkbook.Name
    zadostiMsg = MsgBox("Chcete aktualizovat žádosti PS?", vbYesNo + vbQuestion, "Žádosti PS.xlsm")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''Dialogové okno pro výběr souborů statistik'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With sourceFileDialog
        If zadostiMsg = vbYes Then
            .Title = "Zadejte cestu k souboru: STAT03, STAT05, STAT06, STAT12, STAT13"
        Else
            .Title = "Zadejte cestu k souboru: STAT05, STAT06, STAT12, STAT13"
        End If
        .Filters.Clear
        .Filters.Add "Soubory MS Excel", "*.xl*"
        .ButtonName = "Načíst data"
        .AllowMultiSelect = True
        'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou
        If .Show <> -1 Then
            msgVal = MsgBox("Storno - Načítání dat a skript přerušen. Žádná data nebyla do souboru načtena.", vbCritical)
            Exit Sub
        End If
        'Přiřadí soubor (jeho cestu), který jste vybraly pomocí dialogu do proměné sourceFilePath
            i = 0
            For Each vrtSelectedItem In .SelectedItems
                ReDim Preserve filePath(0 To i)
                ReDim Preserve fileName(0 To i)
                filePath(i) = vrtSelectedItem
                fileName(i) = Dir(filePath(i))
                i = i + 1
            Next
    End With
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Kopírování STAT05ON99 (A6-AE(xlEnd))''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To 6
        If fileName(y) = "stat5on99.xls" Then
            Workbooks.Open (filePath(y))
            lastRow = Cells(6, "F").End(xlDown).row
            Set sourceRange = Range("A6:AE" & lastRow)
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nová produkce").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close savechanges:=False
            Exit For
        End If
    Next y
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Kopírování STAT012 (A7-AV(xlEnd))'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "nstat12.xls" Then
            Workbooks.Open (filePath(y))
            lastRow = Cells(6, "F").End(xlDown).row
            Set sourceRange = Range("A6:AV" & lastRow)
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("nstat12").Range("A3").PasteSpecial
            Workbooks(fileName(y)).Close savechanges:=False
            Exit For
            
        End If
    Next y
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Kopírování STAT06ON99 (A6-W(xlEnd))'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "stat6on.xls" Then
            Workbooks.Open (filePath(y))
            lastRow = Cells(6, "F").End(xlDown).row
            Set sourceRange = Range("A6:W" & lastRow)
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Revolvingy").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close savechanges:=False
            Exit For
        End If
    Next y



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Kopírování STAT13ON99 (A-W(xlEnd))'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "nstat13p.xls" Then
            Workbooks.Open (filePath(y))
            lastRow = Cells(7, "W").End(xlDown).row
            Set sourceRange = Range("A7:W" & lastRow)
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("NSTAT13").Range("A7").PasteSpecial
            Workbooks(fileName(y)).Close savechanges:=False
            Exit For
        End If
    Next y



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Zadání cesty k souboru ŽádostiPS.xlsm''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
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

        If zadostiMsg = vbYes Then
            'presunout for loop sem?! dávalo by to smysl takhle tam je snad zbytečně cyklení pro vyhledávání souboru.
            For y = 0 To UBound(filePath)
                If fileName(y) = "stat3on99.xls" Then

                    Workbooks.Open (zadostFile)
                    Set zadostiData = ThisWorkbook.worksheets("Nstat3")
                    Workbooks.Open (filePath(y))
                    Workbooks(zadostName).Activate
                    lastRow = Cells(5, "AL").End(xlDown).row
                    'tady mažu obsah uplně někde jinde
                    Range("A5:AL" & lastRow).ClearContents
                    datePicker.Show
                    Workbooks(fileName(y)).Activate
                    lastRow = Cells(6, "F").End(xlDown).row
                    Set sourceRange = Range("A6:AL" & lastRow)
                    sourceRange.Copy
                    Workbooks(zadostName).worksheets("Nstat3").Range("A5").PasteSpecial
                    Workbooks(zadostName).Save
        
                    Workbooks(fileName(y)).Close savechanges:=False
                    Workbooks(thisWbName).Activate
                    ThisWorkbook.RefreshAll
                    Workbooks(zadostName).Close savechanges:=False
                    Exit For
                End If
            Next y
        Else
            Workbooks.Open (zadostName)
            Workbooks(thisWbName).Activate
            ThisWorkbook.RefreshAll
            Workbooks(zadostName).Close savechanges:=True
        End If


    newDataRange = zadostiData.Name & "!R4C1:R1048576C57"
    Workbooks.Open (zadostName)
    For Each pivtTableName In Sheets("Nstat3").PivotTables
        pivtTableName.ChangePivotCache _
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            newDataRange)
    
    Next pivtTableName
    Workbooks(zadostName).Close savechanges:=False
        
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
        
'doplnit cestu k souboru žádosti PS a ponechat oblast
End Sub
