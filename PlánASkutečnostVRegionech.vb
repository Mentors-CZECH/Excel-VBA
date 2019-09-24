'Option Explicit

Sub VYPLNIT_PLÁN_A_SKUTEČNOST()

    Dim sourceFileDialog As Object
    Dim fileName() As Variant
    Dim filePath() As Variant
    Dim upperArraySize, lastrow, i, y, t, x As Integer
    Dim zadostName, zadostFile, zadostiMsg As String
    Dim sourceFilePath, thisWbName, msgVal As String
    Dim stat03, stat05, stat06, stat13, stat12 As String
    Dim sourceRange As Range
    Dim vrtSelectedItem As Variant
    Dim rowsArray() As Variant
    Dim productArray() As Variant
    Dim stat9array() As Variant

    'Dim worksheets As Worksheet

    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    productArray() = Array("Produkty AEFG", "Produkt A", "Produkt E", "Produkt F", "Produkt G")
    stat9array() = Array("ABCDEFG", "A", "E", "F", "G")

    ReDim rowsArray(0 To 4)

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    Call odkrytiMKTzdroje
    'Jméno sešitu, který volá dialog
    thisWbName = ActiveWorkbook.Name
    zadostiMsg = MsgBox("Chcete aktualizovat žádosti PS?", vbYesNo + vbQuestion, "Žádosti PS.xlsm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''Zobrazení dialogu pro výběr statistik'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With sourceFileDialog
        If zadostiMsg = vbYes Then
            .Title = "Zadejte cestu k souboru: STAT03, STAT05, STAT13, STAT09"
        Else
            .Title = "Zadejte cestu k souboru: STAT05, STAT013, STAT09"
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''stat05 A6-AE(xlEnd)''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "stat5on99.xls" Then
            Workbooks.Open (filePath(y))
            lastrow = Cells(6, "F").End(xlDown).Row
            Set sourceRange = Range("A6:AE" & lastrow)
            sourceRange.Copy
            Workbooks(thisWbName).Worksheets("Stat5").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            Exit For
        End If
    Next y

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''stat13 A6-AV(xlEnd)'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "nstat13p.xls" Then
            Workbooks.Open (filePath(y))
            lastrow = Cells(7, "L").End(xlDown).Row
            Set sourceRange = Range("A7:L" & lastrow)
            sourceRange.Copy
            Workbooks(thisWbName).Worksheets("nábor").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            Exit For
        End If
    Next y

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''stat03 A6-AL(xlEnd)''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If zadostiMsg = vbYes Then
            For y = 0 To UBound(filePath())
                If fileName(y) = "stat3on99.xls" Then
                    Workbooks.Open (filePath(y))

                    'tohle je rozsah žádosti PS, který chci smazat
                    Workbooks(zadostName).Activate
                    lastrow = Cells(5, "AL").End(xlDown).Row
                    'tady mažu obsah uplně někde jinde
                    Range("A5:AL" & lastrow).ClearContents
                    Workbooks(fileName(y)).Activate
                    lastrow = Cells(6, "F").End(xlDown).Row
                    Set sourceRange = Range("A6:AL" & lastrow)
                    sourceRange.Copy
                    
                    Workbooks(zadostName).Worksheets("Nstat3").Range("A5").PasteSpecial
                    Workbooks(zadostName).Save
                    Set sourceRange = Workbooks(zadostName).Worksheets("Nstat3").Range("BF3:CN67")
                    sourceRange.Copy
                    Workbooks(thisWbName).Worksheets("Nstat3").Range("AU3").PasteSpecial xlPasteValues
                    Workbooks(fileName(y)).Close saveChanges:=False
'                    Workbooks(thisWbName).Activate
'                    ThisWorkbook.RefreshAll
                    Workbooks(zadostName).Close saveChanges:=False
                    Exit For
                End If
            Next y

        Else
            Workbooks.Open (zadostName)
            Set sourceRange = Range("BF3:CN67")
            sourceRange.Copy
            Workbooks(thisWbName).Worksheets("Nstat3").Range("AU3").PasteSpecial xlPasteValues
            Workbooks(zadostName).Close saveChanges:=True
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''Obnovit KTB'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'přepsat do smyčky - přejmenovat tabulky 1-5 na ktbST5_1 a vytvořit for loop
    Workbooks(thisWbName).Worksheets("nábor").Activate
    ActiveSheet.PivotTables("Kontingenční tabulka 1").PivotCache.Refresh
    Worksheets("Stat5").Activate
    ActiveSheet.PivotTables("Kontingenční tabulka 2").PivotCache.Refresh
    ActiveSheet.PivotTables("Kontingenční tabulka 7").PivotCache.Refresh
    ActiveSheet.PivotTables("Kontingenční tabulka 6").PivotCache.Refresh
    ActiveSheet.PivotTables("Kontingenční tabulka 5").PivotCache.Refresh
    ActiveSheet.PivotTables("Kontingenční tabulka 1").PivotCache.Refresh

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''Najdi den'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Worksheets(1).Activate
    lastrow = Cells(4, "D").End(xlDown).Row
    lastrow = lastrow - 1
    Worksheets(1).Activate
    For x = 1 To 8
        Worksheets(x).Select (False)
        Cells(lastrow, 2).Activate
    Next x

    For t = 0 To 4
        rowsArray(t) = ActiveSheet.Range("B:B").EntireColumn. _
        Find(What:=productArray(t), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Row
        rowsArray(t) = rowsArray(t) + lastrow
    Next t

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''Plnění stat9 do listu a volání původního makra'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath)
        If Left(fileName(y), 5) = "STAT9" Then
            'Přesunout for (x), které generuje hodnoty pro produkty A-G sem. potom nebudu muset 2X kopírovat data pro ABCDEFG.
            Workbooks.Open (filePath(y))
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).Worksheets("Nstat9").Range("A1").PasteSpecial
            Worksheets("AC").Activate
            Set sourceRange = Range("G17:U41")
            sourceRange.Copy
            Workbooks(thisWbName).Worksheets("Nstat9").Range("G17").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            Worksheets(12).Activate
            Call Obdélník1_Kliknutí

            Worksheets(1).Activate
            For x = 1 To 8
                Worksheets(x).Select (False)
                Cells(rowsArray(0), 2).Activate
            Next x

            Call prepisNstat9

            For x = 1 To 4
                Workbooks.Open (filePath(y))
                Worksheets(stat9array(x)).Activate

                Set sourceRange = Range("A1:AA79")
                sourceRange.Copy
                Workbooks(thisWbName).Worksheets("Nstat9").Range("A1").PasteSpecial
                Workbooks(fileName(y)).Close saveChanges:=False

                'Produkty A-G
                If x = 1 Then
                Worksheets(12).Activate
                        Call Obdélník2_Kliknutí
                        Worksheets(1).Activate
                        For Z = 1 To 8
                            Worksheets(Z).Select (False)
                            Cells(rowsArray(x), 2).Activate
                        Next Z
                        Call prepisNstat9
                    ElseIf x = 2 Then
                    Worksheets(12).Activate
                        Call Obdélník3_Kliknutí
                        Worksheets(1).Activate
                        For Z = 1 To 8
                            Worksheets(Z).Select (False)
                            Cells(rowsArray(x), 2).Activate
                        Next Z
                        Call prepisNstat9
                    ElseIf x = 3 Then
                    Worksheets(12).Activate
                        Call Obdélník4_Kliknutí
                        Worksheets(1).Activate
                        For Z = 1 To 8
                            Worksheets(Z).Select (False)
                            Cells(rowsArray(x), 2).Activate
                        Next Z
                        Call prepisNstat9
                    ElseIf x = 4 Then
                    Worksheets(12).Activate
                        Call Obdélník5_Kliknutí
                        Worksheets(1).Activate
                        For Z = 1 To 8
                            Worksheets(Z).Select (False)
                            Cells(rowsArray(x), 2).Activate
                        Next Z
                        Call prepisNstat9
                End If
                    
            Next x

            Workbooks.Open (filePath(y))
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).Worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            Call skrytiMKTzdroje

        End If
    Next y

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set sourceFileDialog = Nothing
    ReDim fileName(0)
    ReDim filePath(0)
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub






