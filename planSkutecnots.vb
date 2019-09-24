Option Explicit

'_REQUIRES 
'MSCOMCT2.OCX - Pro ovládání calendar control (chyby v Excel 2013 32b/64b)
'
'

'_RELEASE_NOTES
'01.04.2015 V0.9 - První verze.
'
'

Sub A_VYPLNIT_PLÁN_A_SKUTEČNOST()
    'Definice proměnných
    Dim sourceFileDialog As Object
    Dim fileName() As Variant
    Dim filePath() As Variant
    Dim lastrow, i, y, t, x, z As Integer
    Dim zadostName, zadostFile, zadostiMsg As String
    Dim sourceFilePath, thisWbName, msgVal As String
    Dim sourceRange As Range
    Dim vrtSelectedItem As Variant
    Dim rowsArray() As Variant
    Dim productArray() As Variant
    Dim stat9array() As Variant

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Nastavení proměnných
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    'Pole pro nalezení řádků, na kterých se nachází produkt
    productArray() = Array("Produkty AEFG", "Produkt A", "Produkt E", "Produkt F", "Produkt G")
    'Seznam listů v sešitu STAT09, které se budou kopírovat do
    stat9array() = Array("ABCDEFG", "A", "E", "F", "G")
    'nastavení rowsArray pro 5 produktů (AEFG, A ,E ,F, G)
    ReDim rowsArray(0 To 4)
    Call odkrytiMKTzdroje
    DoEvents
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    'Jméno sešitu, který volá dialog
    thisWbName = ActiveWorkbook.Name
    'Dotaz na aktualizaci souboru žádosti PS. Pokud již byl soubor aktualizován stiskněte ne -> rychlejší skript
    zadostiMsg = MsgBox("Chcete aktualizovat žádosti PS?", vbYesNo + vbQuestion, "Žádosti PS.xlsm")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Zobrazení dialogu pro výběr statistik
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
'   stat05 A6-AE(xlEnd) - Nalezení dokumentu stat5 a nakopírování na příslušný list
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "stat5on99.xls" Then
            worksheets("Stat5").Activate
            lastrow = Cells(2, "AE").End(xlDown).Row
            Range("A2:AE" & lastrow).ClearContents
            Workbooks.Open (filePath(y))
            lastrow = Cells(6, "F").End(xlDown).Row
            Set sourceRange = Range("A6:AE" & lastrow)
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Stat5").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            Exit For
        End If
    Next y

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   stat13 A6-AV(xlEnd) - nalezení stat13 a kopírování na příslušný list
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath())
        If fileName(y) = "nstat13p.xls" Then
            worksheets("nábor").Activate
            lastrow = Cells(2, "L").End(xlDown).Row
            Range("A2:L" & lastrow).ClearContents
            Workbooks.Open (filePath(y))
            lastrow = Cells(7, "L").End(xlDown).Row
            Set sourceRange = Range("A7:L" & lastrow)
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("nábor").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            Exit For
        End If
    Next y

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   zobrazeni dialogu pro zadání cesty k souboru žádost PS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
'   Pokud bude aktualizován soubor žádost PS, tak načti stat03 (smaž stará data a nakopíruj nová) a přepiš datum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If zadostiMsg = vbYes Then
        For y = 0 To UBound(filePath())
            If fileName(y) = "stat3on99.xls" Then
                Workbooks.Open (filePath(y))
                Workbooks(zadostName).Activate
                lastrow = Cells(5, "AL").End(xlDown).Row
                Range("A5:AL" & lastrow).ClearContents
                'Výběr data, které bude v buňce A1 souoboru žádost PS
                datePicker.Show
                Workbooks(fileName(y)).Activate
                lastrow = Cells(6, "F").End(xlDown).Row
                Set sourceRange = Range("A6:AL" & lastrow)
                sourceRange.Copy
                Workbooks(zadostName).worksheets("Nstat3").Range("A5").PasteSpecial
                Workbooks(zadostName).Save
                Set sourceRange = Workbooks(zadostName).worksheets("Nstat3").Range("BF3:CN67")
                sourceRange.Copy
                Workbooks(thisWbName).worksheets("Nstat3").Range("AU3").PasteSpecial xlPasteValues
                Workbooks(fileName(y)).Close saveChanges:=False
                Workbooks(zadostName).Close saveChanges:=False
                Exit For
            End If
        Next y
    Else
        Workbooks.Open (zadostName)
        Set sourceRange = Range("BF3:CN67")
        sourceRange.Copy
        Workbooks(thisWbName).worksheets("Nstat3").Range("AU3").PasteSpecial xlPasteValues
        Workbooks(zadostName).Close saveChanges:=True
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Obnovení KTB na listech stat5 a nábor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'přepsat do smyčky - přejmenovat tabulky 1-5 na ktbST5_1 a vytvořit for loop
    Workbooks(thisWbName).worksheets("nábor").Activate
    ActiveSheet.PivotTables("Kontingenční tabulka 1").PivotCache.Refresh
    worksheets("Stat5").Activate
    ActiveSheet.PivotTables("KTB_ST5_0").PivotCache.Refresh
    ActiveSheet.PivotTables("KTB_ST5_1").PivotCache.Refresh
    ActiveSheet.PivotTables("KTB_ST5_2").PivotCache.Refresh
    ActiveSheet.PivotTables("KTB_ST5_3").PivotCache.Refresh
    ActiveSheet.PivotTables("KTB_ST5_4").PivotCache.Refresh

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Vyhledání posledního vyplněného dne na prvním listu (R01 nejsou tam vzorce ale data)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    worksheets(1).Activate
    lastrow = Cells(3, "D").End(xlDown).Row
    lastrow = lastrow - 1
    If lastrow = 42 Then
        lastrow = 4
    End If
    'pokud je tabulka prázdná bude lastRow 42 - v tom případě nastavuji první řádek v tabulce
    If lastrow = 41 Then
    lastrow = 2
    End If

    worksheets(1).Activate
    For x = 1 To 8
        worksheets(x).Select (False)
        Cells(lastrow, 2).Activate
    Next x
    'Hledání řádků jednotlivých produktů a provedení offsetu
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
'   Plnění dat ze STAT09 do listu a následné volání původních maker (obdelník kliknutí dle akutálně plněného produktu) a poté prepisNstat9
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For y = 0 To UBound(filePath)
        If Left(fileName(y), 5) = "STAT9" Then
            'Přesunout for (x), které generuje hodnoty pro produkty A-G sem. potom nebudu muset 2X kopírovat data pro ABCDEFG.
            Workbooks.Open (filePath(y))
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            worksheets("AC").Activate
            Set sourceRange = Range("G17:U41")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("G17").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
            worksheets("Nstat3").Activate
            Call Obdélník1_Kliknutí

            worksheets(1).Activate
            For x = 1 To 8
                worksheets(x).Select (False)
                Cells(rowsArray(0), 2).Activate
            Next x

            Call prepisNstat9

            For x = 1 To 4
                Workbooks.Open (filePath(y))
                worksheets(stat9array(x)).Activate

                Set sourceRange = Range("A1:AA79")
                sourceRange.Copy
                Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
                Workbooks(fileName(y)).Close saveChanges:=False

                'Produkty A-G
                If x = 1 Then
                    worksheets("Nstat3").Activate
                    Call Obdélník2_Kliknutí
                    worksheets(1).Activate
                    For z = 1 To 8
                        worksheets(z).Select (False)
                        Cells(rowsArray(x), 2).Activate
                    Next z
                    Call prepisNstat9
                ElseIf x = 2 Then
                    worksheets("Nstat3").Activate
                    Call Obdélník3_Kliknutí
                    worksheets(1).Activate
                    For z = 1 To 8
                        worksheets(z).Select (False)
                        Cells(rowsArray(x), 2).Activate
                    Next z
                    Call prepisNstat9
                ElseIf x = 3 Then
                    worksheets("Nstat3").Activate
                    Call Obdélník4_Kliknutí
                    worksheets(1).Activate
                    For z = 1 To 8
                        worksheets(z).Select (False)
                        Cells(rowsArray(x), 2).Activate
                    Next z
                    Call prepisNstat9
                ElseIf x = 4 Then
                    worksheets("Nstat3").Activate
                    Call Obdélník5_Kliknutí
                    worksheets(1).Activate
                    For z = 1 To 8
                        worksheets(z).Select (False)
                        Cells(rowsArray(x), 2).Activate
                    Next z
                    Call prepisNstat9
                End If
            Next x

            Workbooks.Open (filePath(y))
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
        End If
    Next y

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Nastavení proměnných na 0 a objekty na nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set sourceFileDialog = Nothing
    ReDim fileName(0)
    ReDim filePath(0)
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    Call odkrytiMKTzdroje

End Sub
