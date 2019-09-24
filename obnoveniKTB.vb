Option Explicit

Sub A_VYPLNIT_PRŮBĚŽNÉ_PLNĚNÍ()

        Dim sourceFileDialog As Object
        Dim fileName() As Variant
        Dim filePath() As Variant
        Dim upperArraySize, lastRow, i, y, ktbIndex As Integer
        Dim zadostName, zadostFile, zadostiMsg, STRthisWorkbookPath, STRnstat5DataFile  As String
        Dim sourceFilePath, thisWbName, msgVal As String
        Dim stat03, stat05, stat06, stat13, stat12 As String
        Dim sourceRange As Range
        Dim vrtSelectedItem As Variant
        Dim worksheets As Worksheet
        Dim ktbArrayDataNstat3, ktbArrayDataNstat5 As Variant
        'Nastavení objektu pro výběr souborů
        Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

        'Pole s čísly kontingenčních tabulek na listech STAT3 a Nová produkce
        ktbArrayDataNstat3 = Array(8, 9, 2, 5, 6, 7, 13, 12, 11, 3, 10)
        ktbArrayDataNstat5 = Array(1, 4, 5, 6)

        'Cesta k aktuálnímu sešitu a název souboru
        STRthisWorkbookPath = Application.ActiveWorkbook.Path
        STRnstat5DataFile = Application.ActiveWorkbook.FullName
        STRnstat5DataFile = Right(STRnstat5DataFile, Len(STRnstat5DataFile) - InStrRev(STRnstat5DataFile, "\"))
        'zobrazení formuláře pro odemknutí listů
        UserForm1.Show

        With Application
                .ScreenUpdating = False
                .Calculation = xlCalculationAutomatic
                .DisplayAlerts = False
        End With

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
                'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou, ukončit průběh skriptu
                If .Show <> -1 Then
                        msgVal = MsgBox("Storno - Načítání dat a skript přerušen. Žádná data nebyla do souboru načtena.", vbCritical)
                        Exit Sub
                End If
                'Přiřadí soubor (jeho cestu), který byl vybrán pomocí dialogu do proměné sourceFilePath
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
                        'vymazání stávajících (starých) dat (STAT5) na listu nové produkce
                        lastRow = sheets("Nová produkce").Range("G2").End(xlDown).row
                        sheets("Nová produkce").Range("A2:AE" & lastRow).ClearContents
                        
                        'Vykopírování dat ze souboru stat5on99.xls do listu nové produkce
                        Workbooks.Open (filePath(y))
                        lastRow = Cells(6, "G").End(xlDown).row
                        Set sourceRange = Range("A6:AE" & lastRow)
                        sourceRange.Copy
                        Workbooks(thisWbName).worksheets("Nová produkce").Range("A2").PasteSpecial
                        Workbooks(fileName(y)).Close savechanges:=False
                        
                        'Úprava cesty kontingenční tabulek na listu nové produkce a aktualizace pivotCashe pro ostatní KTB
                        sheets("Nová produkce").PivotTables("Kontingenční tabulka 3").ChangePivotCache ActiveWorkbook.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=STRthisWorkbookPath & "\[" & STRnstat5DataFile & "]Nová produkce!C1:C34", _
                        Version:=xlPivotTableVersion15)
                
                        For ktbIndex = 0 To UBound(ktbArrayDataNstat5)
                                sheets("Nová produkce").PivotTables("Kontingenční tabulka " & ktbArrayDataNstat5(ktbIndex)).ChangePivotCache ("Kontingenční tabulka 3")
                        Next ktbIndex
                        Exit For
                End If
        Next y

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''Kopírování STAT012 (A7-AV(xlEnd))'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        For y = 0 To UBound(filePath())
                If fileName(y) = "nstat12.xls" Then
                        'vymazání stávajících (starých) dat (STAT12) na listu nstat12
                        sheets("nstat12").Activate
                        lastRow = sheets("nstat12").Range("G3").End(xlDown).row
                        sheets("nstat12").Range("A3:AW" & lastRow).Select
                        Selection.ClearContents
                        
                        'Vykopírování dat ze souboru stat5on99.xls do listu nové produkce
                        Workbooks.Open (filePath(y))
                        lastRow = Cells(6, "F").End(xlDown).row
                        Set sourceRange = Range("A6:Ax" & lastRow)
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
                        'vymazání stávajících (starých) dat (STAT06) na listu revolvingy
                        lastRow = sheets("Revolvingy").Range("E2").End(xlDown).row
                        sheets("Revolvingy").Range("A2:W" & lastRow).ClearContents
                        'Vykopírování dat ze souboru stat6on99.xls do listu revolvingy
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
                        'vymazání stávajících (starých) dat (STAT13) na listu NSTAT13
                        sheets("NSTAT13").Activate
                        lastRow = sheets("NSTAT13").Range("A7").End(xlDown).row
                        sheets("NSTAT13").Range("A7:W" & lastRow).Select
                        Selection.ClearContents
                        
                        'Vykopírování dat ze souboru stat13on99.xls do listu NSTAT13
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

'''''''''''''Nacteni STAT03ON99 (A6-AL(xlEnd)) pro obnoveni zadostiPS (pokud dialog = 1)''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If zadostiMsg = vbYes Then
                'presunout for loop sem?! dávalo by to smysl takhle tam je snad zbytečně cyklení pro vyhledávání souboru.
                For y = 0 To UBound(filePath)
                        If fileName(y) = "stat3on99.xls" Then
                                'Vymazání starých dat ze souboru žádostPS
                                Workbooks.Open (zadostFile)
                                Workbooks.Open (filePath(y))
                                Workbooks(zadostName).Activate
                                lastRow = Cells(5, "AL").End(xlDown).row
                                Range("A5:AL" & lastRow).ClearContents
                                
                                'Výběr datumu pro žádostPS
                                datePicker.Show
                                Workbooks(fileName(y)).Activate
                                
                                'Kopírování dat ze souboru stat03on99 do souboru žádostPS
                                lastRow = Cells(6, "F").End(xlDown).row
                                Set sourceRange = Range("A6:AL" & lastRow)
                                sourceRange.Copy
                                Workbooks(zadostName).worksheets("Nstat3").Range("A5").PasteSpecial
                                Workbooks(zadostName).Save
                                Workbooks(fileName(y)).Close savechanges:=False
                                Workbooks(thisWbName).Activate
                                
                                'Obnovení cesty ke kontingenčním tabulkám a aktualizace pivot cashe pro ostatní KTB
                                sheets("Nstat3").PivotTables("Kontingenční tabulka 4").ChangePivotCache _
                                ActiveWorkbook.PivotCaches.Create( _
                                SourceType:=xlDatabase, _
                                SourceData:=STRthisWorkbookPath & "\[žádost PS.xlsm]Nstat3!R4C1:R1048576C57", _
                                Version:=xlPivotTableVersion15)
                                
                                For ktbIndex = 0 To UBound(ktbArrayDataNstat3)
                                        sheets("Nstat3").PivotTables("Kontingenční tabulka " & ktbArrayDataNstat3(ktbIndex)).ChangePivotCache ("Kontingenční tabulka 4")
                                Next ktbIndex
                                
                                ThisWorkbook.RefreshAll
                                Workbooks(zadostName).Close savechanges:=False
                                Exit For
                        End If
                Next y
        Else
                'Pokud uživatel nechce obnovovat data v žádostPS dojde pouze k jejich načtení a aktualizaci cest KTB
                Workbooks.Open (zadostName)
                Workbooks(thisWbName).Activate
                ThisWorkbook.RefreshAll
                Workbooks(zadostName).Close savechanges:=True
        End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''Nulování proměnných a volání původních skriptů pro vyplnění reportu'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Set sourceFileDialog = Nothing
        ReDim fileName(0)
        ReDim filePath(0)

        Call selectWorksheetsToFill
        Call denni_vyplneni

        With Application
                .ScreenUpdating = True
                .DisplayAlerts = True
                .Calculation = xlCalculationAutomatic
        End With
End Sub
