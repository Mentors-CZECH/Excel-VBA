Option Explicit

Sub VYPLNIT_PRŮBĚŽNÉ_PLNĚNÍ()

    Dim sourceFileDialog As Object
    Dim fileName() As Variant
    Dim filePath() As Variant
    Dim lastRow, i, y As Integer
    Dim zadostName, zadostFile, zadostiMsg As String
    Dim sourceFilePath, thisWbName, msgVal As String
    Dim sourceRange As Range
    Dim fileItem As Variant
    Dim worksheets As Worksheet
 
    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    UserForm1.Show

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    thisWbName = ActiveWorkbook.Name
    zadostiMsg = MsgBox("Chcete aktualizovat žádosti PS?", vbYesNo + vbQuestion, "Žádosti PS.xlsm")
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''Dialogové okno pro výběr souborů statistik'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With sourceFileDialog
        If zadostiMsg = vbYes Then
            .Title = "Zadejte cestu k souboru: STAT03, STAT05, STAT06, STAT12"
        Else
            .Title = "Zadejte cestu k souboru: STAT05, STAT06, STAT12"
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
            For Each fileItem In .SelectedItems
                ReDim Preserve filePath(0 To i)
                ReDim Preserve fileName(0 To i)
                filePath(i) = fileItem
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
            Workbooks(fileName(y)).Close saveChanges:=False
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
            Workbooks(thisWbName).worksheets("nstat12").Range("A2").PasteSpecial
            Workbooks(fileName(y)).Close saveChanges:=False
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
            Workbooks(fileName(y)).Close saveChanges:=False
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
                    Workbooks.Open (zadostFile)
                    Workbooks.Open (filePath(y))
                    'tohle je rozsah žádosti PS, který chci smazat
                    Workbooks(zadostName).Activate
                    lastRow = Cells(5, "AL").End(xlDown).row
                    'tady mažu obsah uplně někde jinde
                    Range("A5:AL" & lastRow).ClearContents
                    Workbooks(fileName(y)).Activate
                    lastRow = Cells(6, "F").End(xlDown).row
                    Set sourceRange = Range("A6:AL" & lastRow)
                    sourceRange.Copy
                    Workbooks(zadostName).worksheets("Nstat3").Range("A5").PasteSpecial
                    Workbooks(zadostName).Save

                    Workbooks(fileName(y)).Close saveChanges:=False
                    Workbooks(thisWbName).Activate
                    ThisWorkbook.RefreshAll
                    Workbooks(zadostName).Close saveChanges:=False
                    Exit For
                End If
            Next y
        Else
            Workbooks.Open (zadostName)
            Workbooks(thisWbName).Activate
            ThisWorkbook.RefreshAll
            Workbooks(zadostName).Close saveChanges:=True
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
    End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
