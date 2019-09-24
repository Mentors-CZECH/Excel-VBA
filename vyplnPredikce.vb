Option Explicit

Sub A_VYPLNIT_PREDIKCE()

Dim querySheet As Worksheet
Dim stat2Query As QueryTable
Dim sourceFileDialog As Object
Dim sourceRange As Range

Dim dirPath, monthType, newFileName, dayType As String
Dim thisWbName As String
Dim lastrow, x, t, z, y As Integer
Dim filePath, fileName, msgVal, zadostiMsg As String
Dim productArray(), stat9array(), rowsArray() As Variant

Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)
Set querySheet = worksheets("Stat2")
Set stat2Query = querySheet.ListObjects.Item(1).QueryTable

    Call odkrytiMKTzdroje

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Definice jednotlivých produktů. productArray obsahuje názvy jednotlivých produktů, které se vyhledávají v prvním sloupečku tabulky a podle nalezených záznamů
'   se přiřazují řádky do pole rowsArray()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    productArray() = Array("Produkty AEFG", "Produkt A", "Produkt E", "Produkt F", "Produkt G")
    stat9array() = Array("AC", "A", "E", "F", "G", "ABCDEFG")
    ReDim rowsArray(0 To 4)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Dotaz, zda vytvořit nový soubor pro tvorbu predikce, jelikož u predikce se mažou vzorce ve STAT3, pokud už uživatel má dokument predikce
'   vytvořený, může pokračovat s makrem v aktuálním dokumentu.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    dirPath = Application.ActiveWorkbook.Path
    thisWbName = ActiveWorkbook.Name

    stat2Query.Refresh False
    DoEvents

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Dotaz na uživatele pro zadání souboru se STAT09, který bude použit při vyplňování predikce
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    With sourceFileDialog
        .Title = "Zadejte cestu k souboru: STAT09"
        .Filters.Clear
        .Filters.Add "Soubory MS Excel", "*.xl*"
        .ButtonName = "Načíst data"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            msgVal = MsgBox("Storno - Načítání dat a skript přerušen. Žádná data nebyla do souboru načtena.", vbCritical)
            Exit Sub
        End If
        filePath = .SelectedItems(1)
        fileName = Dir(filePath)
    End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Vyhledání posledního řákdu, který obsahuje data a provedení výpočtu offsetu pro jednotlivé produkty a jejich uložení do pole. Ošetřen je i případ, kdy uživatel vyplňuje
'   predikci pro první den v měsíci.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    worksheets(1).Activate
    lastrow = Cells(3, "D").End(xlDown).Row
    lastrow = lastrow - 1
    'pokud je tabulka prázdná bude lastRow 42 - v tom případě nastavuji první řádek v tabulce
    If lastrow = 41 Then
        lastrow = 2
    End If

    worksheets(1).Activate
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Mazání stat3 a stat9 'Kopírování STAT09 a STAT02 do STAT09 dle jednotlivých produktů. Postupně se projdou všechny produkty dle stat9array.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    worksheets("Nstat3").Range("AL1:AT5,AL8:AN14").ClearContents
    worksheets("Nstat9").Range("A1:AA79").ClearContents
    worksheets("Nstat9").Range("M6:N15,V6:W15").UnMerge

    For x = 0 To 5
    
        Workbooks.Open (filePath)
        worksheets(stat9array(x)).Activate
        'Produkty A-G (0=A 1=E 2=F 3=G 4=ABCEFG (Přepisování STAT09 a kopírování STAT05)
        If x = 0 Then
            'data pro AEFG
            Workbooks(fileName).Close saveChanges:=False
            Set sourceRange = worksheets("Stat2").Range("B5:B12")
            sourceRange.Copy
            worksheets("Nstat9").Range("M6").PasteSpecial xlPasteValues
            Set sourceRange = worksheets("Stat2").Range("F5:F12")
            sourceRange.Copy
            worksheets("Nstat9").Range("V6").PasteSpecial xlPasteValues
            For z = 1 To 8
                worksheets(z).Select (False)
                Cells(rowsArray(x), 2).Activate
            Next z
            Call prepisNstat9_predikce
            
        ElseIf x = 1 Then
            'data pro A (B)
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName).Close saveChanges:=False
            worksheets("Nstat9").Range("M6:N15,V6:W15").UnMerge
            Set sourceRange = worksheets("Stat2").Range("J5:J12")
            sourceRange.Copy
            worksheets("Nstat9").Range("M6").PasteSpecial xlPasteValues
            Set sourceRange = worksheets("Stat2").Range("F5:F12")
            sourceRange.Copy
            worksheets("Nstat9").Range("V6").PasteSpecial xlPasteValues
            worksheets(1).Activate
            For z = 1 To 8
                worksheets(z).Select (False)
                Cells(rowsArray(x), 2).Activate
            Next z
            Call prepisNstat9_predikce
            
        ElseIf x = 2 Then
            'data pro E (N)
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName).Close saveChanges:=False
            worksheets("Nstat9").Range("M6:N15,V6:W15").UnMerge
            Set sourceRange = worksheets("Stat2").Range("N5:N12")
            sourceRange.Copy
            worksheets("Nstat9").Range("M6").PasteSpecial xlPasteValues
            worksheets(1).Activate
            For z = 1 To 8
                worksheets(z).Select (False)
                Cells(rowsArray(x), 2).Activate
            Next z
            Call prepisNstat9_predikce
            
        ElseIf x = 3 Then
            'data pro F (R)
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName).Close saveChanges:=False
            worksheets("Nstat9").Range("M6:N15,V6:W15").UnMerge
            Set sourceRange = worksheets("Stat2").Range("R5:R12")
            sourceRange.Copy
            worksheets("Nstat9").Range("M6").PasteSpecial xlPasteValues
            worksheets(1).Activate
            For z = 1 To 8
                worksheets(z).Select (False)
                Cells(rowsArray(x), 2).Activate
            Next z
            Call prepisNstat9_predikce
            
        ElseIf x = 4 Then
            'data pro G (V)
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName).Close saveChanges:=False
            
            worksheets("Nstat9").Range("M6:N15,V6:W15").UnMerge
            Set sourceRange = worksheets("Stat2").Range("V5:V12")
            sourceRange.Copy
            worksheets("Nstat9").Range("M6").PasteSpecial xlPasteValues
            worksheets(1).Activate
            For z = 1 To 8
                worksheets(z).Select (False)
                Cells(rowsArray(x), 2).Activate
            Next z
            Call prepisNstat9_predikce
            
        ElseIf x = 5 Then
            Set sourceRange = Range("A1:AA79")
            sourceRange.Copy
            Workbooks(thisWbName).worksheets("Nstat9").Range("A1").PasteSpecial
            Workbooks(fileName).Close saveChanges:=False
        Exit For
        End If
    Next x

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Dotaz zda chce uživatel vytvořit soubor predikce. Vytváří se z konstatního názvu souboru a z aktuálního data dle Now(). Nastavení aktivního listu, ukončení skriptu.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Sheets("ČR CELKEM").Activate
    zadostiMsg = MsgBox("Uložit jako nový soubor predikce?", vbYesNo + vbQuestion, "Uložit jako?")

    If zadostiMsg = vbYes Then
        If Len(Month(Now)) = 1 Then
        monthType = "0" & Month(Now)
        Else
        monthType = Month(Now)
        End If
        
        If Len(Day(Now)) = 1 Then
        dayType = "0" & Day(Now)
        Else
        dayType = Day(Now)
        End If

        newFileName = "Plán a skutečnost v regionech predikce k " & dayType & "." & monthType & "." & Year(Now)
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & newFileName & ".xlsm"
    End If

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

    Call odkrytiMKTzdroje
End Sub