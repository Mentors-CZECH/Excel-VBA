Option Explicit

Private Sub CopyDataFromSources()
'Skripty slou�� pro zad�n� cesty k souboru report SD salesreport. Je povolen v�b�r pouze jednoho souboru a defaultn� filtr je nastaven jako MS Excel 93
'Pot� jsou kop�rov�na data ze souboru report SD salesreport do souboru Sales Report, na p��slu�n� listy. Skript neumo��uje na��tat v�t�� mno�stv� soubor� (pokud
'chcete vybrat v�ce soubor� p�epi�te AllowMultiSelect = TRUE a na��tejte jednotliv� Itemy do pole).

Dim sourceFilePath As String
Dim sourceFileDialog As Object
Dim sourceWorkbookName, targetWorkbookName, targetWorksheetName As String
Dim copyRange, copyRange2 As Variant

    Application.ScreenUpdating = False
    
    'P�i�azen� n�zv� se��tu a listu
    targetWorkbookName = ActiveWorkbook.Name
    targetWorksheetName = "Counties statistics - country"
    
    'Vytvo�� objekt sourceFileDialog jako Application,FileDialog(msoFileDialogFilePicker). Vlastnosti objektu jsou:
    'Title(nadpis v z�hlav� objektu)
    'Filters.Clear - vy�ist� v�echny defaultn� filtry p�i�azen� k objektu
    'Filters.Add p��dan� dvou filtr� pro soubor MS Excel 93-13 (defaultn� filtr je 93)
    'AllowmultiSelect = FALSE - neumo��ujeme na��tat v�ce ne� jeden soubor.
    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    With sourceFileDialog
        .Title = "Zadejte cestu k souboru: report SD salesreport"
        .Filters.Clear
        .Filters.Add "Soubory MS Excel 93", "*.xls"
        .Filters.Add "Soubory MS Excel 2007 - 2013", "*.xlsm"
        .FilterIndex = 1
        .ButtonName = "Na��st data"
        .AllowMultiSelect = False
        
        'Pokud u�ivatel klikne na storno, zobrzit promt - nebyly vypln�ny listy
        If .Show <> -1 Then
            MsgBox "Storno - N�kter� listy nejsou vypln�ny"
            Exit Sub
        End If
        
    'P�i�azen� cesty k souboru reports SD salesreport do prom�nn�
        sourceFilePath = .SelectedItems(1)
    End With
    
    'Otev�en� souboru report SD salesreport
    sourceWorkbookName = Workbooks.Open(sourceFilePath).Name
    
    'Kop�rov�n� dat
    copyRange = Worksheets("souhrn �R").Range("A9:AP87").Value
    Workbooks(targetWorkbookName).Worksheets(targetWorksheetName).Range("A9:AP87") = copyRange
    'Kop�rov�n� dat
    copyRange2 = Workbooks(sourceWorkbookName).Worksheets("souhrn regiony").Range("A9:AN16").Value
    Workbooks(targetWorkbookName).Worksheets("Counties statistics per regions").Range("A9:AN16") = copyRange2

    'Zav�en� se�itu otev�en�ho pomoc� dialogu
    Workbooks(sourceWorkbookName).Close
    'Vyma�e objekt sourceFileDialog z pam�ti
    Set sourceFileDialog = Nothing
    Application.ScreenUpdating = True
End Sub

