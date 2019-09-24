
Sub buildData()

Dim TM                     As Range
Dim UP                     As Range
Dim wks                     As Worksheet
Dim sheetName               As String
Dim sourceFolderDialog      As Object
Dim wSheet                  As Worksheet
Dim wBook, sourceWorkbook   As Workbook
Dim folderPath, msgVal      As String
Dim regionNum               As Integer
Dim lastrow              As Integer

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    Set TM = Sheets("PT").Range("A4:A6")
    Set sourceFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    Set sourceWorkbook = Workbooks(ActiveWorkbook.Name)
    Set UP = Sheets("PT").Range("D6:D6")
    
        With sourceFolderDialog
                .Title = "Destinace pro export"
                .AllowMultiSelect = False
                If .Show <> -1 Then
                        msgVal = MsgBox("Storno export", vbCritical)
                        Exit Sub
                End If
                folderPath = .SelectedItems(1)
        End With
'Přepsal jsem Item na arrayItem, aby mi bylo jasný že to je prvek pole, které procházím
        For Each arrayItem In TM
        
                If arrayItem.Value = "" Then
                        Exit For
                End If
                sheetName = arrayItem.Value
                Sheets("TM").Range("F6") = sheetName
                
                caPointer = 1
                'Vyhledávám arrayItem (Kód TM) v rozsahu C:C na listu PT. Vyhledávání tam je proto, že nemám jistotu, že pořadí arrayItem = pořadí položek v sheet("PT").Range("A:A")
                resultRow = Sheets("PT").Range("C:C"). _
                Find(What:=arrayItem, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False).Row
                        
                'Vnitřní smyčka. Dokud se kód ve sloupci C:C rovná kódu, který máme v poli(tmArray), tak načítej kódy ÚP

                Do While Sheets("PT").Range("C" & resultRow) = arrayItem
                        'vytváří jméno listu ÚP a caPointer. Do rozsahu B1 vkládá kód TM
                        Sheets("ÚP" & caPointer).Range("B1") = sheetName
                        'vytváří jméno listu ÚP a caPointer. Do rozsahu F6 vkládá kód ÚP
                        Sheets("ÚP" & caPointer).Range("F6") = Sheets("PT").Range("D" & resultRow)
                        'msgbox je tu pro debug. Když si ho odkomentuješ, tak bude vracet kódy jednotlivých ÚP. Když si pouštím makro, tak chci vidět, jestli smyčka opravdu vrací všechny kódy.
                        'MsgBox Range("D" & resultRow)
                        'result row tu používám jako index pro řádky. Při každém průchodu smyčkou zvýšíme o +1 -> v přístím průchodu smyčkou se posuunu o řádek dolu.
                        resultRow = resultRow + 1
                        'caPointer určuje pořadí listu ÚP & X
                        caPointer = caPointer + 1
                Loop
        
'                Sheets("TM").Copy after:=Sheets(Sheets.Count)
                
                Set wSheet = ActiveSheet
                ActiveSheet.Name = sheetName
                
                Set wBook = Workbooks.Add
                With wBook
                        .SaveAs Filename:=folderPath & "\Akční plánovač - " & sheetName & ".xlsx"
                End With
                
                sourceWorkbook.Sheets(sheetName).Copy after:=wBook.Sheets(wBook.Sheets.Count)
                Worksheets("List1").Delete
                wBook.Close savechanges:=True
                sourceWorkbook.Sheets(sheetName).Delete

        Next arrayItem
        
        For caPointer = 1 To 50
        
        Next caPointer
        
        
        With Application
                .ScreenUpdating = True
                .DisplayAlerts = True
        End With
End Sub

