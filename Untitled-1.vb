
Sub buildData()

        Dim MOS                                                       As Range
        Dim wks                                                         As Worksheet
        Dim sheetName                                            As String
        Dim sourceFolderDialog                              As Object
        Dim wSheet                                                    As Worksheet
        Dim wBook, sourceWorkbook                     As Workbook
        Dim folderPath, msgVal                                As String
        Dim regionNum                                             As Integer

        'Neukazovat obnovování obrazovky a chybové hlášky, dotazy
        With Application
                .ScreenUpdating = False
                .DisplayAlerts = False
        End With
        'Lastrow obsahuje poslední řádek ve slupci B (nesmí obsahovat prázdné řádky). Skript dělá to semé jako by jsi klikla na B4 a udělala CTRL+SHIFT+Šipku dolu.
        'Vlastnost Range.Row vrací řádek, na kterém je kurzor. Kontroluj si zda ti to opravdu vrátí správný řádek. Například pokud tam budeš mít vzorec, který vrací "", tak to bude pořád brát jako vyplňěnou buňku
        lastrow = Sheets("Model").Range("B4").End(xlDown).Row
        'Nastaví rozsah, který se bude prohledávat jako B4 až poslední řádek ve sloupci B
        Set MOS = Sheets("Model").Range("B4:B" & lastrow)
        'Vytvoření objektu pro výběr destinace, kam chceme soubor uložit.
        Set sourceFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
        'uloží právě otevřený soubor jako sourceWorkbook. ten se potom používá při přepínání mezi jednotlivýma souborama, které budeme exportovat.
        Set sourceWorkbook = Workbooks(ActiveWorkbook.Name)
    
        'Nastavení dialogového okna pro export
        With sourceFolderDialog
                .Title = "Destinace pro export"
                .AllowMultiSelect = False
                'Pokud uživatel klikne na storno, ukaž hlášku, a ukonči makro
                If .Show <> -1 Then
                        msgVal = MsgBox("Storno export", vbCritical)
                        Exit Sub
                End If
                'do proměnné folderPath uloží cestu ke složce, kterou uživatel vybral.
                folderPath = .SelectedItems(1)
        End With

        For Each Item In MOS
                'Item nemůže být prázdná hodnota (vzorec?)
                If Item.Value = "" Then
                        Exit For
                End If
                
                sheetName = Item.Value
                Sheets("Výstup MNG").Range("B5") = sheetName
                Sheets("Výstup MNG").Copy after:=Sheets(Sheets.Count)
                
                With ActiveSheet.UsedRange
                .Value = .Value
                End With
                
                ActiveSheet.Rows("1:3").Select
                Selection.Delete Shift:=xlUp
                ActiveSheet.Shapes.Range(Array("ComboBox1")).Select
                Selection.Delete
                
                Set wSheet = ActiveSheet
                ActiveSheet.Name = sheetName
                Set wBook = Workbooks.Add
                
                With wBook
                .SaveAs Filename:=folderPath & "\Rozvojový plán - " & sheetName & ".xlsx"
                End With
                
                sourceWorkbook.Sheets(sheetName).Copy after:=wBook.Sheets(wBook.Sheets.Count)
                Worksheets("List1").Delete
                wBook.Close savechanges:=True
                sourceWorkbook.Sheets(sheetName).Delete
        Next Item

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub

