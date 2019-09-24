
Sub buildData()

Dim MOS As Range
Dim wks As Worksheet
Dim sheetName As String
Dim sourceFolderDialog As Object
Dim wSheet As Worksheet
Dim rng As Range
Dim wBook, sourceWorkbook As Workbook
Dim mesic, folderPath, msgVal As String
Dim regionNum As Integer
Dim TEST As Integer


    With Application
        .ScreenUpdating = False
    End With
    
    Set MOS = Sheets("Model").Range("B4:B100")
    Set sourceFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    Set sourceWorkbook = ActiveWorkbook
    
    
        With sourceFolderDialog
        .Title = "Destinace pro export"
        .AllowMultiSelect = False
        'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou
        If .Show <> -1 Then
            msgVal = MsgBox("Storno export", vbCritical)
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    
    
    
    For Each Item In MOS
        If Item.Value = "" Then
            Exit For
        End If
        sheetName = Item.Value
        Sheets("Výstup MNG").Range("B5") = sheetName
        Sheets("Výstup MNG").Copy after:=Sheets(Sheets.Count)
        
        With ActiveSheet.UsedRange
            .Value = .Value
        End With
        Set wSheet = ActiveSheet
        ActiveSheet.Name = sheetName
        folderPath = ActiveWorkbook.Path

        Set wBook = Workbooks.Add
        With wBook
            .SaveAs Filename:=folderPath & "\1Rozvojový plán - " & sheetName & ".xlsx"
        End With
                sourceWorkbook.Sheets(Sheets.Count).Copy after:=wBook.Sheets(wBook.Sheets.Count)
        Worksheets("List1").Delete
        wBook.Close savechanges:=True

    Next Item
    
    With Application
        .ScreenUpdating = True
    End With

End Sub
