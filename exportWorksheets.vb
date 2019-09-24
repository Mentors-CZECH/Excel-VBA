Option Explicit
'Vytvoří nové listy ze seznamu, který se nachází na listu "Seznam Variant" a začíná v buňce "F2" až poslední záznam ve sloupci.

Sub CreateSheetsFromAList()
    Dim MyCell As Range, MyRange As Range
    
    'Seznam variant je název listu, který můžete libovolně měnit
    'Buňka F2 je první záznam v seznamu, který chcete převést na nové listy
    Set MyRange = Sheets("Seznam Variant").Range("D41")
    Set MyRange = Range(MyRange, MyRange.End(xlDown))

    For Each MyCell In MyRange
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = MyCell.Value
    Next MyCell
End Sub

Option Explicit

Sub ExportByRegion()

Dim sourceFileDialog As Object
Dim wSheet As Worksheet, wBook, sourceWorkbook As Workbook
Dim mesic, folderPath, msgVal As String
Dim i As Integer

    With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    End With
    
    Set sourceFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    Set sourceWorkbook = Workbooks(ActiveWorkbook.Name)
    
    mesic = StrConv(MonthName(Month(Now)), vbProperCase)
    folderPath = ActiveWorkbook.Path

    With sourceFileDialog
        .Title = "Destinace pro export"
        .InitialFileName = folderPath
        .AllowMultiSelect = False
        'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou
        If .Show <> -1 Then
            msgVal = MsgBox("Storno export", vbCritical)
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

        For Each wSheet In sourceWorkbook.Worksheets
            Set wBook = Workbooks.Add
            With wBook
                .Title = "TC"
                .Subject = "wSheet.name"
                .SaveAs Filename:=folderPath & "\TC_S_P_01-" & wSheet.Name & ".xlsx"
            End With

            wSheet.Copy After:=wBook.Sheets(wBook.Sheets.Count)
            Worksheets("List1").Delete
            If wSheet.Name = "PopisVariant" Then
            Else
                Range("E1").EntireColumn.Insert
                Range("E1").Value = "Smlouva"
                Rows("1:12").EntireRow.Delete
                Columns("K:T").EntireColumn.Delete

                wSheet.Range("A2:A36").FillDown
            End If
            wBook.Close savechanges:=True
        Next wSheet

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub

