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

Sub ExportBySheetItem()

Dim sourceFileDialog As Object
Dim wSheet As Worksheet, wBook, sourceWorkbook As Workbook
Dim mesic, folderPath, msgVal As String
Dim i, lastrow, nextrow, overflow, test As Long

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
            'řádky které budou smazány
            If Cells(2, 1).Value = "" Then
             Columns("A").EntireColumn.Delete
             End If
                Rows("1:11").EntireRow.Delete
                Columns("J:T").EntireColumn.Delete
                Range("E1").EntireColumn.Insert
                Range("E1").Value = "Smlouva"
                Cells(1, 1).Activate
                overflow = 0
                For i = 1 To 50
                If ActiveCell.Value = "" Then
                Exit For
                Else
                    lastrow = Cells(1, "A").End(xlDown).Row
                    On Error Resume Next:
                    nextrow = Cells(lastrow, "A").End(xlDown).Row
                    If nextrow > 10000 Then
                    overflow = Cells(lastrow, "C").End(xlDown).Row
                    Range("A" & lastrow + 1 & ":B" & overflow).Value = Range("A" & lastrow & ":B" & lastrow).Value
                    nextrow = overflow
                    Else
                    Range("A" & lastrow + 1 & ":B" & nextrow - 1).Value = Range("A" & lastrow & ":B" & lastrow).Value
                    End If
                End If
                
                If Cells(nextrow + 1, 1) = "" Then
                    nextrow = nextrow - 1
                Else
                    test = Cells(nextrow - 1, 1).End(xlDown).Row
                    nextrow = test
                End If
                
                Cells(nextrow + overflow, 1).Activate
                Next i
            End If
            wBook.Close savechanges:=True
        Next wSheet

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub




