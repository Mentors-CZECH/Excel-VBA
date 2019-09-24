Option Explicit

Sub A_EXPORT_DATA()
Dim sourceFolderDialog As Object
Dim wSheet As Worksheet
Dim wBook, sourceWorkbook As Workbook
Dim mesic, folderPath, msgVal, lastNameStr, buildSheetName As String
Dim regionNum, lastName, lastRow, lastCol, EndRow, EndCol, EndRow2, endSheet As Integer
Dim rng, copyrange As Range
Dim TEST As Integer
Dim productRowArray() As Variant
Dim item As Variant


' Nastavení aplikační vrstvy
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    
    'definice pole pro offset jednotlivých produktů a následné skrývání řádků
    productRowArray() = Array(9, 37, 65, 93, 121)
    Set rng = Range("B5:R100").SpecialCells(xlCellTypeBlanks)
    Set sourceFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    Set sourceWorkbook = Workbooks(ActiveWorkbook.Name)

    mesic = StrConv(MonthName(Month(Now)), vbProperCase)
    folderPath = ActiveWorkbook.Path
    'Dialog pro výběr destinace (folderPath) složky, do které se bude ukládat exportovaný soubor.
    With sourceFolderDialog
        .Title = "Destinace pro export"
        .AllowMultiSelect = False

        If .Show <> -1 Then
            msgVal = MsgBox("Storno export", vbCritical)
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

                Set wBook = Workbooks.Add
                With wBook
                    .Title = "Průběžné plnění "
                    .Subject = "Průběžné plnění skupin"
                    .SaveAs fileName:=folderPath & "\Průběžné plnění manažesrkých skupin_9.xlsx"
                End With
                ActiveWindow.DisplayGridlines = False
                For Each wSheet In sourceWorkbook.worksheets
                    If wSheet.Cells(5, 3).Value = "R01" Or wSheet.Cells(5, 3).Value = "R02" Or wSheet.Cells(5, 3).Value = "R03" Or wSheet.Cells(5, 3).Value = "R04" Or wSheet.Cells(5, 3).Value = "R05" Or wSheet.Cells(5, 3).Value = "R06" Or wSheet.Name = "Pořadí" Then
                        Set copyrange = wSheet.Range("A1:BJ147")
                        copyrange.Copy
                        wBook.sheets.Add after:=worksheets(worksheets.Count)
                        
                        If wSheet.Name = "Pořadí" Then
                            worksheets(worksheets.Count).Range("A1").PasteSpecial Paste:=xlPasteAll
                            wSheet.Activate
                            Cells.Select
                            Selection.Copy
                            wBook.Activate
                            Range("A1").Select
                            ActiveSheet.Paste
                            Range("B2").ClearContents
                            ActiveSheet.Name = "Pořadí"
                            
                        Else
                        
                            With worksheets(worksheets.Count).Range("A1")
                                .PasteSpecial xlPasteValues
                                .PasteSpecial xlPasteFormats
                                
                            End With
                            lastNameStr = Range("C3")
                            lastName = Len(lastNameStr) - InStr(lastNameStr, " ")
                            lastName = Right(lastNameStr, lastName)
                            buildSheetName = Range("C2").Value & " - " & lastName
                            
                            If Len(buildSheetName) > 31 Then
                            ActiveSheet.Name = Left(buildSheetName, 31)
                            
                            Else
                            ActiveSheet.Name = Range("C2").Value & " - " & lastName
                            
                            End If
                            If ActiveSheet.Range("C5") = "R01" Then
                                ActiveSheet.Tab.ColorIndex = 23
                                
                            ElseIf ActiveSheet.Range("C5") = "R02" Then
                                ActiveSheet.Tab.ColorIndex = 27
                                
                            ElseIf ActiveSheet.Range("C5") = "RO3" Then
                                ActiveSheet.Tab.ColorIndex = 10
                                
                            ElseIf ActiveSheet.Range("C5") = "R04" Then
                                ActiveSheet.Tab.ColorIndex = 3
                                
                            ElseIf ActiveSheet.Range("C5") = "R05" Then
                                ActiveSheet.Tab.ColorIndex = 16
                                
                            Else
                                ActiveSheet.Tab.ColorIndex = 46
                                
                            End If
                            lastRow = Range("A1").End(xlDown).row
                            lastCol = Range("D1").End(xlToRight).Column
                            EndCol = 62
                            EndRow = 160
                            ActiveSheet.Cells(1, EndCol + 1).Resize(, lastCol - EndCol).Columns.Hidden = True

                            For Each item In productRowArray
                                EndRow = ActiveSheet.Range("B" & item).End(xlDown).row
                                EndRow2 = (ActiveSheet.Range("B" & EndRow).End(xlDown).row - EndRow) - 2
                                Range("B" & EndRow & ":B" & EndRow2 + EndRow).EntireRow.Hidden = True
                                
                            Next item
                            Range("B147:B1048576").EntireRow.Hidden = True
                            
                        End If
                        ActiveWindow.Zoom = 70
                        Columns.AutoFit
                        Range("C2").Select

                    Else
                    End If
                    
                Next wSheet
                
                wBook.worksheets("List1").Delete
                wBook.worksheets("Pořadí").Activate
                wBook.Save
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

