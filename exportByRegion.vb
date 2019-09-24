'Option Explicit

Private Sub buttonAll_Click()
    
    Dim chk As Control
    Dim colors As String, delimiter As String

    For Each chk In Me.Controls
        If TypeOf chk Is msforms.CheckBox Then
    chk.Value = True
        End If
    Next
    
End Sub

Private Sub buttonInvert_Click()
    For Each chk In Me.Controls
        If TypeOf chk Is msforms.CheckBox Then
            If chk.Value = True Then
            chk.Value = False
            Else
            chk.Value = True
            End If
        End If
    Next
    
End Sub

Private Sub buttonStorno_Click()
    Unload Me
End Sub

Private Sub buttonExport_Click()
    Dim sourceFolderDialog As Object
    Dim wSheet As Worksheet
    Dim wBook, sourceWorkbook As Workbook
    Dim mesic, folderPath, msgVal As String
    Dim regionNum As Integer
    Dim rng As Range
    Dim TEST As Integer

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With

    Set rng = Range("B5:R100").SpecialCells(xlCellTypeBlanks)
    Set sourceFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    Set sourceWorkbook = Workbooks(ActiveWorkbook.Name)

    mesic = StrConv(MonthName(Month(Now)), vbProperCase)
    folderPath = ActiveWorkbook.Path
    
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

    For Each chk In Me.Controls
        If TypeOf chk Is msforms.CheckBox Then
            If chk.Value = True Then
                
                'nastavit export
                
                Set wBook = Workbooks.Add
                With wBook
                    .Title = "skupin_" & mesic & "_" & Right(chk.Name, 1)
                    .Subject = "Region_R0" & Right(chk.Name, 1)
                    .SaveAs fileName:=folderPath & "\Průběžné plnění manažesrkých skupin_" & mesic & "_R0" & Right(chk.Name, 1) & ".xlsx"
                End With
                sourceWorkbook.Charts("Produkty_kod_jmeno").Copy after:=wBook.sheets(wBook.sheets.Count)
                sourceWorkbook.Charts("Plán_kod_jmeno").Copy after:=wBook.sheets(wBook.sheets.Count)
                sourceWorkbook.Charts("Plany_kod_jmeno").Copy after:=wBook.sheets(wBook.sheets.Count)
        
                For Each wSheet In sourceWorkbook.worksheets
                    'Podmínka, podle které se určuje, zda se list kopírovat do cílového souboru. Zde je to číslo regionu uvedené v záhlaví listu
                    If wSheet.Cells(5, 3).Value = "R0" & CStr(Right(chk.Name, 1)) Or wSheet.Name = "Pořadí" Then
                        wSheet.Copy after:=wBook.sheets(wBook.sheets.Count)
                    Else
                    End If
                    
                Next wSheet
                
                worksheets("List1").Delete
                worksheets("Pořadí").Activate
                Range("B2").Clear
                TEST = ActiveSheet.Range("B5").End(xlDown).row
                worksheets("Pořadí").ListObjects("Tabulka1").Unlist
                worksheets("Pořadí").ListObjects("Tabulka2").Unlist
                For i = 5 To TEST
                    If CInt(Cells(i, 2).Value) = CInt(Right(chk.Name, 1)) Then
                    Else
                    Range("B" & i & ":" & "I" & i).ClearContents
                    End If
                    If CInt(Cells(i, 11).Value) = CInt(Right(chk.Name, 1)) Then
                    Else
                    Range("K" & i & ":" & "R" & i).ClearContents
                    End If
                Next
                Range("B5:I" & TEST).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
                Range("K5:R" & TEST).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
                Columns("E:E").HorizontalAlignment = xlCenter
                Columns("M:M").HorizontalAlignment = xlCenter
                wBook.Close savechanges:=True
            End If
        End If
    Next
    Unload Me
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

