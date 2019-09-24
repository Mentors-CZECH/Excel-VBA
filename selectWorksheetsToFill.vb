Sub selectWorksheetsToFill()
'vybere listy MOS a vybere buňku aktuálního dne
    Dim dayOffset As Integer
    Dim x As Integer
    
    dayOffset = worksheets("dny").Range("C7")
    
    worksheets(9).Activate
    For x = 9 To ThisWorkbook.worksheets.Count
        worksheets(x).Select (False)
        Range("B7").Offset(dayOffset + 1, 0).Select
    Next x
    
End Sub