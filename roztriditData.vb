Sub A_ROZTRIDIT_DATA()

Dim pasteRange  As Range
Dim copyRange   As Range
Dim lastrow     As Integer

    With Application
        .ScreenUpdating = False
        .CutCopyMode = False
    End With

        Sheets("Přehled Změn OS").Activate
        lastrow = Worksheets("Přehled Změn OS").Range("B" & Rows.Count).End(xlUp).Row
        lastrow = lastrow - 1

        Set copyRange = Range("B4:O" & lastrow)
        copyRange.Copy
        Set pasteRange = Range("B4:O" & lastrow)
        pasteRange.PasteSpecial xlPasteValues

        lastrow = Worksheets("Přehled Změn OS").Range("B" & Rows.Count).End(xlUp).Row
        'Nástupy MOS - kopíruje data na separátní list
        Sheets("Přehled Změn OS").Range("$B$3:$O$" & lastrow).AutoFilter Field:=6, Criteria1:="Nástup"
        Sheets("Přehled Změn OS").AutoFilter.Range.Copy
        Sheets("Nástup").Activate
        Sheets("Nástup").Range("C3").Activate
        Sheets("Nástup").Paste
        Columns.AutoFit

        Columns("B:B").EntireColumn.Hidden = True
        Range("A1").Select
        Sheets("Přehled Změn OS").Range("$B$3:$O$" & lastrow).AutoFilter Field:=6, Criteria1:="Změna primární oblasti"
        Sheets("Přehled Změn OS").AutoFilter.Range.Copy
        Sheets("Změna primární oblasti").Activate
        Sheets("Změna primární oblasti").Range("B3").Activate
        Sheets("Změna primární oblasti").Paste
        Columns.AutoFit

        Range("A1").Select
        Sheets("Přehled Změn OS").Range("$B$3:$O$" & lastrow).AutoFilter Field:=6, Criteria1:="Změna sekundární oblasti"
        Sheets("Přehled Změn OS").AutoFilter.Range.Copy
        Sheets("Změna sekundární oblasti").Activate
        Sheets("Změna sekundární oblasti").Range("B3").Activate
        Sheets("Změna sekundární oblasti").Paste
        Columns.AutoFit

        Range("A1").Select
        Sheets("Přehled Změn OS").Range("$B$3:$O$" & lastrow).AutoFilter Field:=6, Criteria1:="Změna pozice"
        Sheets("Přehled Změn OS").AutoFilter.Range.Copy
        Sheets("Změna pozice").Activate
        Sheets("Změna pozice").Range("B3").Activate
        Sheets("Změna pozice").Paste
        Columns.AutoFit

        Range("A1").Select
        Sheets("Přehled Změn OS").Range("$B$3:$O$" & lastrow).AutoFilter Field:=6, Criteria1:="Ukončení"
        Sheets("Přehled Změn OS").AutoFilter.Range.Copy
        Sheets("Ukončení").Activate
        Sheets("Ukončení").Range("B3").Activate
        Sheets("Ukončení").Paste
        Columns.AutoFit

        Range("A1").Select
        Sheets("Přehled Změn OS").Range("$B$3:$O$" & lastrow).AutoFilter Field:=6
        Sheets("Přehled Změn OS").Activate

    With Application
        .ScreenUpdating = True
        .CutCopyMode = True
    End With

End Sub