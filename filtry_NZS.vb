'Option Explicit

Function FILTR_PRIMARNI_OBLASTI(mosCode As Range, filterDate As Range)
Dim item As Range
Dim endRow, i As Integer
Dim dataRange As Range
Dim changeDate() As Variant
Dim changeType() As Variant
Dim changeRow() As Variant

endRow = Sheets("Přehled Změn OS").Range("B2").End(xlDown).Row
Set dataRange = Sheets("Přehled Změn OS").Range("B2:B" & endRow)

i = 0
'Najdi všechny položky pro daný kód

For Each item In dataRange
    If item = mosCode Then
        itemsInArr = itemsInArr + 1
        itemrow = item.Row
        ReDim Preserve changeType(itemsInArr - 1)
        ReDim Preserve changeDate(itemsInArr - 1)
        ReDim Preserve changeRow(itemsInArr - 1)
        changeRow(i) = itemrow
        changeType(i) = Sheets("Přehled Změn OS").Range("G" & itemrow).Value
        changeDate(i) = Sheets("Přehled Změn OS").Range("D" & itemrow).Value
        i = i + 1
    End If
Next item

itemPointer = 0
For Each arrItem In changeDate

    If arrItem >= filterDate Then
        Exit For
    Else
    End If
    itemPointer = itemPointer + 1
Next arrItem

If itemPointer = 0 Then
FILTR_PRIMARNI_OBLASTI = Sheets("Přehled Změn OS").Range("I" & changeRow(0))
Else
FILTR_PRIMARNI_OBLASTI = Sheets("Přehled Změn OS").Range("I" & changeRow(itemPointer - 1))
End If

End Function


Function FILTR_MOS(mosCode As Range, filterDate As Range)

Dim item As Range
Dim endRow, i As Integer
Dim dataRange As Range
Dim changeDate() As Variant
Dim changeType() As Variant
Dim changeRow() As Variant

endRow = Sheets("Přehled Změn OS").Range("B2").End(xlDown).Row
Set dataRange = Sheets("Přehled Změn OS").Range("B2:B" & endRow)

i = 0
'Najdi všechny položky pro daný kód

For Each item In dataRange
    If item = mosCode Then
        itemsInArr = itemsInArr + 1
        itemrow = item.Row
        ReDim Preserve changeType(itemsInArr - 1)
        ReDim Preserve changeDate(itemsInArr - 1)
        ReDim Preserve changeRow(itemsInArr - 1)
        changeRow(i) = itemrow
        changeType(i) = Sheets("Přehled Změn OS").Range("G" & itemrow).Value
'        changeDate(i) = Sheets("Přehled Změn OS").Range("D" & itemrow).Value
        changeDate(i) = DateSerial(Year(Sheets("Přehled Změn OS").Range("D" & itemrow).Value), Month(Sheets("Přehled Změn OS").Range("D" & itemrow).Value), Day(1))
        i = i + 1
    End If
    
    
'    DateSerial(Year(Sheets("Přehled Změn OS").Range("D" & itemrow).Value), Month(Sheets("Přehled Změn OS").Range("D" & itemrow).Value), Day(1))
Next item

itemPointer = 0
For Each arrItem In changeDate
    If arrItem > DateSerial(Year(filterDate), Month(filterDate + 1), Day(1)) Then
        Exit For
    Else
    End If
    itemPointer = itemPointer + 1
Next arrItem


FILTR_MOS = Sheets("Přehled Změn OS").Range("F" & changeRow(itemPointer - 1))

End Function
