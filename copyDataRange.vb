
Sub copyDataRange()
	dim sourceRange as Range
	Dim lastrow as Integer
	
	lastrow = Cells(6, "F").End(xlDown).Row
	sourceRange = Range("BF3:CN67" & lastrow).Copy
	Workbooks(thisWbName).Worksheets("Nstat3").Range("AU3").PasteSpecial xlPasteValues
End Sub

'Pokud první verze nepujde tak tohle. Nevím proč musí být Set před setrange
Sub copyDataSetRange()
	dim sourceRange as Range
	Dim lastrow as Integer
	
	lastrow = Cells(6, "F").End(xlDown).Row
	set sourceRange = Range("BF3:CN67" & lastrow)
	sourceRange.copy
	Workbooks(thisWbName).Worksheets("Nstat3").Range("AU3").PasteSpecial xlPasteValues
End Sub
