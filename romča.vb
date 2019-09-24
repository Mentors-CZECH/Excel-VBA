Option Explicit
Sub run()
	Dim brandArray(), partArray() As Variant
	Dim buildWorksheet As String
	Dim startCell As Variant
	Dim lowerBoundBrand, upperBoundBrand, lowerBoundPart, upperBoundPart, i, y, z, cPointer As Integer

	Application.ScreenUpdating = False
	cPointer = 0
	brandArray = Array("ALFA ROMEO", "AUDI", "BMW", "CITROËN", "DACIA", "DAEWOO", "DAIHATSU", "FIAT", "FORD", "HONDA", "HYUNDAI", "CHRYSLER", "ISUZU", "JEEP", "KIA", "LANCIA", "LAND ROVER", "LEXUS", "MAZDA", "MERCEDES", "MINI", "MITSUBISHI", "NISSAN", "OPEL", "PEUGEOT", "PONTIAC", "PROTON", "RENAULT", "ROVER", "SAAB", "SEAT", "SMART", "SSANGYONG", "SUBARU", "SUZUKI", "ŠKODA", "TOYOTA", "VOLKSWAGEN", "VOLVO")
	partArray = Array("Karoserie", "Motory", "Převodovky", "Výfuky", "Brzdy", "Chlazení", "Vytápění", "Nápravy", "Elektrosoučástky", "Univerzál", "Interiér", "Palivová", "Kola", "Osvětlení")
	lowerBoundBrand = LBound(brandArray)
	upperBoundBrand = UBound(brandArray)
	lowerBoundPart = LBound(partArray)
	upperBoundPart = UBound(partArray)
	
	For i = 0 To upperBoundBrand
		For y = 0 To upperBoundPart
			Sheets.Add.Name = brandArray(i) & " " & partArray(y)
			buildWorksheet = brandArray(i) & " " & partArray(y)
			Worksheets(buildWorksheet).Range("B2:B600") = brandArray(i)
			Range("B2").Activate
			For z = 0 To 5
			ActiveCell.Offset(0, z).FormulaR1C1 = "=značky!R[" & y & "]C[" & cPointer + 2 & "]"
			Next z
			z = 0
			cPointer = 0
		Next y
	Next i
	
	Application.ScreenUpdating = True
End Sub