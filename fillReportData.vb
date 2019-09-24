Sub denni_vyplneni()
    Dim wSheet As Worksheet
' cela_operace_pokus Makro

    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If

     ActiveCell.Offset(0, 7).Range("a1").Select

    'PRODUKTY ABCEF

    'nová produkce
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C36:R120C44,R6C8,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C36:R120C44,R6C8,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 10).Range("a1").Select
    
        'nová produkce-vlastní
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R6C18,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R6C18,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'revolvingy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Revolvingy!C28:C29,2,FALSE)),0,VLOOKUP(R2C3,Revolvingy!C28:C29,2,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'smlouvy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R124C36:R400C50,R6C28,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R124C36:R400C50,R6C28,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 6).Range("a1").Select

    'žádosti
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R6C33,FALSE)),0,VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R6C33,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]+R[-1]C"
    'ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",RC[1],R[-1]C),IF(RC[1]="""","""",IF(R[-1]C=""Skutečnost"",RC[1],RC[1]+R[-1]C)))"
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",IF(ISERROR(RC[1]),0,RC[1]),R[-1]C),IF(RC[1]="""","""",IF(OR(R[-1]C=""Skutečnost"",R[-1]C="""")=TRUE,RC[1],RC[1]+R[-1]C)))"
    ActiveCell.Offset(0, 6).Range("a1").Select
   
   'schvalovatelnost
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R152C48:R300C55,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C48:R300C55,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C48:R300C55,4,FALSE))),0,VLOOKUP(R2C3,Nstat3!R152C48:R300C55,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C48:R300C55,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C48:R300C55,4,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 3).Range("A1").Select
    
    'recidivy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,nstat12!C7:C46,27,FALSE)),0,VLOOKUP(R2C3,nstat12!C7:C46,27,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,nstat12!C7:C46,30,FALSE)),0,VLOOKUP(R2C3,nstat12!C7:C46,30,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,nstat12!C7:C46,33,FALSE)),0,(VLOOKUP(R2C3,nstat12!C7:C46,33,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Nábor
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,NSTAT13!C[-25]:C[-23],3,FALSE)),0,VLOOKUP(R2C3,NSTAT13!C[-25]:C[-23],3,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Efektivní
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,NSTAT13!C32:C35,3,FALSE)),0,VLOOKUP(R2C3,NSTAT13!C32:C35,3,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'PRODUKT A
       
    ActiveCell.Offset(28, -49).Range("a1").Select
    
    'nová produkce
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C36:R120C44,R34C8,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C36:R120C44,R34C8,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 10).Range("a1").Select
    
    'nová produkce-vlastní
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R34C18,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R34C18,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'revolvingy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Revolvingy!C28:C29,2,FALSE)),0,VLOOKUP(R2C3,Revolvingy!C28:C29,2,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'smlouvy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R124C36:R400C50,R34C28,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R124C36:R400C50,R34C28,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 6).Range("a1").Select

    'žádosti
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R34C33,FALSE)),VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R34C34,FALSE),VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R34C33,FALSE)+VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R34C34,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]+R[-1]C"
    'ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",RC[1],R[-1]C),IF(RC[1]="""","""",IF(R[-1]C=""Skutečnost"",RC[1],RC[1]+R[-1]C)))"
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",IF(ISERROR(RC[1]),0,RC[1]),R[-1]C),IF(RC[1]="""","""",IF(OR(R[-1]C=""Skutečnost"",R[-1]C="""")=TRUE,RC[1],RC[1]+R[-1]C)))"
    ActiveCell.Offset(0, 6).Range("a1").Select
    
    'schvalovatelnost
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R152C55:R300C59,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C55:R300C59,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C55:R300C59,4,FALSE))),0,VLOOKUP(R2C3,Nstat3!R152C55:R300C59,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C55:R300C59,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C55:R300C59,4,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    'PRODUKT E
       
    ActiveCell.Offset(28, -31).Range("a1").Select
    
    'nová produkce
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C36:R114C44,R62C8,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C36:R114C44,R62C8,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 10).Range("a1").Select
    
            'nová produkce-vlastní
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R62C18,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R62C18,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'revolvingy
 
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'smlouvy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R117C36:R245C50,R62C28,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R117C36:R245C50,R62C28,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 6).Range("a1").Select

    'žádosti
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R62C33,FALSE)),0,VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R62C33,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]+R[-1]C"
    'ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",RC[1],R[-1]C),IF(RC[1]="""","""",IF(R[-1]C=""Skutečnost"",RC[1],RC[1]+R[-1]C)))"
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",IF(ISERROR(RC[1]),0,RC[1]),R[-1]C),IF(RC[1]="""","""",IF(OR(R[-1]C=""Skutečnost"",R[-1]C="""")=TRUE,RC[1],RC[1]+R[-1]C)))"
    ActiveCell.Offset(0, 6).Range("a1").Select
    
    'schvalovatelnost
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R152C61:R300C65,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C61:R300C65,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C61:R300C65,4,FALSE))),0,VLOOKUP(R2C3,Nstat3!R152C61:R300C65,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C61:R300C65,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C61:R300C65,4,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
 
     'PRODUKT F
       
    ActiveCell.Offset(28, -31).Range("a1").Select
    
    'nová produkce
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C36:R114C44,R90C8,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C36:R114C44,R90C8,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 10).Range("a1").Select
    
    'nová produkce-vlastní
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R90C18,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R90C18,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'revolvingy

    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'smlouvy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R117C36:R245C50,R90C28,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R117C36:R245C50,R90C28,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 6).Range("a1").Select

    'žádosti
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R90C33,FALSE)),0,VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R90C33,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]+R[-1]C"
    'ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",RC[1],R[-1]C),IF(RC[1]="""","""",IF(R[-1]C=""Skutečnost"",RC[1],RC[1]+R[-1]C)))"
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",IF(ISERROR(RC[1]),0,RC[1]),R[-1]C),IF(RC[1]="""","""",IF(OR(R[-1]C=""Skutečnost"",R[-1]C="""")=TRUE,RC[1],RC[1]+R[-1]C)))"
    ActiveCell.Offset(0, 6).Range("a1").Select
    
    'schvalovatelnost
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R152C67:R300C71,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C67:R300C71,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C67:R300C71,4,FALSE))),0,VLOOKUP(R2C3,Nstat3!R152C67:R300C71,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C67:R300C71,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C67:R300C71,4,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'PRODUKT G
    ActiveCell.Offset(28, -31).Range("a1").Select
    
    'nová produkce
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C36:R114C44,R118C8,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C36:R114C44,R118C8,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 10).Range("a1").Select
    
    'nová produkce-vlastní
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R118C18,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R2C50:R642C58,R118C18,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'revolvingy
    ActiveCell.Offset(0, 5).Range("a1").Select
    
    'smlouvy
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,'Nová produkce'!R117C36:R245C50,R118C28,FALSE)),0,VLOOKUP(R2C3,'Nová produkce'!R117C36:R245C50,R118C28,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, 6).Range("a1").Select

    'žádosti
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R118C33,FALSE)),0,VLOOKUP(R2C3,Nstat3!R4C48:R147C60,R118C33,FALSE))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    'ActiveCell.FormulaR1C1 = "=RC[1]+R[-1]C"
    'ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",RC[1],R[-1]C),IF(RC[1]="""","""",IF(R[-1]C=""Skutečnost"",RC[1],RC[1]+R[-1]C)))"
    ActiveCell.FormulaR1C1 = "=IF(ISERROR(RC[1]),IF(R[-1]C=""Skutečnost"",IF(ISERROR(RC[1]),0,RC[1]),R[-1]C),IF(RC[1]="""","""",IF(OR(R[-1]C=""Skutečnost"",R[-1]C="""")=TRUE,RC[1],RC[1]+R[-1]C)))"
    ActiveCell.Offset(0, 6).Range("a1").Select
    
    'schvalovatelnost
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R2C3,Nstat3!R152C73:R300C77,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C73:R300C77,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C73:R300C77,4,FALSE))),0,VLOOKUP(R2C3,Nstat3!R152C73:R300C77,3,FALSE)/(VLOOKUP(R2C3,Nstat3!R152C73:R300C77,3,FALSE)+VLOOKUP(R2C3,Nstat3!R152C73:R300C77,4,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("c3").Select
    ActiveWindow.SmallScroll up:=117
    ActiveWindow.SmallScroll ToLeft:=50
    
    sheets("nstat12").Visible = False
    sheets("revolvingy").Visible = False
    sheets("nová produkce").Visible = False
    sheets("nstat3").Visible = False
    sheets("NSTAT13").Visible = False
    sheets("TeamStat").Visible = False

    sheets("Pořadí").Select
    
    For Each wSheet In worksheets
    If wSheet.Name = "Pořadí" Then
    Else

    wSheet.Protect Password:="sim"
    End If

    Next
    sheets("Pořadí").Select
    Range("A1").Select


    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

  End Sub
