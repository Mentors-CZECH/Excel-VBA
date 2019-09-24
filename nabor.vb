Sub nabor()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   nabor Makro
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 33).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)),0,(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(1, -33).Range("A1").Select
    
       Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 33).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)),0,(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(1, -33).Range("A1").Select
    
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 33).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)),0,(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(1, -33).Range("A1").Select
    
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 33).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)),0,(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(1, -33).Range("A1").Select
    
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 33).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)),0,(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(1, -33).Range("A1").Select
                   
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 33).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)),0,(VLOOKUP(R1C35,nábor!C17:C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(1, -33).Range("A1").Select
    
    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
    
End Sub
