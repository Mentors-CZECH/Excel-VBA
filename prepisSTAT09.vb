Sub prepisNstat9()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' prepisNstat9 Makro - Klávesová zkratka: Ctrl+j
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Nábor

    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If

    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select

    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
                   
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    vložení proukce dle MKT zdroje z listu Stat5
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL8").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
    
    Sheets("Nstat3").Select
    Range("AL9").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
    
    Sheets("Nstat3").Select
    Range("AL10").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
    
    Sheets("Nstat3").Select
    Range("AL11").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
    
    Sheets("Nstat3").Select
    Range("AL12").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
         
    Sheets("Nstat3").Select
    Range("AL13").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
    
    Sheets("Nstat3").Select
    Range("AL14").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
         
      
    Sheets("Nstat3").Select
    Range("AL15").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    ActiveCell.Offset(0, -5).Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Uprava stat9
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat9").Select
    Range("M6:N6,M7:N7,M8:N8,M9:N9,M10:N10,M11:N11,M12:N12,M13:N13,M14:N14").Select
    Range("M14").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Columns("M:M").EntireColumn.AutoFit
    Range("V6:W6,V7:W7,V8:W8,V9:W9,V10:W10,V11:W11,V12:W12,V13:W13,V14:W14").Select
    Range("V14").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Columns("V:V").EntireColumn.AutoFit
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    Sheets("Nstat9").Select
    Range("M6").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select

    Range("M8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select

    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
       
    Sheets("Nstat9").Select
    Range("M12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   přepis natypovaných žádostí
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL1").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    Sheets("Nstat3").Select
    Range("AM1").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    Sheets("Nstat3").Select
    Range("AN1").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    Sheets("Nstat3").Select
    Range("AO1").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    Sheets("Nstat3").Select
    Range("AP1").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("AQ1").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
       Sheets("Nstat3").Select
    Range("AR1").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Přepis natypovaných žádostí sítě
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets("Nstat3").Select
    Range("AM8").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
          Sheets("Nstat3").Select
    Range("AM9").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

          Sheets("Nstat3").Select
    Range("AM10").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM11").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM12").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM13").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM14").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'přepis rozhodnuté žádosti žádostí
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL2").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("AM2").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("AN2").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("AO2").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("AP2").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("AQ2").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
    
    Sheets("Nstat3").Select
    Range("Ar2").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C=""Skutečnost"",RC[1],IF(RC[1]="""","""",R[-1]C+RC[1]))"
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Přepis rozhodnutých žádostí sítě
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Sheets("Nstat3").Select
    Range("An8").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Nstat3").Select
    Range("An9").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Sheets("Nstat3").Select
    Range("An10").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Nstat3").Select
    Range("An11").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Nstat3").Select
    Range("An12").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Nstat3").Select
    Range("An13").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("Nstat3").Select
    Range("An14").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'přepis schvalovatelnost
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL3").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Nstat3").Select
    Range("AM3").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Nstat3").Select
    Range("AN3").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Nstat3").Select
    Range("AO3").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Nstat3").Select
    Range("AP3").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Nstat3").Select
    Range("AQ3").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Sheets("Nstat3").Select
    Range("Ar3").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    Sheets("Nstat3").Select
    Range("At3").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
     ActiveCell.Offset(0, 43).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Přepis recidivy
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat9").Select
    Range("J19").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T19").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R32").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J20").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T20").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R33").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J21").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T21").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R34").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J22").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T22").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R35").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J23").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T23").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R36").Select
    Selection.Copy
    Sheets("Nstat9").Select
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J24").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T24").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R37").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J25").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T25").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R38").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J28").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T28").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R41").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    Sheets("Nstat9").Select
    Range("a1").Select
    
    Sheets("ČR CELKEM").Select

    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
    
End Sub

Sub prepisNstat9_novy_mesic()
'Sub prepisNstat9()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  MsgBox ("Aktualizace může trvat i několik sekund!")
        Styl = vbYes

    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Nábor
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
   Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
       Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
           Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
 ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
               Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
 ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
               Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
                   
                Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 56).Range("A1").Select
ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)),0,(VLOOKUP(R1C58,nábor!R5C17:R12C18,2,FALSE)))"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    ActiveCell.Offset(0, -56).Range("A1").Select
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    vložení proukce dle MKT zdroje z listu Stat5
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sheets("Nstat3").Select
Range("AL8").Select
Selection.Copy
Sheets("R01 Pardubice").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select

Sheets("Nstat3").Select
Range("AL9").Select
Selection.Copy
Sheets("R02 Praha").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select

Sheets("Nstat3").Select
Range("AL10").Select
Selection.Copy
Sheets("R03 Brno").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select

Sheets("Nstat3").Select
Range("AL11").Select
Selection.Copy
Sheets("R04 Ostrava").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select

Sheets("Nstat3").Select
Range("AL12").Select
Selection.Copy
Sheets("R05 Mladá Boleslav").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select
     
Sheets("Nstat3").Select
Range("AL13").Select
Selection.Copy
Sheets("R06 České Budějovice").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select

Sheets("Nstat3").Select
Range("AL14").Select
Selection.Copy
Sheets("R07").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select
     
  
Sheets("Nstat3").Select
Range("AL15").Select
Selection.Copy
Sheets("ČR CELKEM").Select
ActiveCell.Offset(0, 5).Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
ActiveCell.Offset(0, -5).Select
     
     
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Uprava stat9
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat9").Select
    Range("M6:N6,M7:N7,M8:N8,M9:N9,M10:N10,M11:N11,M12:N12,M13:N13,M14:N14").Select
    Range("M14").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Columns("M:M").EntireColumn.AutoFit
    Range("V6:W6,V7:W7,V8:W8,V9:W9,V10:W10,V11:W11,V12:W12,V13:W13,V14:W14").Select
    Range("V14").Activate
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Columns("V:V").EntireColumn.AutoFit
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    Sheets("Nstat9").Select
    Range("M6").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    'ActiveCell.Offset(1, -9).Range("A1").Select
    Range("M8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V9").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("M11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
       
    Sheets("Nstat9").Select
    Range("M12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 2).Range("a1").Select
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("V12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'přepis natypovaných žádostí
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL1").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AM1").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AN1").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AO1").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AP1").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    
    Sheets("Nstat3").Select
    Range("AQ1").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    
       Sheets("Nstat3").Select
    Range("AR1").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 11).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Přepis natypovaných žádostí sítě
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets("Nstat3").Select
    Range("AM8").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
          Sheets("Nstat3").Select
    Range("AM9").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

          Sheets("Nstat3").Select
    Range("AM10").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM11").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM12").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM13").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("AM14").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'přepis rozhodnuté žádosti žádostí
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL2").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AM2").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AN2").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AO2").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    Sheets("Nstat3").Select
    Range("AP2").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    
    Sheets("Nstat3").Select
    Range("AQ2").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
    
       Sheets("Nstat3").Select
    Range("Ar2").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 9).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -1).Range("A1").Select
    Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[1]"
 
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Přepis rozhodnutých žádostí sítě
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Sheets("Nstat3").Select
    Range("An8").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
          Sheets("Nstat3").Select
    Range("An9").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

          Sheets("Nstat3").Select
    Range("An10").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("An11").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("An12").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("An13").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
           Sheets("Nstat3").Select
    Range("An14").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'přepis schvalovatelnost
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Nstat3").Select
    Range("AL3").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Nstat3").Select
    Range("AM3").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Nstat3").Select
    Range("AN3").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Nstat3").Select
    Range("AO3").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Nstat3").Select
    Range("AP3").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    Sheets("Nstat3").Select
    Range("AQ3").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Sheets("Nstat3").Select
    Range("Ar3").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    Sheets("Nstat3").Select
    Range("At3").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
     ActiveCell.Offset(0, 43).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Přepis recidivy
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Sheets("Nstat9").Select
    Range("J19").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T19").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R32").Select
    Selection.Copy
    Sheets("R01 Pardubice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J20").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T20").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R33").Select
    Selection.Copy
    Sheets("R02 Praha").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J21").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T21").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R34").Select
    Selection.Copy
    Sheets("R03 Brno").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J22").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T22").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R35").Select
    Selection.Copy
    Sheets("R04 Ostrava").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J23").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T23").Select
    Selection.Copy
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R36").Select
    Selection.Copy
    Sheets("Nstat9").Select
    Sheets("R05 Mladá Boleslav").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J24").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T24").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R37").Select
    Selection.Copy
    Sheets("R06 České Budějovice").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    
    Sheets("Nstat9").Select
    Range("J25").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T25").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R38").Select
    Selection.Copy
    Sheets("R07").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select

    Sheets("Nstat9").Select
    Range("J28").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("T28").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Nstat9").Select
    Range("R41").Select
    Selection.Copy
    Sheets("ČR CELKEM").Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(1, -44).Range("A1").Select
    Sheets("Nstat9").Select
    Range("a1").Select
    Sheets("ČR CELKEM").Select
    
    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
    
End Sub
