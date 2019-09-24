Sub Obdélník1_Kliknutí()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Obdélník1_Kliknutí Makro - Produkty AEFG
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    
    Range("Av6:bc6").Select
    Selection.Copy
    Range("Al1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av8:bc8").Select
    Selection.Copy
    Range("Al4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
            
    Range("Av29:bc29").Select
    Selection.Copy
    Range("AL2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("Av31:bc31").Select
    Selection.Copy
    Range("AL5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av50:bc50").Select
    Selection.Copy
    Range("AL3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Sheets("Stat5").Select
    Range("AM7:AN13").Select
    Selection.Copy
    Sheets("Nstat3").Select
    Range("AK8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("AV8:BB8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
      
    Range("AV31:BB31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AN8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

End Sub

Sub Obdélník2_Kliknutí()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Obdélník2_Kliknutí Makro - Produkt A
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    
    Range("Av10:bc10").Select
    Selection.Copy
    Range("Al1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av12:bc12").Select
    Selection.Copy
    Range("Al4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av33:bc33").Select
    Selection.Copy
    Range("AL2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av35:bc35").Select
    Selection.Copy
    Range("AL5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av54:bc54").Select
    Selection.Copy
    Range("AL3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Sheets("Stat5").Select
    Range("AM23:AN29").Select
    Selection.Copy
    Sheets("Nstat3").Select
    Range("AK8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("AV12:BB12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

    Range("AV35:BB35").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AN8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
        
    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

End Sub

Sub Obdélník3_Kliknutí()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Obdélník2_Kliknutí Makro - Produkt E
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    
    Range("Av14:bc14").Select
    Selection.Copy
    Range("Al1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    
    Range("Av16:bc16").Select
    Selection.Copy
    Range("Al4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    
    Range("Av37:bc37").Select
    Selection.Copy
    Range("AL2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

    Range("Av39:bc39").Select
    Selection.Copy
    Range("AL5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("Av58:bc58").Select
    Selection.Copy
    Range("AL3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Sheets("Stat5").Select
    Range("Am39:An45").Select
    Selection.Copy
    Sheets("Nstat3").Select
    Range("AK8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
  
    Range("AV16:BB16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
      
    Range("AV39:BB39").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AN8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
        
    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

End Sub

Sub Obdélník5_Kliknutí()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Obdélník2_Kliknutí Makro - Produkt G
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    
    Range("Av22:bc22").Select
    Selection.Copy
    Range("Al1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    
    Range("Av24:bc24").Select
    Selection.Copy
    Range("Al4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    
    Range("Av45:bc45").Select
    Selection.Copy
    Range("AL2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("Av47:bc47").Select
    Selection.Copy
    Range("AL5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("Av66:bc66").Select
    Selection.Copy
    Range("AL3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Sheets("Stat5").Select
    Range("Am72:An78").Select
    Selection.Copy
    Sheets("Nstat3").Select
    Range("AK8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("AV24:BB24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
      
    Range("AV47:BB47").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AN8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
        
    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If

End Sub

Sub Obdélník4_Kliknutí()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Obdélník2_Kliknutí Makro - Produkt F
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Application.ScreenUpdating = False Then
        appFlag = False
        Else
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        appFlag = True
    End If
    
    Range("Av18:bc18").Select
    Selection.Copy
    Range("Al1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    
    Range("Av20:bc20").Select
    Selection.Copy
    Range("Al4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    
    Range("Av41:bc41").Select
    Selection.Copy
    Range("AL2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("Av43:bc43").Select
    Selection.Copy
    Range("AL5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Range("Av62:bc62").Select
    Selection.Copy
    Range("AL3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      
    Sheets("Stat5").Select
    Range("Am55:An61").Select
    Selection.Copy
    Sheets("Nstat3").Select
    Range("AK8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
       
    Range("AV20:BB20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
      
    Range("AV43:BB43").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AN8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
        
    If appFlag = True Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub
