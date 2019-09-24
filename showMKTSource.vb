Sub odkrytiMKTzdroje()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Odkryje/Skryje MKT zdroj - Klávesová zkratka: Ctrl+t
' .Visible = 2 (xlVeryhidden - uživatel nemůže list zobrazit)
' .Visible = -1 (xlVisible - List je zobrazen)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim worksheetsArray, backendArray As Variant
    Dim Item As Variant
    
    'Listy regionů, které chcete odkrývat/skrývat
    worksheetsArray = Array("R01 Pardubice", "R02 Praha", "R03 Brno", "R04 Ostrava", "R05 Mladá Boleslav", "R06 České Budějovice", "R07", "ČR CELKEM")
    'Listy statistiky a ostatní listy, které uživatel v normálním zobrazení nevydí
    backendArray = Array("Nstat3", "Stat5", "nábor", "plnění ročního plánu", "RizeneVplaceni_Smlouvy", "Stat2", "Plány", "ČR CELKEM vč nevyplacené NH")

    With Application
        .ScreenUpdating = False
    End With

    If Sheets("Nstat3").Visible = True Then 'Protect
        'Projde všechny listy a zahesluje je. Nastaví outline a zarovná k A1
        For Each Item In worksheetsArray

            With Sheets(Item)
                .Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
                .Protect Password:="sim"
            End With
            ActiveWindow.ScrollColumn = 1
            ActiveWindow.ScrollRow = 1
            Range("A1").Select
        Next Item
        'Skryje listy statistik jako xlVeryHidden (uživatel je nemůže zobrazit)
        For Each Item In backendArray
            Sheets(Item).Visible = 2
        Next Item
   Else 'Unprotect
        'Odhesluje všechny listy regionů
        For Each Item In worksheetsArray
        Sheets(Item).Select
            With Sheets(Item)
                ActiveSheet.Unprotect Password:="sim"
                ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
            End With
        Next Item
        'Odkryhy listy statistik
        For Each Item In backendArray
            If Item = "plnění ročního plánu" Or Item = "Plány" Or Item = "ČR CELKEM vč nevyplacené NH" Or Item = "RizeneVplaceni_Smlouvy" Then
            Else
                Sheets(Item).Visible = -1
            End If
        Next Item
    End If
    
    With Application
        .ScreenUpdating = True
    End With

End Sub
