sub copyStatToWorkbook()

 Dim sourceFileDialog As Object
    Dim USERNAME, sourceFilePath, thisWbName, msgVal As String

    Set sourceFileDialog = Application.FileDialog(msoFileDialogFilePicker)

	'Nastavení jména uživatele (systémového jména) pro vytváření například ceskty k desktop
    USERNAME = Environ("username")
	
	
	'Jméno sešitu, který volá dialog
    thisWbName = ActiveWorkbook.Name
    
	'Je možné přidávat další atributy, a měnit jejich parametry
    With sourceFileDialog
        .Title = "Zadejte cestu k souboru: xxx Název souboru xxx"
        .Filters.Clear
        .Filters.Add "Soubory MS Excel", "*.xl*"
        .ButtonName = "Načíst data"
        .AllowMultiSelect = False
        'Pokud uživatel klikne na storno, zobrzit promt s chybovou hlaskou
        If .Show <> -1 Then
            msgVal = MsgBox("Storno - chybova hlaska", vbCritical)
            Exit Sub
        End If
        'Přiřadí soubor (jeho cestu), který jste vybraly pomocí dialogu do proměné sourceFilePath
        sourceFilePath = .SelectedItems(1)
    End With


end sub