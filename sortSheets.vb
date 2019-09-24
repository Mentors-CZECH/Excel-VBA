Public thisWbName As String
Public saveFlag As Boolean

Sub SortWorksheets()
	Dim N As Integer
	Dim M As Integer
	Dim FirstWSToSort As Integer
	Dim LastWSToSort As Integer
	Dim SortDescending As Boolean

    SortDescending = False
    Application.ScreenUpdating = False

    If ActiveWindow.SelectedSheets.Count = 1 Then
			FirstWSToSort = 2
			LastWSToSort = Worksheets.Count
		Else
			With ActiveWindow.SelectedSheets
				For N = 2 To .Count
					If .Item(N - 1).Index <> .Item(N).Index - 1 Then
						MsgBox "Chyba, nelze setřídit"
						Exit Sub
					End If
				Next N
				FirstWSToSort = .Item(1).Index
				LastWSToSort = .Item(.Count).Index
			 End With
    End If
    
    For M = FirstWSToSort To LastWSToSort
        For N = M To LastWSToSort
            If SortDescending = True Then
                If UCase(Worksheets(N).Name) > UCase(Worksheets(M).Name) Then
                    Worksheets(N).Move Before:=Worksheets(M)
                End If
            Else
                If UCase(Worksheets(N).Name) < UCase(Worksheets(M).Name) Then
                   Worksheets(N).Move Before:=Worksheets(M)
                End If
            End If
         Next N
    Next M
    Application.ScreenUpdating = True
End Sub


