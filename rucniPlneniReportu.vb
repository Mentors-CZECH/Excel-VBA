Option Explicit

Sub Vypln_Report_Rucne()
	Application.ScreenUpdating = False
	Call odkrytiMKTzdroje
	Call prepisNstat9
	Call odkrytiMKTzdroje
	Application.ScreenUpdating = True

End Sub
