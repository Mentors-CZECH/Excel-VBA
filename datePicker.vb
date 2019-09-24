Option Explicit

'_REQUIRES 
'MSCOMCT2.OCX - Pro ovládání calendar control (chyby v Excel 2013 32b/64b)
'
'

Private Sub datePicker_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    Cells(1, 1).Value = DateClicked
    Unload Me
End Sub

