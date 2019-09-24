Private Sub Workbook_Open()
    Dim array_example(100)
    Dim uidArray(100)
    Dim lastrowcolb As Integer
    Dim thisSheet As String

Application.EnableEvents = False
    List33.Activate
    thisSheet = ActiveSheet.Name
    
    lastrowcolb = Worksheets(thisSheet).Range("C" & Rows.Count).End(xlUp).Row
    'MsgBox lastrowcolb
    'Storing values in the array
    'ComboBox2
    
    For i = 28 To lastrowcolb
         array_example(i) = Range("C" & i)
    '     uidArray(i) = Range("B" & i)
         'MsgBox array_example(i)
    Next
    List32.Activate
    thisSheet = ActiveSheet.Name
    
    With List32.ComboBox1
        For i = 28 To lastrowcolb
        .AddItem array_example(i)
        Next
     End With
     
    'With List32.ComboBox2
    '   For i = 28 To lastrowcolb
    '   .AddItem uidArray(i)
    '   Next
    'End With

 
    Application.EnableEvents = True
 End Sub

