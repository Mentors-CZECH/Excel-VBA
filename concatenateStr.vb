Sub Concatenate_Formula(bConcat As Boolean, bOptions As Boolean)
    Dim rSelected   As Range
    Dim c           As Range
    Dim sArgs       As String
    Dim bCol        As Boolean
    Dim bRow        As Boolean
    Dim sArgSep     As String
    Dim sSeparator  As String
    Dim rOutput     As Range
    Dim vbAnswer    As VbMsgBoxResult
    Dim lTrim       As Long
    Dim sTitle      As String


    Set rOutput = ActiveCell
    bCol = False
    bRow = False
    sSeparator = ""
    sTitle = IIf(bConcat, "CONCATENATE", "Ampersand")
    
    On Error Resume Next
    Set rSelected = Application.InputBox(Prompt:= _
                    "Select cells to create formula", _
                    Title:=sTitle & " Creator", Type:=8)
    On Error GoTo 0

    If Not rSelected Is Nothing Then

        sArgSep = IIf(bConcat, ",", "&")
        
        If bOptions Then
        
            vbAnswer = MsgBox("Columns Absolute? $A1", vbYesNo)
            bCol = IIf(vbAnswer = vbYes, True, False)
            
            vbAnswer = MsgBox("Rows Absolute? A$1", vbYesNo)
            bRow = IIf(vbAnswer = vbYes, True, False)
                
            sSeparator = Application.InputBox(Prompt:= _
                        "Type separator, leave blank if none.", _
                        Title:=sTitle & " separator", Type:=2)
        End If
        
        For Each c In rSelected.Cells
            sArgs = sArgs & c.address(bRow, bCol) & sArgSep
            If sSeparator <> "" Then
                sArgs = sArgs & Chr(34) & sSeparator & Chr(34) & sArgSep
            End If
        Next
        
        lTrim = IIf(sSeparator <> "", 4 + Len(sSeparator), 1)
        sArgs = Left(sArgs, Len(sArgs) - lTrim)

        If bConcat Then
            rOutput.Formula = "=CONCATENATE(" & sArgs & ")"
        Else
            rOutput.Formula = "=" & sArgs
        End If
    End If

End Sub