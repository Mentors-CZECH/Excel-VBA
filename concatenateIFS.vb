'1) Přidávat jednotlivé funkce postupně pod sebe
'2) Název funkce velkým písmem
'3) Při vytváření změn ve funkci napsat kdo provedl změnu a kde
'4) K funkcím přidávat popisky, a vysvětlení jak funkce funguje
'5) You DO NOT talk about FIGHT CLUB

'CONCATENATEIFS spojuje řetězce, pokud jsou splněny podmínky (minimálně 1) řetězce jsou odděleny znakem (Separator)
'=CONCATENATEIFS(RozsahProConcatenate; Criteria1; Operator1; Podmínka1; Criteria2; Operator2; Podmínka2; ..., Separator)
'=CONCATENATEIFS(A:A; B:B; ">"; 0; C:C; "="; "AM"; ...; ";")

Function CONCATENATEIFS(ConcatenateRange As Range, ParamArray Criteria() As Variant) As Variant
    Dim I As Long
    Dim c As Long
    Dim n As Long
    Dim f As Boolean
    Dim Separator As String
    Dim strResult As String
    On Error GoTo ErrHandler
    n = UBound(Criteria)
    If n < 3 Then
        'Too few arguments
        GoTo ErrHandler
    End If
    If n Mod 3 = 0 Then
        'Separator specified explicitly
        Separator = Criteria(n)
    Else
        'Use default separator
        Separator = ";"
    End If
    'Loop through the cells of the concatenate range
    For I = 1 To ConcatenateRange.Count
        'Start by assuming that we have a match
        f = True
        'Loop through the conditions
        For c = 0 To n - 1 Step 3
            'Does cell in criteria range match the condition?
            Select Case Criteria(c + 1)
                Case "<="
                    If Criteria(c).Cells(I).Value > Criteria(c + 2) Then
                        f = False
                        Exit For
                    End If
                Case "<"
                    If Criteria(c).Cells(I).Value >= Criteria(c + 2) Then
                        f = False
                        Exit For
                    End If
                Case ">="
                    If Criteria(c).Cells(I).Value < Criteria(c + 2) Then
                        f = False
                        Exit For
                    End If
                Case ">"
                    If Criteria(c).Cells(I).Value <= Criteria(c + 2) Then
                        f = False
                        Exit For
                    End If
                Case "<>"
                    If Criteria(c).Cells(I).Value = Criteria(c + 2) Then
                        f = False
                        Exit For
                    End If
                Case Else
                    If Criteria(c).Cells(I).Value <> Criteria(c + 2) Then
                        f = False
                        Exit For
                    End If
            End Select
        Next c
        'Were all criteria satisfied?
        If f Then
            'If so, add separator and value to result
            strResult = strResult & Separator & ConcatenateRange.Cells(I).Value
        End If
    Next I
    If strResult <> "" Then
        'Remove first separator
        strResult = Mid(strResult, Len(Separator) + 1)
    End If
    CONCATENATEIFS = strResult
    Exit Function
	
	ErrHandler:
		CONCATENATEIFS = CVErr(xlErrValue)
End Function

Sub RegisterUDF()
    Dim s As String
    s = "Spojí řetězce v zadaném rozsahu, podle vybraných kritérií" & vbLf _
    & "IF(CONCATENATEIFS(<expression>, <default>, <expression>)"

    Application.MacroOptions Macro:="CONCATENATEIFS", Description:=s, Category:=9
End Sub

Sub UnregisterUDF()
    Application.MacroOptions Macro:="CONCATENATEIFS", Description:=Empty, Category:=Empty
End Sub