Function CZETOENG(source As Variant) As String

	Const cz As String = "áÁčČďĎéÉěĚíÍňŇóÓřŘšŠťŤúÚůŮýÝžŽ"
	Const en As String = "aAcCdDeEeEiInNoOrRsStTuUuUyYzZ"
	
	Dim TmpS As String
	Dim OutS As String
	
	Dim I As Integer
	
	OutS = ""
	If IsNull(source) Or source = "" Then
	 	CZETOENG = ""
	Else
	 	For I = 1 To Len(source)
	 	TmpS = Mid(source, I, 1)
	 	If InStr(1, cz, TmpS, vbBinaryCompare) > 0 Then TmpS = Mid(en, InStr(1, cz, TmpS, vbBinaryCompare), 1)
	 	OutS = OutS & TmpS
	 	Next I
 		CZETOENG = OutS
	End If

End Function
