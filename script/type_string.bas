REM  *****  BASIC  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/ko/text/sbasic/shared/03/sf_string.html?DbPAR=BASIC#bm_id151579602147056
'

Sub main

	Dim s as String : s = "test string"
	
	s = s & " 2"
	
	MsgBox( s )

End Sub
