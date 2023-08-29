REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/ko/text/sbasic/shared/03/sf_string.html?DbPAR=BASIC#bm_id151579602147056
'

Sub type_string___declaration_1

	Dim s as String
	
	s = "test string"
	
	MsgBox( s )

End Sub
Sub type_string___declaration_2

	Dim s as String : s = "test string"
	
	MsgBox( s )

End Sub
Sub type_string___declaration_3

	Dim s as String
	
	s = String( 10, "s" )
	
	MsgBox( s )

End Sub




Sub type_string___append

	Dim s as String : s = "test string"
	
	s = s & " append"
	
	MsgBox( s )

End Sub



Sub type_string___tab_linefeed

	Dim s as String : s = "test string"
	
	s = s & Chr( 9 ) & Chr( 9 ) & "append 1"
	
	s = s & Chr( 10 ) & "append 2"
	
	'
	' 안돼
	'
	s = s & Chr( 10 ) & "\t \\t \n \\n" & " ------ " & "Not Working"
	
	MsgBox( s )

End Sub




Sub type_string___StrReverse

	Dim s as String : s = "test string"
	
	StrReverse( s )
	
	MsgBox( s )

End Sub

