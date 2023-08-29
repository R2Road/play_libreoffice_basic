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




Sub type_string___Left_Right_Mid

	Dim view as String
	
	Dim s as String : s = "test string"
	
	
	'
	' Left : index 가 1부터 시작
	'
	view = view & "Left : " & Left( s, 0 )
	view = view & Chr( 10 )
	
	
	'
	' Left : index 가 1부터 시작
	'
	view = view & "Left : " & Left( s, 1 )
	view = view & Chr( 10 )
	
	
	'
	' Right
	'
	view = view & "Right : " & Right( s, 1 )
	view = view & Chr( 10 )
	
	
	'
	' Mid : 지정한 위치를 기준으로 Index를 적용
	'
	view = view  & "Mid : "& Mid( s, 6, 1 )
	view = view & Chr( 10 )
	
	
	MsgBox( view )

End Sub

