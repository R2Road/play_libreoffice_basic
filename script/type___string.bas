REM  *****  LibreOffice VBA  *****

Option VBASupport 1 'for StrConv
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
	view = view & "Right : " & Right( s, 2 )
	view = view & Chr( 10 )
	
	
	'
	' Mid : 지정한 위치를 기준으로 Index를 적용
	'
	view = view  & "Mid : "& Mid( s, 6, 1 )
	view = view & Chr( 10 )
	
	
	MsgBox( view )

End Sub



Sub type_string___Len

	Dim s as String : s = "test string"	
	
	'
	' Len : 문자열의 길이 반환
	'
	MsgBox( Len( s ) )

End Sub



Sub type_string___Instr

	Dim s as String : s = "test string"	
	
	'
	' Instr : 지정한 문자열을 찾아 위치를 반환
	'
	MsgBox( Instr( 1, s, "st" ) )

End Sub



'REF : https://help.libreoffice.org/latest/en-US/text/sbasic/shared/strconv.html
Sub type_string___StrConv

	Dim s as String : s = "가"
	
	
	'
	' '기본 인코딩은 Unicode. "가" 는 0:172, "각" 은 1:172
	'
	Dim x0() As Byte
	x0 = s
	MsgBox( x0( 0 ) & " " & x0( 1 ) )
	
	
	'
	'  EUC-KR 로 변환된다. "가" 는 176:161
	'
	Dim x1() As Byte
	x1 = StrConv( s, vbFromUnicode )
	MsgBox( x1( 0 ) & " " & x1( 1 ) )
	
	
	'
	' Unicode를 Unicode로 인코딩 시도하면 실패한다.
	'
	Dim x2() As Byte
	x2 = StrConv( s, vbUnicode)
	'MsgBox( x2( 0 ) & " " & x2( 0 ) )

End Sub

