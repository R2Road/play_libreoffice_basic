REM  *****  LibreOffice VBA  *****

Function base___formatiing

	Dim s as String
	
	s = s & "16진수 출력 : " & Hex( 16 ) '16진수 출력
	
	s = s & Chr( 10 )
	
	s = s & "8진수 출력 : " & Oct( 16 ) '8진수 출력
	
	MsgBox( s )

End Function
