REM  *****  LibreOffice VBA  *****

Sub type___array___declaration_1

	'
	' 기본
	'
	Dim a( 3 ) as Integer
	a( 1 ) = 100
	a( 2 ) = 20
	a( 3 ) = 3
	
	Dim result as Integer
	For Each i In a
		result = result + i
	Next i
	
	MsgBox( result )

End Sub



Sub type___array___declaration_2

	'
	' Index의 범위 조절 가능
	'
	Dim a( 0 to 2 ) as Integer
	a( 0 ) = 100
	a( 1 ) = 20
	a( 2 ) = 3
	
	Dim result as Integer
	For Each i In a
		result = result + i
	Next i
	
	MsgBox( result )

End Sub



Sub type___array___declaration_3

	'
	' 동적 배열
	'
	Dim a as Variant
	a = Array( 10, 20, 30 )
	
	MsgBox( a( 0 ) & a( 1 ) & a( 2 ) )

End Sub
