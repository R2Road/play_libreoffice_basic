REM  *****  LibreOffice VBA  *****

Sub type___array___declaration_1

	'
	' 기본
	'
	Dim a( 3 ) as Integer
	a( 1 ) = 100
	a( 2 ) = 20
	a( 3 ) = 3
	
	MsgBox( a( 1 ) + a( 2 ) + a( 3 ) )

End Sub



Sub type___array___declaration_2

	'
	' Index의 범위 조절 가능
	'
	Dim a( 0 to 2 ) as Integer
	a( 0 ) = 200
	a( 1 ) = 30
	a( 2 ) = 4
	
	MsgBox( a( 0 ) + a( 1 ) + a( 2 ) )

End Sub



Sub type___array___declaration_3

	'
	' 동적 배열
	'
	Dim a as Variant
	a = Array( 300, 20, 1 )
	
	MsgBox( a( 0 ) + a( 1 ) + a( 2 ) )

End Sub



Sub type___array___iteration_1

	Dim a as Variant : a = Array( 100, 20, 3 )
	
	'
	' For Each
	'
	Dim result as Integer
	For Each i In a
		result = result + i
	Next i
	
	MsgBox( result )

End Sub



Sub type___array___iteration_2

	Dim a as Variant : a = Array( 200, 30, 4 )
	
	'
	' For
	'
	Dim result as Integer
	For i = 0 To 2
		result = result + a( i )
	Next i
	
	MsgBox( result )

End Sub
