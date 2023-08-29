REM  *****  LibreOffice VBA  *****

Sub type___array___declaration_1

	Dim a as Variant
	a = Array( 10, 20, 30 )
	
	MsgBox( a( 0 ) & a( 1 ) & a( 2 ) )

End Sub



Sub type___array___declaration_2

	Dim a( 1 to 3 ) as Integer
	a( 1 ) = 100
	a( 2 ) = 20
	a( 3 ) = 3
	
	Dim result as Integer
	For Each i In a
		result = result + i
	Next i
	
	MsgBox( result )

End Sub
