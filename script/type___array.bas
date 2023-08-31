REM  *****  LibreOffice VBA  *****

Sub type___array___declaration_1

	'
	' 기본 : 길이가 0인 배열( 이걸 어따 쓰지? )
	'
	Dim a() as Integer

End Sub



Sub type___array___declaration_2

	'
	' 기본 : 길이가 1인 배열
	'
	Dim a( 0 ) as Integer
	a( 0 ) = 100
	
	MsgBox( a( 0 ) )

End Sub



Sub type___array___declaration_3

	'
	' 기본
	'
	Dim a( 2 ) as Integer
	a( 0 ) = 100
	a( 1 ) = 20
	a( 2 ) = 3
	
	MsgBox( a( 0 ) + a( 1 ) + a( 2 ) )

End Sub



Sub type___array___declaration_4

	'
	' Index의 범위 조절 가능
	'
	Dim a( 1 to 3 ) as Integer
	a( 1 ) = 200
	a( 2 ) = 30
	a( 3 ) = 4
	
	MsgBox( a( 1 ) + a( 2 ) + a( 3 ) )

End Sub



Sub type___array___declaration_5

	'
	' 동적 배열
	'
	Dim a as Variant
	a = Array( 300, 20, 1 )
	
	MsgBox( a( 0 ) + a( 1 ) + a( 2 ) )

End Sub



Sub type___array___lbound_ubound

	'
	' LBound : 배열 시작 인덱스 반환
	' UBound : 배열 마지막 인덱스 반환
	'
	Dim a_0() as Integer
	MsgBox( LBound( a_0 ) & " " & UBound( a_0 ) )
	
	
	Dim a_1( 2 ) as Integer
	MsgBox( LBound( a_1 ) & " " & UBound( a_1 ) )
	
	
	Dim a_2( 1 to 3 ) as Integer
	MsgBox( LBound( a_2 ) & " " & UBound( a_2 ) )
	
	
	Dim a_3 as Variant
	a_3 = Array( 300, 20, 1 )
	MsgBox( LBound( a_3 ) & " " & UBound( a_3 ) )

End Sub



Sub type___array___size

	'
	' Size : UBound - LBound + 1
	'
	Dim a( 2 ) as Integer
	MsgBox( "Size : " & ( UBound( a ) - LBound( a ) + 1 ) )

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
