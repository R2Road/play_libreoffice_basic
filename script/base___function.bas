REM  *****  BASIC  *****

Sub base___function

	'
	' 없는 함수도 호출 된다.
	'
	MsgBox( base___function___nothing() )
	
	'
	' 반환 값을 설정하지 않은 함수 호출 : 0이 반환된다.
	'
	MsgBox( base___function___return_x() )
	
	'
	' 반환 값을 설정한 함수 호출
	'
	MsgBox( base___function___return_o() )
	
	'
	' 인자가 있는 함수
	'
	MsgBox( base___function___has_argument( 9800, 76 ) )

End Sub



Function base___function___return_x as Integer
End Function



Function base___function___return_o as Integer

	'
	' 함수 이름? 과 같은 변수에 값을 할당 하면 반환 값으로 처리 된다.
	'
	base___function___return_o = 123

End Function



Function base___function___has_argument( a as Integer, b as Integer ) as Integer

	base___function___has_argument = ( a + b )

End Function
