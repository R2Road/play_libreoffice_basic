REM  *****  BASIC  *****

'
' Enum 을 사용하려면 이 "Option VBASupport 1" 이 필요하다
'
Option VBASupport 1



'
' Enum 정의
'
Enum eTest

	Hana
	Dul
	Sam

End Enum



Sub base___enum

	'
	' Basic
	'
	MsgBox( _
						"eTest.Hana : " &  eTest.Hana	_
		& Chr( 10 ) & 	"eTest.Dul : " &  eTest.Dul		_
		& Chr( 10 ) & 	"eTest.Sam : " &  eTest.Sam		_
	)
	
	'
	' For-Next
	'
	For i = eTest.Hana to eTest.Sam
		Print i
	Next

End Sub
