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

Sub base___enum___declaration_1

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



'
' Enum 정의
'
Enum eTest2

	Hana = 3
	Dul = 2
	Sam = 1

End Enum

Sub base___enum___declaration_2

	'
	' Basic
	'
	MsgBox( _
						"eTest2.Hana : " &  eTest2.Hana	_
		& Chr( 10 ) & 	"eTest2.Dul : " &  eTest2.Dul		_
		& Chr( 10 ) & 	"eTest2.Sam : " &  eTest2.Sam		_
	)
	
	'
	' For-Next
	'
	For i = eTest2.Hana to eTest2.Sam
		Print i
	Next

End Sub
