REM  *****  BASIC  *****

'
' Class Module 을 사용하려면 "Option Compatible", "Option ClassModule" 이 필요하다
'
Option Compatible
Option ClassModule

'
' 멤버 변수
'
Private i as Integer
Public l as Long

'
' 생성자
'
Private Sub Class_Initialize()

	i = 1
	l = 2

    Print "Class Module : Initialize"
End Sub

'
' 소멸자
'
Private Sub Class_Terminate()
    Print "Class Module : Terminate"
End Sub ' Destructor



'
' 프로퍼티 : Get
'
Public Property Get PI() as Integer

	PI = i + 10

End Property

Public Property Get PL() as Integer

	PL = l + 20

End Property



'
' 프로퍼티 : Let
'
Public Property Let PI( arg as Integer )

	i = arg + 10

End Property

Public Property Let PL( arg as Integer )

	l = arg + 20

End Property
