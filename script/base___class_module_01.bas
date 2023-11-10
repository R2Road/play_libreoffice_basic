REM  *****  BASIC  *****

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

	i = 10

    Print "Class Module : Initialize"
End Sub

'
' 소멸자
'
Private Sub Class_Terminate()
    Print "Class Module : Terminate"
End Sub ' Destructor
