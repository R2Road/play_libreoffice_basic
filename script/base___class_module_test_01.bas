REM  *****  BASIC  *****

Option Explicit



Sub base___class_module_test_01___generate_release

	Dim c as Object
	
	
	
	'
	' Generate : Call - Class_Initialize
	'
	Set c = New base___class_module_01
	
	
	
	'
	' 멤버 변수에 접근
	' > Private 키워드가 작동을 안하는데?
	'
	Print "Member : " & c.i & " " & c.l
	
	
	
	'
	' 프로퍼티에 접근 : Get
	'
	Print "Property : Get : " & c.PI & " " & c.PL
	
	
	
	'
	' 프로퍼티에 접근 : Let
	'
	c.PI = 1000
	c.PL = 2000
	Print "Property : Let : " & c.PI & " " & c.PL
	
	
	
	'
	' Release : Call - Class_Terminate
	'
	Set c = Nothing



End Sub



Sub base___class_module_test_01___over_write

	Dim c as Object
	
	
	
	'
	' Generate : Call - Class_Initialize
	'
	Set c = New base___class_module_01
	
	
	
	'
	' Overwrite : 소멸자가 불리냐?
	'
	Set c = New base___class_module_01
	
	
	
	MsgBox( "소멸자 확인" )
	
	
	
	'
	' Release : Call - Class_Terminate
	'
	Set c = Nothing



End Sub
