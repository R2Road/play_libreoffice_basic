REM  *****  BASIC  *****

Sub base___class_module_test_01___class_module_generate_release

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
	' 프로퍼티에 접근
	'
	Print "Property : " & c.PI & " " & c.PL
	
	
	
	'
	' Release : Call - Class_Terminate
	'
	Set c = Nothing



End Sub
