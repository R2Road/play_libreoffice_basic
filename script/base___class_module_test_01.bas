REM  *****  BASIC  *****

Sub base___class_module_test_01___class_module_generate_release

	Dim c as Object
	
	'
	' Generate : Call - Class_Initialize
	'
	Set c = New base___class_module_01
	
	'
	' Release : Call - Class_Terminate
	'
	Set c = Nothing

End Sub
