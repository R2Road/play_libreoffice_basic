REM  *****  BASIC  *****

Option Explicit



'REF : https://renenyffenegger.ch/notes/development/languages/VBA/language/null-and-nothing-etc

Function CheckEmpty4Variant( i as Integer, a as Variant )
	
	If IsEmpty( a ) Then
	
		MsgBox( i & ". Empty : a" )
	
	Else
		
		MsgBox( i & ". Not Empty : a" )
		
	EndIf
	
End Function

Sub base___new___empty

	Dim a as Variant	
	
	CheckEmpty4Variant( 1, a )
	
	
	
	Set a = New base___class_module_01	
	
	CheckEmpty4Variant( 2, a )
	
	
	
	Set a = Nothing	
	
	'
	' Set a = Nothing 의 결과가 c++ 의 a = nullptr 과는 다른가 보다.
	'
	CheckEmpty4Variant( 3, a )
	
	
	
	a = Empty	
	
	CheckEmpty4Variant( 4, a )

End Sub
