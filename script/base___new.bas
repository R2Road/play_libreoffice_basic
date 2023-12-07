REM  *****  BASIC  *****

Option Explicit



Sub base___new___empty

	Dim a as Variant
	
	
	If IsEMpty( a ) Then
	
		MsgBox( "1. Empty : a" )
		
	EndIf
	
	
	Set a = New base___class_module_01
	
	
	If IsEmpty( a ) = false Then
	
		MsgBox( "2. Not Empty : a" )
		
	EndIf
	
	
	Set a = Nothing
	
	
	'
	' Set a = Nothing 의 결과가 c++ 의 a = nullptr 과는 다른가 보다.
	'
	If IsEmpty( a ) Then
	
		MsgBox( "3. Empty : a" )
		
	EndIf
	
	
	a = Empty
	
	
	If IsEmpty( a ) Then
	
		MsgBox( "4. Empty : a" )
		
	EndIf

End Sub
