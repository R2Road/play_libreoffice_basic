REM  *****  BASIC  *****

Option Explicit



Sub base___new___validation

	Dim a as Variant
	
	If IsEMpty( a ) Then
	
		MsgBox( "Empty : a" )
		
	EndIf
	
	
	Set a = New base___class_module_01
	
	
	If IsEmpty( a ) = false Then
	
		MsgBox( "Not Empty : a" )
		
	EndIf

End Sub
