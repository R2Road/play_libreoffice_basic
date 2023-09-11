REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/ko/text/sbasic/shared/03/sf_dictionary.html
'

Sub script_forge___dictionary___declaration

	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
	
	'
	' Declaration
	'
	Dim dic as Variant	
	dic = CreateScriptService("Dictionary")

End Sub



Sub script_forge___dictionary___add

	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
	
	Dim dic as Variant	
	dic = CreateScriptService("Dictionary")
	
	'
	' Add
	'
	dic.Add( "a", 100 )
	dic.Add( "b", 200 )
	dic.Add( "c", 300 )
	
	MsgBox( dic.Keys( 0 ) & dic.Item( "a" ) )

End Sub



Sub script_forge___dictionary___iteration

	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
	
	Dim dic as Variant	
	dic = CreateScriptService("Dictionary")
	
	dic.Add( "a", 100 )
	dic.Add( "b", 200 )
	dic.Add( "c", 300 )
	
	'
	' Iteration
	'
	Dim keys as Variant, i as Variant
	keys = dic.Keys()
	For Each i in keys
		MsgBox( i & " : " & dic.Item( i ) )
	Next i

End Sub
