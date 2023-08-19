REM  *****  BASIC  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/ko/text/sbasic/shared/03/sf_dictionary.html
'

Sub main

	GlobalScope.BasicLibraries.loadLibrary("ScriptForge")
	
	Dim dic as Variant
	
	dic = CreateScriptService("Dictionary")
	
	dic.Add( "a", 100 )
	dic.Add( "b", 200 )
	dic.Add( "c", 300 )
	
	Dim keys as Variant, i as Variant
	keys = dic.Keys()
	For Each i in keys
		MsgBox( i & " : " & dic.Item( i ) )
	Next i

End Sub
