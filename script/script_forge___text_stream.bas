REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://help.libreoffice.org/7.6/ko/text/sbasic/shared/03/sf_textstream.html
'

Sub script_forge___text_stream___writeline

	GlobalScope.BasicLibraries.loadLibrary("Tools") ' for Tools
	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for FileSystem
	
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/../resources/" & "script_forge___text_stream___writeline.txt" )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	Dim pf As Variant
	Set pf = file_system.CreateTextFile(file_path, Overwrite := true)
	
	
	
	
	
	'
	' WriteLine
	'
	pf.WriteLine( "[" )
	pf.WriteLine( Chr( 9 ) & "{ ""idx"" : 1234 }" )
	pf.WriteLine( "]" )
	
	
	
	
	
	pf.CloseFile()
	pf = pf.Dispose()
	
	MsgBox( "파일 삭제 대기" & Chr( 10 ) & Chr( 10 ) & file_path )
	
	file_system.DeleteFile( file_path )

End Sub
