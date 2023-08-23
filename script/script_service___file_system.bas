REM  *****  BASIC  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/he/text/sbasic/shared/03/sf_filesystem.html
'

Sub scriptservice_filesystem_readline

	GlobalScope.BasicLibraries.loadLibrary("Tools") ' for Tools
	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for FileSystem
	
	
	'
	'
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/../resources/" & "dummy_text.txt" )
	MsgBox( file_path )
	
	
	'
	' File System 개체 생성
	'
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	
	'
	' File Open
	'
	Dim pf As Variant
	Set pf = file_system.OpenTextFile(file_path, file_system.ForReading)
	
	
	'
	' Read Line
	'
	Dim file_string as String
	file_string = pf.ReadLine()
	MsgBox( file_string )
	
	
	file_string = ""
	
	
	'
	'
	'
	Do While Not pf.AtEndOfStream
		file_string = file_string & pf.ReadLine() & Chr( 10 )
	Loop
	MsgBox( file_string )
	
	
	'
	' Release
	'
	pf = pf.Dispose()

End Sub



Sub scriptservice_filesystem_open_close_readall

	GlobalScope.BasicLibraries.loadLibrary("Tools") ' for Tools
	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for FileSystem
	
	
	'
	'
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/../resources/" & "dummy_text.txt" )
	MsgBox( file_path )
	
	
	'
	' File System 개체 생성
	'
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	
	'
	' File Open
	'
	Dim pf As Variant
	Set pf = file_system.OpenTextFile(file_path, file_system.ForReading)
	
	
	'
	' Read All File
	'
	Dim file_string as String
	file_string = pf.ReadAll()
	MsgBox( file_string )
	
	
	'
	' File Close
	'
	pf.CloseFile()
	
	
	'
	' Release
	'
	pf = pf.Dispose()

End Sub



Sub scriptservice_filesystem_create_delete

	GlobalScope.BasicLibraries.loadLibrary("Tools") ' for Tools
	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for FileSystem
	
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/../resources/" & "dummy_text_4_scriptservice_filesystem_create_delete.txt" )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	
	'
	' File Create
	'
	Dim pf As Variant
	Set pf = file_system.CreateTextFile(file_path, Overwrite := true)
	
	
	
	MsgBox( "이 창을 닫지 말고 파일이 생성 됐는지 확인해봐" & Chr( 10 ) & Chr( 10 ) & file_path )
	
	
	
	pf.CloseFile()
	pf = pf.Dispose()
	
	
	'
	' 삭제
	'
	file_system.DeleteFile( file_path )

End Sub
