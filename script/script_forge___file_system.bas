REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/he/text/sbasic/shared/03/sf_filesystem.html
'

Sub script_forge___file_system___open_close

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
	' File Close
	'
	pf.CloseFile()
	
	
	'
	' File Release
	'
	pf = pf.Dispose()

End Sub



Sub script_forge___file_system___create_delete

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
