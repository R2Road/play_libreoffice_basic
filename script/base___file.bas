REM  *****  LibreOffice VBA  *****

Sub base___file___open_write_print_close_kill

	GlobalScope.BasicLibraries.loadLibrary("Tools")
	
	
	
	'
	' 파일 경로 생성
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/../resources/" & "base___file.txt" )
	
	
	
	'
	' 파일 포인터?
	'
	Dim file_line_count as Integer
	file_line_count = FreeFile '초기화
	
	
	
	'
	' Open, Write, Print, Close
	'
	Open file_path For Output as #file_line_count
		
		Write #file_line_count, "a", 200
		Write #file_line_count, "b"
		Write #file_line_count, "c"
		
		Print #file_line_count, "a", 200
		Print #file_line_count, "b"
		Print #file_line_count, "c"
	
	Close #file_line_count
	
	
	
	MsgBox( "이 창을 닫지 말고 파일이 생성 됐는지 확인해봐" & Chr( 10 ) & Chr( 10 ) & file_path )
	
	
	
	'
	' 파일 삭제
	'
	Kill file_path

End Sub

