REM  *****  BASIC  *****

Sub main

GlobalScope.BasicLibraries.loadLibrary("Tools")



'
' 파일 경로 생성
'
Dim file_path as String
file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/" & "test.txt" )



'
' 파일 포인터?
'
Dim file_line_count as Integer
file_line_count = FreeFile



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

End Sub
