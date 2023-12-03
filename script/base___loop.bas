REM  *****  BASIC  *****

Option Explicit



Sub base___loop___for_next

	Dim result_string as String
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' For i = 0 to j 는 j <= 10 과 같다.
	'
	For i = 0 to j
		result_string = result_string & i & " "
	Next i
	
	MsgBox( result_string )

End Sub



Sub base___loop___for_next_step

	Dim result_string as String
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' For i = 0 to j 는 j <= 10 과 같다.
	'
	For i = 0 To j step 2
		result_string = result_string & i & " "
	Next i
	
	MsgBox( result_string )

End Sub
