REM  *****  BASIC  *****

Option Explicit



' REF : https://wiki.documentfoundation.org/Documentation/BASIC_Guide



Sub base___loop___for_next

	Dim result_string as String : result_string = "for_next" & Chr( 10 ) & Chr( 10 )
	
	
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



Sub base___loop___for_step_next

	Dim result_string as String : result_string = "for_step_next" & Chr( 10 ) & Chr( 10 )
	
	
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



Sub base___loop___do_while

	Dim result_string as String : result_string = "do_while" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	'
	'
	Do While i <= j
	
		result_string = result_string & i & " "
		
		i = i + 1
		
	Loop
	
	MsgBox( result_string )

End Sub



Sub base___loop___while_wend

	Dim result_string as String : result_string = "while_wend" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	'
	'
	While i <= j
	
		result_string = result_string & i & " "
		
		i = i + 1
		
	Wend
	
	MsgBox( result_string )

End Sub
