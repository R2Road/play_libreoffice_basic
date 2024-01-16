﻿REM  *****  BASIC  *****

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



Sub base___loop___for_next___step

	Dim result_string as String : result_string = "for_next___step" & Chr( 10 ) & Chr( 10 )
	
	
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



Sub base___loop___for_next___exit_for

	Dim result_string as String : result_string = "for_next___exit_for" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' For i = 0 to j 는 j <= 10 과 같다.
	'
	For i = 0 To j
	
		result_string = result_string & i & " "
		
		If i = 4 Then
			Exit For
		End If
		
	Next i
	
	MsgBox( result_string )

End Sub
