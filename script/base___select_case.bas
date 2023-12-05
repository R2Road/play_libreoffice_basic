REM  *****  BASIC  *****

Option Explicit



' REF : https://wiki.documentfoundation.org/Documentation/BASIC_Guide



Sub base___select_case

	Dim result_string as String	
	result_string = "select_case" & Chr( 10 ) & Chr( 10 )

	Dim value as Integer : value = 1


	Select Case value
		Case 1:
			result_string = result_string & " 1"
		Case 2:
			result_string = result_string & " 2"
	End Select
	
	
	
	MsgBox( result_string )

End Sub



Sub base___select_case___case_advance

	Dim result_string as String	
	result_string = "select_case" & Chr( 10 ) & Chr( 10 )

	Dim value as Integer : value = 10


	Select Case value
		Case 1 to 5
			result_string = result_string & " 1 to 5"
		Case > 100
			result_string = result_string & " > 100"
		Case 6, 7, 8
			result_string = result_string & " 6, 7, 8"
		Case 9, 10, 11, > 15, < 0
			result_string = result_string & " 9, 10, 11, > 15, < 0"
	End Select
	
	
	
	MsgBox( result_string )

End Sub



Sub base___select_case___case_else

	Dim result_string as String	
	result_string = "select_case" & Chr( 10 ) & Chr( 10 )

	Dim value as Integer : value = 6


	Select Case value
		Case 1 to 5
			result_string = result_string & " 1 to 5"
		Case Else
			result_string = result_string & " Else"
	End Select
	
	
	
	MsgBox( result_string )

End Sub
