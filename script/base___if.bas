REM  *****  BASIC  *****

Option Explicit



Sub base___if___if_then

	If 4 > 0 Then
		
		MsgBox( "base___if" & Chr( 10 ) & Chr( 10 ) & "If 4 > 0 Then" )
		
	End If

End Sub



Sub base___if___if_then_else

	If 4 < 0 Then
		
		MsgBox( "base___if" & Chr( 10 ) & Chr( 10 ) & "If 4 < 0 Then" )
	
	Else
	
		MsgBox( "base___if" & Chr( 10 ) & Chr( 10 ) & "Else" )
		
	End If

End Sub



Sub base___if___if_then_else_if

	'
	' Else If : 안된다.
	' ElseIf : 된다.
	'
	If 4 < 0 Then
		
		MsgBox( "base___if" & Chr( 10 ) & Chr( 10 ) & "If 4 < 0 Then" )
	
	ElseIf 4 > 0 Then
	
		MsgBox( "base___if" & Chr( 10 ) & Chr( 10 ) & "ElseIf 4 > 0 Then" )
		
	End If

End Sub
