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
Sub base___loop___for_next_exit

	Dim result_string as String : result_string = "for_next_exit" & Chr( 10 ) & Chr( 10 )
	
	
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



Sub base___loop___for_each___with_variant

	' REF : https://wiki.documentfoundation.org/Documentation/BASIC_Guide - For Each 항목에서 가져왔다.
	
	Dim result_string as String : result_string = "for_each ??" & Chr( 10 ) & Chr( 10 )
	
	
	Const a1 = 1
	Const a2 = 2
	Const a3 = 3
	
	'
	' Variant
	'
	Dim a( a1, a2, a3 )
	
	
	'
	' 24 Loop
	'
	' 0 ~ 1 : 2
	' 0 ~ 2 : 3
	' 0 ~ 3 : 4
	'
	' 2 * 3 * 4 = 24
	' 맞냐? 이거?
	'
	Dim i as Integer : i = 0
	Dim e
	For Each e in a()
	
		i = i + 1
		
	Next e
	
	
	result_string = result_string & " Loop : " & i
	MsgBox( result_string )

End Sub



Sub base___loop___do_while_loop

	Dim result_string as String : result_string = "do_while_loop" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' 조건을 만족 하면 돌아간다.
	'
	Do While i <= j
	
		result_string = result_string & i & " "
		
		i = i + 1
		
	Loop
	
	MsgBox( result_string )

End Sub
Sub base___loop___do_loop_while

	Dim result_string as String : result_string = "do_loop_while" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' 조건을 만족 하면 돌아간다.
	'
	Do
	
		result_string = result_string & i & " "
		
		i = i + 1
		
	Loop While i <= j
	
	MsgBox( result_string )

End Sub



Sub base___loop___do_until_loop

	Dim result_string as String : result_string = "do_until_loop" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' 조건을 만족 할 때 까지 돌아간다.
	'
	Do Until i > j
	
		result_string = result_string & i & " "
		
		i = i + 1
		
	Loop
	
	MsgBox( result_string )

End Sub
Sub base___loop___do_loop_until

	Dim result_string as String : result_string = "do_loop_until" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' 조건을 만족 할 때 까지 돌아간다.
	'
	Do
	
		result_string = result_string & i & " "
		
		i = i + 1
		
	Loop Until i > j
	
	MsgBox( result_string )

End Sub



Sub base___loop___do_loop

	Dim result_string as String : result_string = "do_loop" & Chr( 10 ) & Chr( 10 )
	
	
	Dim i as Integer
	Dim j as Integer : j = 10
	
	
	'
	' while( 1 ) 과 비슷하다.
	'
	Do
	
		result_string = result_string & i & " "
		
		i = i + 1
		
		If i = j Then
			Exit Do
		end If
		
	Loop
	
	MsgBox( result_string )

End Sub



Sub base___loop___for_each

	Dim result_string as String : result_string = "for_each" & Chr( 10 ) & Chr( 10 )
	
	
	'
	' 기본
	'
	Dim a( 2 ) as Integer
	a( 0 ) = 100
	a( 1 ) = 20
	a( 2 ) = 3
	
	
	Dim i as Integer
	For Each i in a
	
		result_string = result_string & i & " "
		
	Next i
	
	
	MsgBox( result_string )

End Sub
