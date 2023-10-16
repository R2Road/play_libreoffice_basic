REM  *****  BASIC  *****

Function base___collection___declaration

	Dim c as New Collection

End Function



Function base___collection___add

	Dim c as New Collection
	
	Dim result as String
	
	'
	'
	'
	c.Add "A"
	c.Add "B"
	
	result = "c.Count : " & c.Count
	
	For Each i In c
		result = result & Chr( 10 ) & i
	Next i
	
	MsgBox( result )

End Function
