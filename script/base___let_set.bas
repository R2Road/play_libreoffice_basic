REM  *****  BASIC  *****

Function base___let

	Dim c as Integer : c = 10
	Dim d as Integer : d = 20
	
	'
	' 대입과 같다.
	'
	let c = d
	
	MsgBox( c )

End Function



Function base___set

	Dim c as New Collection
	
	Dim d as New Collection
	d.Add "d"
	d.Add "e"
	d.Add "f"
	
	MsgBox( c.Count )
	
	'
	' Set은 해당 객체의 포인터를 저장하게 한다.
	'
	Set c = d
	
	MsgBox( c.Count )
	
	d.Add "g"
	
	MsgBox( c.Count )

End Function
