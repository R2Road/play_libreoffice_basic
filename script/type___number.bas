REM  *****  LibreOffice VBA  *****

Sub type___number___prefix

	'
	' Hexadecimal
	'
	Dim hex as Byte : hex = &HFF ' 255
	
	MsgBox( hex )
	
	
	'
	' Octal
	'
	Dim oct as Byte : oct = &O07 ' 7
	
	MsgBox( oct )

End Sub
