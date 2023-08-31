REM  *****  LibreOffice VBA  *****

Option Explicit

Sub research___multibyte

	Dim document as Object
	document = ThisComponent
	
	Dim sheets as Object
	sheets = document.Sheets
	
	Dim sheet as Object
	sheet = sheets.getByName( "data_2" )	
	
	
	
	
	
	'
	'
	'
	Dim cell_0_1 as Object
	cell_0_1 = sheet.getCellByPosition( 0, 1 )
	
	
	Dim s as String : s =  cell_0_1.String
	MsgBox( s )
	
'	Dim i as Integer
'	For i = 1 To 3
	
		Dim c() as Byte : c = Mid( s, 1, 1 )

		MsgBox( c )
		MsgBox( c( 0 ) & " " & c( 1 ) )
		
		If ( c( 0 ) And &HF0 ) = &HE0 Then
			MsgBox( "3Byte Word" )
		ElseIf ( c( 0 ) And &HE0 ) = &HC0 Then
			MsgBox( "2Byte Word" )
		Else
			MsgBox( "1Byte Word" )
		EndIf
		
'	Next i

End Sub
