REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://www.debugpoint.com/libreoffice-workbook-worksheet-and-cell-processing-using-macro/
'

Sub document___sheets___count_index_name
	
	Dim document as Object
	document = ThisComponent
	
	
	'
	' Sheets
	'
	Dim sheets as Object
	sheets = document.Sheets
	
	
	'
	' Sheet Count
	'
	Dim i, cnt as Integer : cnt = sheets.Count - 1
	Dim s as String
	
	
	'
	' Index 로 Sheet 가져오기
	'
	s = s & document.Title & Chr( 10 ) & Chr( 10 )
	For i = 0 to cnt
		s = s & i & " : " & sheets( i ).Name & Chr( 10 )
	Next i
	
	
	
	MsgBox( s )

End Sub


	
Sub E
	
	Dim document as Object
	document = ThisComponent
	
	Dim sheets as Object
	sheets = document.Sheets
	
	Dim sheet_count as Integer
	sheet_count = sheets.Count
	
	Dim sheet as Object
	sheet = sheets.getByName( "data_1" )
	
	Dim cell_0_0 as Object
	cell_0_0 = sheet.getCellByPosition( 0, 0 )
	
	Dim cell_1_0 as Object
	cell_1_0 = sheet.getCellByPosition( 1, 0 )
	
	MsgBox "Sheet Count : " & sheet_count
	
	PrintCell( cell_0_0 )
	PrintCell( cell_1_0 )

End Sub



Sub PrintCell( cell as Object )

	Select Case cell.Type
		Case com.sun.star.table.CellContentType.VALUE
			MsgBox cell.Value
		Case com.sun.star.table.CellContentType.TEXT
			MsgBox cell.String
	End Select

End Sub
