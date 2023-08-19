REM  *****  BASIC  *****

Option Explicit

'
' REF : https://www.debugpoint.com/libreoffice-workbook-worksheet-and-cell-processing-using-macro/
'
	
Sub main
	
	Dim doc as Object
	doc = ThisComponent
	
	Dim sheets as Object
	sheets = doc.Sheets
	
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
