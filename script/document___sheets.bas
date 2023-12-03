REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://help.libreoffice.org/latest/en-US/text/sbasic/guide/read_write_values.html
' REF : https://www.debugpoint.com/libreoffice-basic-macro-tutorial-index/
' REF : https://www.debugpoint.com/libreoffice-workbook-worksheet-and-cell-processing-using-macro/
'

Sub document___sheets___count_index_name
	
	Dim result_string as String
	
	
	'
	'
	'
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
	Dim cnt as Integer : cnt = sheets.Count - 1
	
	
	'
	' Index 로 Sheet 가져오기
	'
	Dim i as Integer
	For i = 0 to cnt
		result_string = result_string & i & " : " & sheets( i ).Name & Chr( 10 )
	Next i
	
	
	
	result_string = document.Title & Chr( 10 ) & Chr( 10 ) & result_string
	MsgBox( result_string )

End Sub



Sub document___sheets___row_and_column
	
	'
	' ThisComponent 는 현재 문서를 가리킨다.
	'
	Dim document as Object
	document = ThisComponent
	
	
	'
	' Get Sheet
	'
	Dim sheets as Object
	sheets = document.Sheets
	
	Dim sheet as Object
	sheet = sheets.getByName( "row_and_column" )
	
	
	'
	' Long 을 써야 한다.
	'
	Dim end_x as Long
	Dim end_y as Long
	end_x = sheet.Columns.Count
	end_y = sheet.Rows.Count
	
		
	MsgBox( "End : " & end_x & " : " & end_y )

End Sub



Sub document___sheets___active_row_and_column
	
	'
	' ThisComponent 는 현재 문서를 가리킨다.
	'
	Dim document as Object
	document = ThisComponent
	
	
	'
	' Get Sheet
	'
	Dim sheets as Object
	sheets = document.Sheets
	
	Dim sheet as Object
	sheet = sheets.getByName( "row_and_column" )
	
	
	
	'
	' Active X
	'
	Dim cur_x as Long
	Dim end_x as Long : end_x = sheet.Columns.Count
	
	For cur_x = 0 to end_x
		If sheet.getCellByPosition( cur_x, 0 ).String = "" Then
			cur_x = cur_x - 1
			Exit For
		EndIf
	Next cur_x
	
	
	
	'
	' Active Y
	'
	Dim cur_y as Long
	Dim end_y as Long : end_y = sheet.Rows.Count
	
	For cur_y = 0 to end_y
		If sheet.getCellByPosition( 0, cur_y ).String = "" Then
			cur_y = cur_y - 1
			Exit For
		EndIf
	Next cur_y
	
		
	MsgBox( _
						"Data : " & end_x & " : " & end_y _
		& Chr( 10 ) _
		& Chr( 10 ) & 	"Active X : " & "0" & " ~ " & cur_x _ 
		& Chr( 10 ) & 	"Active Y : " & "0" & " ~ " & cur_y _
	)

End Sub


	
Sub document___sheets___getcellbyposition
	
	Dim document as Object
	document = ThisComponent
	
	Dim sheets as Object
	sheets = document.Sheets
	
	Dim sheet as Object
	sheet = sheets.getByName( "data_1" )	
	
	
	
	
	
	'
	'
	'
	Dim cell_0_0 as Object
	cell_0_0 = sheet.getCellByPosition( 0, 0 )
	
	Dim cell_1_0 as Object
	cell_1_0 = sheet.getCellByPosition( 1, 0 )
	
	Dim cell_2_0 as Object
	cell_2_0 = sheet.getCellByPosition( 2, 0 )
	
	
	
	
	
	PrintCell( cell_0_0 )
	PrintCell( cell_1_0 )
	PrintCell( cell_2_0 )

End Sub



Sub PrintCell( cell as Object )

	Select Case cell.Type
		Case com.sun.star.table.CellContentType.VALUE
			MsgBox cell.Value
		Case com.sun.star.table.CellContentType.TEXT
			MsgBox cell.String
	End Select

End Sub



Sub document___sheets___selection

	'
	'
	'
    Dim selection_cell as Object
    selection_cell = ThisComponent.getCurrentSelection()
    
    
    
    MsgBox( "w : " & selection_cell.Columns.Count & Chr( 10 ) & "h : " &selection_cell.Rows.Count )
    
End Sub



Sub document___sheets___selected_cell

	'
	'
	'
    Dim selection_cell as Object
    selection_cell = ThisComponent.getCurrentSelection()
    
    
    '
    ' NOTE : selection_cell.getCellAddress 는 문서에서 1개 이상의 셀이 선택되어 있을때 터진다.
    '
    Dim CAC as Object
    CAC = ThisComponent.createInstance("com.sun.star.table.CellAddressConversion")
    CAC.Address = selection_cell.getCellAddress
    
    
    
    MsgBox( CAC.UserInterfaceRepresentation & Chr( 10 ) & CAC.PersistentRepresentation )
    
End Sub



Sub document___sheets___selected_range

	'
	'
	'
    Dim selection_cell as Object
    selection_cell = ThisComponent.getCurrentSelection()
    
    
    '
    '
    '
    Dim CRAC as Object
    CRAC = ThisComponent.createInstance("com.sun.star.table.CellRangeAddressConversion")
    CRAC.Address = selection_cell.getRangeAddress
    
    
    
    MsgBox( CRAC.UserInterfaceRepresentation & Chr( 10 ) & CRAC.PersistentRepresentation )
    
End Sub



'
' REF : https://wiki.documentfoundation.org/Macros/Calc/ba024
'
Sub document___sheets___sort

	'
	'
	'
    Dim sheet as Object
    sheet = ThisComponent.Sheets.getByName( "sorting_list" )
    
    
    '
    '
    '
    Dim range
    range = sheet.getCellRangeByName("A2:C8")
    
    
    '
    '
    '
    Dim sort_field(1) as new com.sun.star.util.SortField
    
	sort_field(0).Field = 0
    sort_field(0).SortAscending = TRUE
    sort_field(0).FieldType = com.sun.star.util.SortFieldType.ALPHANUMERIC 'com.sun.star.util.SortFieldTypeNUMERIC
    
    sort_field(1).Field = 1
    sort_field(1).SortAscending = TRUE
    sort_field(1).FieldType = com.sun.star.util.SortFieldType.ALPHANUMERIC 'com.sun.star.util.SortFieldTypeNUMERIC
    
    
    '
    '
    '
    Dim sort_description(0) as new com.sun.star.beans.PropertyValue
    
    sort_description(0).Name = "SortFields"
    sort_description(0).Value = sort_field()
    
    
    range.Sort( sort_description() )
    
    
    '
    '
    '
    MsgBox( "Sorting : Pivot Field 0 and 1" )
    
    
    '
	' Rollback
	'
	sort_field(0).Field = 2
	sort_field(1).Field = 2
	sort_description(0).Value = sort_field()
	range.Sort( sort_description() )

    
End Sub
















