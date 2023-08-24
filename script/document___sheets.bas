REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://www.debugpoint.com/libreoffice-basic-macro-tutorial-index/
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
