REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://www.debugpoint.com/libreoffice-workbook-worksheet-and-cell-processing-using-macro/
'

Sub document___title
	
	'
	' ThisComponent 는 현재 문서를 가리킨다.
	'
	Dim document as Object
	document = ThisComponent
	
	
	
	MsgBox( document.Title )

End Sub



Sub document___getActiveSheet
	
	'
	' ThisComponent 는 현재 문서를 가리킨다.
	'
	Dim document as Object
	document = ThisComponent
	
	
	'
	' getActiveSheet
	'
	Dim sheet as Object
	sheet = document.CurrentController.getActiveSheet()
	
	
	'
	' Sheet Size
	'
	MsgBox( sheet.Name & "  w : " & sheet.Columns.Count & "  h : " & sheet.Rows.Count )

End Sub
