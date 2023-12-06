REM  *****  LibreOffice VBA  *****

Option Explicit

'
' REF : https://help.libreoffice.org/7.6/ko/text/sbasic/shared/03010102.html
'

Sub base___message_box

	Dim result as Integer
	
	result = MsgBox( "Test Message Box" )
	
	MsgBox( result )

End Sub



Sub base___message_box___ok

	Dim result as Integer

	result = MsgBox( "Test Message Box : 1", MB_OK )
	
	MsgBox( result )

End Sub



Sub base___message_box___ok_cancel

	Dim result as Integer

	result = MsgBox( "Test Message Box : 2", MB_OKCANCEL )
	
	MsgBox( result )

End Sub



Sub base___message_box___yes_no_cancel

	Dim result as Integer

	result = MsgBox( "Test Message Box : 3", MB_YESNOCANCEL )
	
	MsgBox( result )

End Sub



Sub base___message_box___title

	Dim result as Integer

	result = MsgBox( "Test Message Box : 4", MB_YESNO, "title" )
	
	MsgBox( result )

End Sub



Sub base___message_box___complex_option

	Dim result as Integer

	'
	' MB_DEFBUTTON2 옵션을 사용하면 Focus 가 2번 버튼에 할당된다.
	'
	result = MsgBox( "Test Message Box : X", MB_ABORTRETRYIGNORE + MB_DEFBUTTON2, "title" )
	
	MsgBox( result )

End Sub
