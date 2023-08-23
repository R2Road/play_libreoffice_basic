REM  *****  LibreOffice VBA  *****

'
' REF : https://help.libreoffice.org/latest/ko/text/sbasic/shared/03/sf_string.html
'

Sub script_forge___string___create

	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for String Utility
	
	Dim s as String : s = "abcd"
	
	
	
	'
	' String Service 생성
	'
	Dim sfs : sfs = CreateScriptService("String")
	
	
	
	'
	' String Service 객체를 직접 만들 필요 없다.
	' "ScriptForge" 를 로드 하면 String Service 객체가 "SF_String" 라는 이름으로 자동 생성된다.
	'
	s = SF_String.Capitalize( s )
	
	
	
	'
	' 다른 곳에 담아도 된다.
	'
	Dim svc : svc = SF_String
	
	
	MsgBox( s )
	
	
	
	'
	' SF_String 다른 함수에서 써보기
	'
	script_forge___string___create_2()

End Sub
Sub script_forge___string___create_2

	Dim s as String : s = "hijk lmn"
	
	
	'
	' "ScriptForge" 는 한 번 로드되면 다른 함수에서도 쓸 수 있다.
	' office 가 유지되는 동안에는 문제가 없는 것 같다.
	'
	s = SF_String.Capitalize( s )
	
	
	MsgBox( s )

End Sub
