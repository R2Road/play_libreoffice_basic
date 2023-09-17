REM  *****  LibreOffice VBA  *****

Option VBASupport 1 'for Len, StrConv
Option Explicit

'
'REF : https://leesumin.tistory.com/78
'

'
' 한글 결합식
'
' (초성 인덱스 * 21 + 중성 인덱스) * 28 + 종성 인덱스 + 0xAC00
' > 21은 중성의 총 수
' > 28은 종성의 총 수
' 
' 즉 한글 결합식은 최소한의 숫자를 사용하면서 인덱스를 하나로 합치기 위해
' 먼저 계산식에 오른 Index에 다음 Index 의 총 수를 곱해서 수를 밀어 올리면서 연산을 반복해서 완성 하는 것이다.
' 
' 초성 배열 19개 : "ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ", "ㅅ", "ㅆ", "ㅇ" , "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"
' 중성 배열 21개 : "ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ", "ㅗ", "ㅘ", "ㅙ", "ㅚ", "ㅛ", "ㅜ", "ㅝ", "ㅞ", "ㅟ", "ㅠ", "ㅡ", "ㅢ", "ㅣ"
' 종성 배열 28개 : "", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ", "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"
'

'
' REF : https://velog.io/@limdumb/%EC%9C%A0%EB%8B%88%EC%BD%94%EB%93%9C-%ED%95%9C%EA%B8%80-%EB%B2%94%EC%9C%84
' 유니코드 한글 범위 : AC00 ~ D7FF
'
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
	
	Dim b() as Byte
	b = Mid( s, 1, 1 )
	MsgBox( b )
	
	Dim result as Boolean : result = False
	If IsMultiByte( b ) Then
		result = IsKorean( b )
	EndIf
	MsgBox( result )

End Sub



Function research___multibyte___check_multibyte

	Dim s as String : s = "가a1"
	Dim b() as Byte
	
	Dim i as Integer
	For i = 1 To 3
		b = Mid( s, i, 1 )
		MsgBox( IsMultiByte( b ) )
	Next i
	
End Function
Function IsMultiByte( b() as Byte )

	'
	' 유니코드 범위 체크 : English + Latin : https://en.wikipedia.org/wiki/List_of_Unicode_characters
	'
	IsMultiByte = ( b( 1 ) <> 0 )

End Function




Function research___multibyte___check_korean

	Dim s as String : s = "가a1"
	Dim b() as Byte
	
	
	Dim i as Integer
	For i = 1 To 3
		b = Mid( s, i, 1 )
		MsgBox( IsKorean( b ) )
	Next i
	
End Function
Function IsKorean( b() as Byte )

	'
	' 유니코드 범위 체크 : 한글 : AC00 ~ D7FF
	'
	IsKorean = ( b( 1 ) >= &HAC And b( 1 ) < &HD8 )

End Function



Function research___multibyte___extract_initial_consonant '초성 : initial_consonant

	Dim s as String : s = "가a1"
	Dim b() as Byte
	
	Dim i as Integer
	For i = 1 To 3
		b = Mid( s, i, 1 )
		MsgBox( "0 : " & b( 0 ) & Chr( 10 ) & "1 : " & b( 1 ) & Chr( 10 ) & Extract_InitialConsonant( b ) )
	Next i
	
End Function
Function Extract_InitialConsonant( b() as Byte )

	'
	' byte array를 하나의 수로 만든다.
	'
	
	Dim i as Long 'Integer : 16bit, Long : 32bit
	
	'
	' b( 1 )
	'
	i = b( 1 )
	i = i * 256 ' 256 : 2의 8승 : 왼쪽 shift 8
	
	'
	' b( 0 )
	'
	i = i + b( 0 )
	
	Extract_InitialConsonant = i
	
End Function



Function research___multibyte___extract_vowel '모음 : vowel

	Dim s as String : s = "가a1"
	Dim b() as Byte
	
	Dim i as Integer
	For i = 1 To 3
		b = Mid( s, i, 1 )
		MsgBox( "0 : " & b( 0 ) & Chr( 10 ) & "1 : " & b( 1 ) & Chr( 10 ) & Extract_Vowel( b ) )
	Next i
	
End Function
Function Extract_Vowel( b() as Byte )
	
End Function



Function research___multibyte___extract_final_consonant '종성 : final_consonant

	Dim s as String : s = "가a1"
	Dim b() as Byte
	
	Dim i as Integer
	For i = 1 To 3
		b = Mid( s, i, 1 )
		MsgBox( "0 : " & b( 0 ) & Chr( 10 ) & "1 : " & b( 1 ) & Chr( 10 ) & Extract_FinalConsonant( b ) )
	Next i
	
End Function
Function Extract_FinalConsonant( b() as Byte )
	
End Function
