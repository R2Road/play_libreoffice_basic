REM  *****  LibreOffice VBA  *****

Option VBASupport 1 'for Len, StrConv
Option Explicit

'
'REF : https://leesumin.tistory.com/78
'

'
' 한글 결합식
'
' (초성 인덱스 * 21 + 중성 인덱스) * 28 + 종성 인덱스 + 0xAC00( 44032 : 가 )
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

'
' 분해 가능한 한글 범위 : AC00( 가 : 44032 ) ~ D7A3( 힣 : 55203 )
' 


Private list_initial_consonaant( 19 ) as String
Private list_vowel( 21 ) as String
Private list_final_consonaant( 28 ) as String

Function InitKoreanPartsList
	
	If list_initial_consonaant( 0 ) = "" Then
		
		list_initial_consonaant = Array( "ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ", "ㅅ", "ㅆ", "ㅇ" , "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ" )
		list_vowel = Array( "ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ", "ㅗ", "ㅘ", "ㅙ", "ㅚ", "ㅛ", "ㅜ", "ㅝ", "ㅞ", "ㅟ", "ㅠ", "ㅡ", "ㅢ", "ㅣ" )
		list_final_consonaant = Array( "", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ", "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ" )
		
	EndIf
	
End Function



Function ConvertBytes2Code( b() as Byte )

	'
	' byte array를 하나의 수로 만든다.
	'
	
	Dim code as Long 'Integer : 16bit, Long : 32bit
	
	'
	' b( 1 )
	'
	code = b( 1 )
	code = code * 256 ' 256 : 2의 8승 : 왼쪽 shift 8
	
	'
	' b( 0 )
	'
	code = code + b( 0 )
	
	ConvertBytes2Code = code
		
End Function

Function IsDecompositionEnable( code as Long )

	IsDecompositionEnable = ( code >= 44032 And code <= 55203 )

End Function



Function research___multibyte___extract_initial_consonant '초성 : initial_consonant

	InitKoreanPartsList

	Dim s as String : s = "가네민ㄷㅏ"
	
	Dim slen as Integer : slen = Len( s )
	
	Dim b() as Byte
	Dim code as Long
	
	Dim result as String
	
	Dim i as Integer
	For i = 1 To slen
		b = Mid( s, i, 1 )
		
		code = ConvertBytes2Code( b )
		
		If IsDecompositionEnable( code ) Then
			result = result & "+ " & b & Chr( 10 ) & "0 : " & b( 0 ) & Chr( 10 ) & "1 : " & b( 1 ) & Chr( 10 ) & list_initial_consonaant( Extract_InitialConsonant( code ) ) & Chr( 10 ) & Chr( 10 )
		Else
			result = result & "+ " & b & Chr( 10 ) & "분해 불가" & Chr( 10 ) & Chr( 10 )
		End If
	Next i
	
	MsgBox( result )
	
End Function
Function Extract_InitialConsonant( code as Long )
	
	'
	' 한글 결합식
	'
	' (초성 인덱스 * 21 + 중성 인덱스) * 28 + 종성 인덱스 + 0xAC00( 44032 : 가 )
	'
	
	'
	' 가 : 44032
	' 각 항목의 인덱스가 모두 0 일때 '가' 이다.
	'
	code = Int( code - 44032 ) '은근슬쩍 반올림을 하고 있어서 Int 를 사용 해서 정수부만 쓰도록 제한한다.
	
	'
	' 종성 떨구기
	'
	code = Int( code / 28 )
	
	'
	' 중성 떨구기
	'
	code = Int( code / 21 )
	
	Extract_InitialConsonant = code
	
End Function



Function research___multibyte___extract_vowel '모음 : vowel

	Dim s as String : s = "가나민ㄷㅏ"
	Dim b() as Byte
	
	Dim result as String
	
	Dim i as Integer
	For i = 1 To 4
		b = Mid( s, i, 1 )
		result = result & "+ " & b & Chr( 10 ) & "0 : " & b( 0 ) & Chr( 10 ) & "1 : " & b( 1 ) & Chr( 10 ) & Extract_Vowel( b ) & Chr( 10 ) & Chr( 10 )
	Next i
	
	MsgBox( result )
	
End Function
Function Extract_Vowel( b() as Byte )
	
End Function



Function research___multibyte___extract_final_consonant '종성 : final_consonant

	Dim s as String : s = "가나민ㄷㅏ"
	Dim b() as Byte
	
	Dim result as String
	
	Dim i as Integer
	For i = 1 To 4
		b = Mid( s, i, 1 )
		result = result & "+ " & b & Chr( 10 ) & "0 : " & b( 0 ) & Chr( 10 ) & "1 : " & b( 1 ) & Chr( 10 ) & Extract_FinalConsonant( b ) & Chr( 10 ) & Chr( 10 )
	Next i
	
	MsgBox( result )
	
End Function
Function Extract_FinalConsonant( b() as Byte )
	
End Function
