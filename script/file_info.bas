REM  *****  BASIC  *****

'
' REF : https://help.libreoffice.org/latest/ro/text/sbasic/shared/03/lib_tools.html?&DbPAR=SHARED&System=WIN#strings_module
'

Sub main

GlobalScope.BasicLibraries.loadLibrary("Tools")

Dim a as String
a = ( _
		Chr(10) _
	&	"+ Path" & Chr(10) _
	& 	ThisComponent.getURL() _
	& 	Chr(10) & Chr(10) _	
	& 	"+ Directory" & Chr(10) _
	&	 Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/") _
	& 	Chr(10) & Chr(10) _
	&	"+ File Name + Extension" & Chr(10) _
	& 	Tools.Strings.FileNameoutofPath(ThisComponent.getURL(),"/") _
	& 	Chr(10) _
	& 	Tools.Strings.GetFileNameWithoutExtension(ThisComponent.getURL(),"/") _
	& 	Chr(10) & Chr(10) _	
	&	"+ Extension" & Chr(10) _
	&  	Tools.Strings.GetFileNameExtension(ThisComponent.getURL(),"/") _
	& 	Chr(10) & Chr(10) _	
)

MsgBox a

End Sub
