Attribute VB_Name = "inieditor"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Function WRITEINI(ByVal SECTION As String, KEY As String, value As String, INIFILE As String)
WritePrivateProfileString SECTION, KEY, value, INIFILE
End Function
Public Function GETINI(ByVal SECTION As String, KEY As String, INIFILE As String, ByRef rVALUE As String) As Boolean
Dim value As String * 256
Dim a As Long
a = GetPrivateProfileString(SECTION, _
    KEY, "?!?", value, 256, _
    INIFILE)
If Left$(value, 3) = "?!?" Then
    GETINI = False
    
Else
    GETINI = True
    rVALUE = Left$(value, a)
End If
End Function
