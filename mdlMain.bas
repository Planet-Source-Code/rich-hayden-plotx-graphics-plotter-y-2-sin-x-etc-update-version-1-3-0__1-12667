Attribute VB_Name = "mdlMain"
Declare Function GetProfileString Lib "kernel32" _
     Alias "GetProfileStringA" (ByVal lpAppName As _
     String, ByVal lpKeyName As String, ByVal _
     lpDefault As String, ByVal lpReturnedString As _
     String, ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" _
     Alias "WriteProfileStringA" (ByVal lpszSection _
     As String, ByVal lpszKeyName As String, ByVal _
     lpszString As String) As Long

Declare Function GetProfileSection Lib "kernel32" _
     Alias "GetProfileSectionA" (ByVal lpAppName _
     As String, ByVal lpReturnedString As String, _
     ByVal nSize As Long) As Long

Declare Function WriteProfileSection Lib "kernel32" _
     Alias "WriteProfileSectionA" (ByVal lpAppName _
     As String, ByVal lpString As String) As Long
     
Declare Function GetPrivateProfileInt Lib "kernel32" _
     Alias "GetPrivateProfileIntA" (ByVal _
     lpApplicationName As String, ByVal lpKeyName As _
     String, ByVal nDefault As Long, ByVal lpFileName _
     As String) As Long

Declare Function GetPrivateProfileString Lib _
     "kernel32" Alias "GetPrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal _
     lpKeyName As Any, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, ByVal lpFileName As String) _
     As Long

Declare Function WritePrivateProfileString Lib _
     "kernel32" Alias "WritePrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal _
     lpKeyName As Any, ByVal lpString As Any, ByVal _
     lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib _
     "kernel32" Alias "GetPrivateProfileSectionA" _
     (ByVal lpAppName As String, ByVal lpReturnedString _
     As String, ByVal nSize As Long, _
     ByVal lpFileName As String) As Long
     
Declare Function WritePrivateProfileSection Lib _
     "kernel32" Alias "WritePrivateProfileSectionA" _
     (ByVal lpAppName As String, ByVal lpString As _
     String, ByVal lpFileName As String) As Long

 


Public Sub WriteIniInfo(iniSection As String, iniItem As String, ItemValue As String, iniFile)
    On Error GoTo EH
    Dim X As Long
    X = WritePrivateProfileString(iniSection, iniItem, ItemValue, iniFile)
    Exit Sub
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: M" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Sub

Public Function GetIniInfo(iniFile As String, Section As String, ItemReturn As String, DefaultValue As String) As String
    On Error GoTo EH
    Dim lResult As Long
    Dim sIniString As String
    sIniString = String(20, 0)
    lResult = GetPrivateProfileString(Section, ItemReturn, DefaultValue, sIniString, Len(sIniString), iniFile)
    sIniString = Left$(sIniString, InStr(sIniString, Chr$(0)) - 1)
    GetIniInfo = sIniString
    Exit Function
EH:
    MsgBox "Fatal Error in PlotX. Please re-install PlotX or alternatively, you may send an e-mail to Richard Hayden, at r_hayden@hotmail.com. Please include full details of the error; the error code (below), your product version (below) and error message (below), and if possible, how to replicate it. (It would be very advisable to write the below details down, before hitting OK, as you may not be able to easily replicate the error.)" & Chr(13) & Chr(13) & "Error Code: L" & Chr(13) & Chr(13) & "Product Version: " & App.Major & "." & App.Minor & "." & App.Revision & Chr(13) & Chr(13) & "Error Message: " & Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "Please remember that the creator of this program, Richard Hayden, has NO guarantee to you to correct any problems and you have NO warranty agreement with this program. I will try and correct any problems, but can not guarantee anything.", vbCritical + vbOKOnly + vbSystemModal, "PlotX"
    End
End Function
