Attribute VB_Name = "mdlRead"
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub Drag_Form(frmForm As Form)
  ReleaseCapture
  SendMessage frmForm.hwnd, &HA1, 2, 0
End Sub

Public Function ReadINI(Section, KeyName, filename As String) As String
Dim sRet As String
  
  sRet = String(998, Chr(0))
  ReadINI = Left(sRet, GetPrivateProfileString(Section, KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(Section, KeyName, NewString As String, filename As String) As String
  Dim sWet As String
  sWet = WritePrivateProfileString(Section, KeyName, NewString, filename)
End Function

