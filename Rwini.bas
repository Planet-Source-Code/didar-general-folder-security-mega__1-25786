Attribute VB_Name = "rwinin"
Option Explicit
'Read-Write INI Sample
'Written by: George Csefai-Keane, Inc.
'email: george.csefai@keaneinc.com

'API Declarations
 Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'Variable Declarations
Global r%       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Sub CenterForm(frm As Form)
    frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
    frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function


