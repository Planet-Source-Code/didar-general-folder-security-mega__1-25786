Attribute VB_Name = "ModKilApp"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_CLOSE = &H10


Public Function ShutDownApplication(ByVal ApplicationName As String) As Boolean

Dim hWnd As Long, Result As Long
hWnd = FindWindow(vbNullString, ApplicationName)
If hWnd <> 0 Then
        Result = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
        ShutDownApplication = True
    End If

End Function

Public Function FileName(WithPath)
Dim WithoutPath, AllLen, Where As String
WithoutPath = WithPath
Do Until InStr(WithoutPath, "\") = 0
AllLen = Len(WithoutPath)
Where = InStr(WithoutPath, "\")
WithoutPath = Right(WithoutPath, AllLen - Where)
Loop
FileName = WithoutPath
End Function






