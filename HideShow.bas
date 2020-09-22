Attribute VB_Name = "Module3"
Dim rtn As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Public Sub sCenterForm(tmpF As Form)

Dim x As Integer, Y As Integer
Y = (Screen.Height - tmpF.Height) \ 2
x = (Screen.Width - tmpF.Width) \ 2

tmpF.Move x, Y

End Sub


