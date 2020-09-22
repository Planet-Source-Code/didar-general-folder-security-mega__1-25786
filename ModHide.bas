Attribute VB_Name = "ModHide"
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
    ' ----Public Declares for this code
    Public Const RSP_SIMPLE_SERVICE = 1
    Public Const RSP_UNREGISTER_SERVICE = 0
  
Public Sub Hide_Program_In_CTRL_ALT_DELETE()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub Show_Program_In_CTRL_ALT_DELETE()
    Dim pid As Long
    Dim reserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub



