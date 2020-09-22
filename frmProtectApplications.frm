VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProtectApplications 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Folder Security"
   ClientHeight    =   4335
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3975
   Icon            =   "frmProtectApplications.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   6960
      TabIndex        =   32
      Top             =   1320
      Width           =   615
   End
   Begin VB.Timer Timer6 
      Left            =   360
      Top             =   3480
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   7080
      TabIndex        =   31
      Top             =   2160
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   5280
      TabIndex        =   30
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   5280
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   2400
      TabIndex        =   18
      Top             =   1080
      Width           =   1455
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveFolder 
         Caption         =   "Remove.."
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddFolder 
         Caption         =   "Add Folders"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   4320
      Width           =   1455
   End
   Begin VB.ListBox ListWindows 
      Height          =   1230
      Left            =   4320
      TabIndex        =   15
      Top             =   2880
      Width           =   3350
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   7000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   3120
      ScaleHeight     =   2595
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   6885
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtCopyText2 
         Height          =   735
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1440
         Top             =   360
      End
      Begin VB.TextBox txtTimer2 
         Height          =   285
         Left            =   0
         TabIndex        =   10
         Text            =   "0"
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCount2 
         Caption         =   "0"
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   5400
      Width           =   975
   End
   Begin VB.ListBox ListFoldersApplications 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   1980
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2145
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   4800
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save "
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtTimer 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   6885
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   4800
   End
   Begin VB.TextBox txtCopyText 
      Height          =   1335
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   6885
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   1
      Top             =   6885
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstTasks 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   6885
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1920
      Top             =   5160
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2280
      Top             =   4800
   End
   Begin MSComDlg.CommonDialog com1 
      Left            =   3465
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   4680
      TabIndex        =   33
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "                   All Rights Reserved.                          © Copyright by General Corporation.2001"
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright by General Corporation"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Folder Security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   465
      TabIndex        =   24
      Top             =   90
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Folder Security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   23
      Top             =   120
      Width           =   2865
   End
   Begin VB.Label lblIndex 
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   6885
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Protected Folders.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label lblCount 
      Caption         =   "0"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   6885
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu change 
         Caption         =   "Change Password"
      End
      Begin VB.Menu Add 
         Caption         =   "Add Folders"
      End
      Begin VB.Menu Remove 
         Caption         =   "Remove"
      End
      Begin VB.Menu endtask 
         Caption         =   "Disable End Task"
      End
      Begin VB.Menu edn 
         Caption         =   "Enable End Task"
      End
      Begin VB.Menu auto 
         Caption         =   "Set As Always Open Autmatically"
      End
      Begin VB.Menu never 
         Caption         =   "Never Run Automatically"
      End
      Begin VB.Menu Disable 
         Caption         =   "Disable Folder Delete"
      End
      Begin VB.Menu Enable1 
         Caption         =   "Enable Folder Delete"
      End
      Begin VB.Menu file 
         Caption         =   "Disable File Delete"
      End
      Begin VB.Menu Enable 
         Caption         =   "Enable File Delete"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu sys 
      Caption         =   "System"
      Begin VB.Menu msconfig 
         Caption         =   "Disable Msconfig"
      End
      Begin VB.Menu econfig 
         Caption         =   "Enable Msconfig"
      End
      Begin VB.Menu dregedit 
         Caption         =   "Disable Regedit"
      End
      Begin VB.Menu eregedit 
         Caption         =   "Enable Regedit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu Abou 
         Caption         =   "About"
      End
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmProtectApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'              General Folder Security
'  Author : Didarul Alam ( Bangladesh )
'Senior Programmer: General Corporation Bangladesh.

'I've edit this code.I didn't write the whole code.
'Some modules I've collected from PSC.

'I think this 'General Folder Security' so strong then ever
'ever before. Becuase if anybody change the locked folders
'Under any other operating system, then 'General Folder Security'
'will not allow that user to access the computer..

'This code is dedicated to Tahmina Nur Chowdhury
'She is my love.

Option Explicit
Dim Worked As Boolean





Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
 Const HKEY_LOCAL_MACHINE = &H80000002
 Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2











Private Const SW_HIDE = 0
Private Const GW_OWNER = 4
Public apiError As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97
Dim x, X1 As Integer
Dim filenumber As Integer
Dim rtn As Long
Dim FileNum, UserInput, i
Dim fName As String
Dim fPath As String
Dim bi As BROWSEINFO
Dim r As Long
Dim pidl As Long
Dim tmpPath As String
Dim pos As Integer








Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub













Private Function ValidateDir(tmpPath As String) As String
          ValidateDir = tmpPath
End Function

Private Function vbGetBrowseDirectory() As String
    pidl = SHBrowseForFolder(bi)
    tmpPath = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl, ByVal tmpPath)
    If r Then
          pos = InStr(tmpPath, Chr$(0))
          tmpPath = Left(tmpPath, pos - 1)
          vbGetBrowseDirectory = ValidateDir(tmpPath)
    Else: vbGetBrowseDirectory = ""
    End If
End Function

Private Sub Abou_Click()
MsgBox "All Rights Reserved. © Copyright by General Corporation Bangladesh 2001.", 32, "About"
End Sub

Private Sub Add_Click()
cmdAddFolder_Click
End Sub

Private Sub auto_Click()
Command2_Click
    r = WritePrivateProfileString("password", "status", "1", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
End Sub

Private Sub change_Click()
Dim didar As String
Dim r As Long
didar = InputBox("Please Enter The New Password", "New Password", "gsi911")
   If didar = "" Then
   MsgBox "You didn't enter any Password", 16, "Error"
   Exit Sub
   End If
    r = WritePrivateProfileString("password", "pass", didar, iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
End Sub

Private Sub cmdAddFolder_Click()
mnuAddFolder_Click
Command5_Click
End Sub

Private Sub cmdExit_Click()
On Error GoTo errhandler
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End
Exit Sub
errhandler:
End Sub


Private Sub cmdSave_Click()
txtCopyText.Text = ""
For i = 0 To ListFoldersApplications.ListCount - 1
    For x = 0 To ListFoldersApplications.ListCount - 1
    If i = x Then GoTo Nextx
        If LCase(ListFoldersApplications.List(x)) = LCase(ListFoldersApplications.List(i)) Then
        ListFoldersApplications.RemoveItem x
    End If
Nextx:
    Next x
Next i
    For i = 0 To ListFoldersApplications.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & ListFoldersApplications.List(i) & vbCrLf
        Next i
FileNum = FreeFile
Open App.Path & "\Folder.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum
Timer3.Enabled = True
End Sub

Private Sub cmdRemoveFolder_Click()
mnuRemoveFolder_Click
End Sub




Private Sub mnuAddFolder_Click()
On Error Resume Next

fPath$ = vbGetBrowseDirectory$()
    If fPath > "" Then
      fName$ = fPath
      
      
      SetAttr fName, vbReadOnly
      
      
      
      ListFoldersApplications.AddItem LCase(FileName(fName))
For i = 0 To ListFoldersApplications.ListCount - 1
    For x = 0 To ListFoldersApplications.ListCount - 1
    If i = x Then GoTo Nextx
        If LCase(ListFoldersApplications.List(x)) = LCase(ListFoldersApplications.List(i)) Then
        ListFoldersApplications.RemoveItem x
    End If
Nextx:
    Next x
Next i
txtCopyText.Text = ""

    For i = 0 To ListFoldersApplications.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & ListFoldersApplications.List(i) & vbCrLf
        Next i
        
FileNum = FreeFile
Open App.Path & "\Folder.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum
Timer3.Enabled = True
      End If
End Sub

Private Sub Command1_Click()
Me.Hide

End Sub

Private Sub Command2_Click()
On Error Resume Next
If Command2.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GMFolder", "GMFolder"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GMFolder", App.Path & "\GMfolder.exe"

End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Command3.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GMFolder", "GMFolder"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GMFolder", "0"
End If
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
On Error Resume Next

'fPath$ = vbGetBrowseDirectory$()
 '   If fPath > "" Then
  '    fName$ = fPath
      
     
   '   SetAttr fName, vbReadOnly
      
      'List1.AddItem LCase(FileName(fName))
      List1.AddItem LCase(fName)
      
      
For i = 0 To List1.ListCount - 1
    For x = 0 To List1.ListCount - 1
    If i = x Then GoTo Nextx
        If LCase(List1.List(x)) = LCase(List1.List(i)) Then
        List1.RemoveItem x
    End If
Nextx:
    Next x
Next i
txtCopyText.Text = ""

    For i = 0 To List1.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & List1.List(i) & vbCrLf
               
        Next i
        
FileNum = FreeFile
Open App.Path & "\history.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum
Timer3.Enabled = True
      
      'End If

End Sub

Private Sub Command6_Click()
Dim didar, check As String


On Error GoTo Err
For i = 0 To List1.ListCount - 1
If List1.List(i) = "" Then
Label6.Caption = List1.List(i)
Else
SetAttr List1.List(i), vbReadOnly
End If
Next i


Exit Sub
Err:
Form2.show
MsgBox "You Are Not A Valid User. You Have Changed Some Locked Folders In Different Operating System.", 16, "Illegal User"
  didar = GetFromINI("password", "pass", iniPath$)
  
  
  check = InputBox("General Folder Security Found You As An Illegal User.Because You Have Changed Some Locked Folders.So System Is Going To Shutdown Immediately.If You Are A Registered User,Please Enter Your Password...", "Enter Password")
  If check = didar Then
  Unload Form2
    Me.show
  Else
    MsgBox "Not A Valid User. Windows Is Now Going To Shutdown...", 16, "Error"
  i = Shell("c:\windows\rundll.exe user.exe,exitwindows", vbNormalFocus)
  End If
  
  
End Sub


Private Sub Disable_Click()
    r = WritePrivateProfileString("password", "disable", "1", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
    Check1.Value = 1
    End Sub

Private Sub dregedit_Click()
On Error Resume Next
Name "c:\windows\regedit.exe" As "c:\windows\regedit.dat"
End Sub

Private Sub econfig_Click()
On Error Resume Next
Name "c:\windows\system\msconfig.dat" As "c:\windows\system\msconfig.exe"
End Sub

Private Sub edn_Click()
Dim Ret As Integer
Dim pOld As Boolean


    r = WritePrivateProfileString("password", "end", "0", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
Check3.Value = 0
End Sub

Private Sub enable_Click()
    r = WritePrivateProfileString("password", "file", "0", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
    Check2.Value = 0
    End Sub

Private Sub Enable1_Click()
    r = WritePrivateProfileString("password", "disable", "0", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
    Check1.Value = 0
End Sub

Private Sub endtask_Click()
Dim Ret As Integer
Dim pOld As Boolean

    r = WritePrivateProfileString("password", "end", "1", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."

Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
Check3.Value = 1
End Sub

Private Sub eregedit_Click()
On Error Resume Next
Name "c:\windows\regedit.dat" As "c:\windows\regedit.exe"
End Sub

Private Sub file_Click()
    r = WritePrivateProfileString("password", "file", "1", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
    Check2.Value = 1
End Sub

Private Sub Help_Click()
MsgBox "              HELP" & Chr(13) & Chr(13) & "    Click 'Add Folder' to select the folder to lock.", 32, "Info"
End Sub


Private Sub mnuExit_Click()
End
End Sub


Private Sub mnuLockFiles_Click()
Dim fileLock As String
Dim filenumber
On Error GoTo errhandler
    Open "C:\Folder\file.txt" For Input As #1


    Do While Not EOF(1)
        Line Input #1, fileLock
        filenumber = FreeFile
        Open fileLock For Binary Shared As #filenumber
         Lock #filenumber
    Loop
    Close #1

    Exit Sub
errhandler:
End Sub



Private Sub mnuRemoveFolder_Click()
Dim a As Integer

On Error Resume Next


a = ListFoldersApplications.ListIndex

ListFoldersApplications.RemoveItem ListFoldersApplications.ListIndex
txtCopyText.Text = ""
    Dim i As Integer
    For i = 0 To ListFoldersApplications.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & ListFoldersApplications.List(i) & vbCrLf
        Next i
FileNum = FreeFile
Open App.Path & "\Folder.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum





'aaaaaaaaaaaaaaaaaaaaaaa


List1.RemoveItem a
txtCopyText.Text = ""
    
    For i = 0 To List1.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & List1.List(i) & vbCrLf
        Next i
FileNum = FreeFile
Open App.Path & "\history.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum




End Sub

Private Sub pGetTasks()
    Call fEnumWindows(lstTasks)
    On Error Resume Next
    lstTasks.ListIndex = -1
End Sub


Private Sub mnuREmoveFolders_Click()
ListFoldersApplications.Clear
txtCopyText.Text = ""
FileNum = FreeFile
Open App.Path & "\Folder.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum
End Sub

Private Sub Form_Load()
Dim didar As String
Dim r As Long
Dim check As String
Dim flag As String
Dim flag2 As String
Dim flag3 As String
Dim flag4 As String
Dim Ret As Integer
Dim pOld As Boolean



'On Error GoTo errhandler
Call sCenterForm(Me)
App.TaskVisible = False
On Error Resume Next
Module1.ListOpen ListFoldersApplications, App.Path & "\folder.txt"


Module1.ListOpen List1, App.Path & "\history.txt"


'FileNum = FreeFile
'Open App.Path & "\password.dat" For Input As FileNum
'txtPassword.Text = Input(LOF(FileNum), FileNum)
'Close FileNum


Me.show
With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = 512
        .hIcon = Me.Icon
        .szTip = "General Folder Security. General Corporation 2001." & Chr(0)
    End With
    apiError = Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA)
'mnuLockFiles_Click







iniPath$ = App.Path + "\rwini.ini"
  didar = GetFromINI("password", "pass", iniPath$)





Command6_Click




flag2 = GetFromINI("password", "disable", iniPath$)
If flag2 = "1" Then
Check1.Value = 1
End If

flag3 = GetFromINI("password", "file", iniPath$)
If flag3 = "1" Then
Check2.Value = 1
End If



flag4 = GetFromINI("password", "end", iniPath$)
If flag4 = "1" Then
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
Check3.Value = 1
End If



flag = GetFromINI("password", "status", iniPath$)
If flag = "1" Then
Me.Hide
Exit Sub
End If





    
If didar = "" Then
didar = InputBox("Please Enter The New Password", "New Password", "gsi911")
  
   If didar = "" Then
   MsgBox "You didn't enter any Password", 16, "Error"
   Exit Sub
   End If

    r = WritePrivateProfileString("password", "pass", didar, iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
End If


  check = InputBox("Please Enter Passord", "Enter Password")
  If check = didar Then
  Else
  MsgBox "Not a valid user", 16, "Error"
  End
  Exit Sub
  End If





Exit Sub
errhandler:
Me.show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim didar, check As String
Dim tmpLong As Single

On Error Resume Next


  tmpLong = x / Screen.TwipsPerPixelX
    
    Select Case tmpLong
        Case WM_LBUTTONUP
            apiError = SetForegroundWindow(Me.hWnd)
            
            
            
            
  didar = GetFromINI("password", "pass", iniPath$)
  check = InputBox("Please Enter Passord", "Enter Password")
  If check = didar Then
  Me.show
    Else
  MsgBox "Not a valid user", 16, "Error"
  Exit Sub
  End If

            
            
            
            
            
'            UserInput = InputBox("1Password")
 '          If UserInput = txtPassword.Text Then
  '          Me.show
            
   '         Else
    '        MsgBox "Sorry Wrong password"
     '       End If
            
        Case WM_RBUTTONUP
            apiError = SetForegroundWindow(Me.hWnd)
            
  didar = GetFromINI("password", "pass", iniPath$)
  check = InputBox("Please Enter Passord", "Enter Password")
  If check = didar Then
  Me.show
    Else
  MsgBox "Not a valid user", 16, "Error"
  Exit Sub
  End If
      
      
    End Select
 End Sub

Private Sub Form_Unload(Cancel As Integer)
With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = vbNull
        .hIcon = Me.Icon
        .szTip = Chr(0)
    End With
    apiError = Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA)
End Sub


Private Sub cmdAddFromWinList_Click()
Dim num As Integer
num = InputBox("Please pick from Window list", "Window Picker")
num = Val(num) - 1
ListFoldersApplications.AddItem ListWindows.List(num)
End Sub

Public Sub mnuLockFolders_Click()
Dim lock1 As Integer
For lock1 = 0 To ListFoldersApplications.ListCount - 1
ListFoldersApplications.ItemData(lock1) = 0
Next lock1
End Sub

Private Sub mnuShowTaskBar_Click()
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar

End Sub

Private Sub mnuSysedit_Click()
ListFoldersApplications.AddItem "system configuration editor"
cmdSave_Click
End Sub

Private Sub mnuUnlockFiles_Click()
On Error GoTo errhandler
Dim FileNum
FileNum = FreeFile
Open "C:\Folder\file.txt" For Input As FileNum
    For x = 1 To FreeFile - 1
    Close #x
    Next x
    Close #FileNum
Exit Sub
errhandler:
End Sub

Public Sub mnuUnloackFolders_Click()
Dim unlock1 As Integer
For unlock1 = 0 To ListFoldersApplications.ListCount - 1
ListFoldersApplications.ItemData(unlock1) = 1
Next unlock1
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++
' MSCONFIG utility and REGEDIT can disable 'General Folder Security' or any other
'Folder protection software ( Which has startup mode). So it's necessary to disable this programs..
'+++++++++++++++++++++++++++++++++++++++++++++++


Private Sub msconfig_Click()
On Error Resume Next
Name "c:\windows\system\msconfig.exe" As "c:\windows\system\msconfig.dat"
End Sub

Private Sub never_Click()
Command3_Click
    r = WritePrivateProfileString("password", "status", "0", iniPath$)
    If r <> 1 Then MsgBox "An error occurred while writing SerialNumber."
End Sub

Private Sub Remove_Click()
cmdRemoveFolder_Click
End Sub

Private Sub Timer1_Timer()
Call pGetTasks
showfile

If Check1.Value = 1 Then
Worked = ShutDownApplication("Confirm Folder Delete")
End If

If Check2.Value = 1 Then
Worked = ShutDownApplication("Confirm file Delete")
Worked = ShutDownApplication("Confirm Multiple File Delete")
Worked = ShutDownApplication("Confirm File rename")
Worked = ShutDownApplication("Confirm Folder move")
End If



End Sub

Private Sub Timer2_Timer()
Dim winname, pro As String
winname = ListWindows.List(0)

'If ListFoldersApplications.ListCount = 0 Then
'Exit Sub
'Else


For X1 = 0 To ListFoldersApplications.ListCount - 1
pro = ListFoldersApplications.List(X1)
If winname = pro Then


            If pro = "" Then
            Exit Sub
            End If

access
Exit Sub
ElseIf winname = "exploring - " + pro Then
access
Exit Sub
End If
Next X1

'End If

End Sub

Private Sub access()
Dim x As Integer
Dim Y As Integer
Dim pass As String
Dim z As Integer
Dim hid As Long
On Error Resume Next
hid = lstTasks.ItemData(0)
Y = ListFoldersApplications.ItemData(X1)
If Y = 0 Then
HideWin hid
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Tasbar
Form1.show
'pass = InputBox("Password", "3Password")   'Always wants password.
'If pass = txtPassword.Text Then
'ListFoldersApplications.ItemData(X1) = 1
'ShowWin hid
'rtn = FindWindow("Shell_traywnd", "") 'get the Window
'Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
'Else
'MsgBox "Invalid Password. The file or folder you tried to open will now close", , "Password"
CloseWin hid
ListFoldersApplications.ItemData(X1) = 0
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
'End If
End If

End Sub



Public Sub showfile()
Dim show As Integer
ListWindows.Clear


For show = 0 To lstTasks.ListCount - 1
ListWindows.AddItem LCase(lstTasks.List(show))
Next show
End Sub


'Private Function GetFileName(WithString)
'Dim WithoutString, AllLen, Where As String
'WithoutString = WithString
'Do Until InStr(WithoutString, "\") = 0
'AllLen = Len(WithoutString)
'Where = InStr(WithoutString, "\")
'WithoutString = Right(WithoutString, AllLen - Where)
'Loop
'GetFileName = WithoutString
'End Function





