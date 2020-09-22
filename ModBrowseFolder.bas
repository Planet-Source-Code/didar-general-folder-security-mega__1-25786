Attribute VB_Name = "ModBrowseFolder"
Option Explicit

Public Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type

Public Declare Function SHGetPathFromIDList Lib _
   "shell32.dll" Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Public Declare Function SHBrowseForFolder Lib _
   "shell32.dll" Alias "SHBrowseForFolderA" _
   (lpBrowseInfo As BROWSEINFO) As Long


   




