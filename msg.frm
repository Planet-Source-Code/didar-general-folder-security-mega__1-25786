VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "©Copyright By General Corporation 2001."
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   2925
      End
      Begin VB.Label Label3 
         Caption         =   "This Folder Is Protected By General Security Folder"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "General Warning!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access Is Denied"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   2115
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   2520
   End
End
Attribute VB_Name = "Form1"
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


Private Sub Timer1_Timer()
Unload Me
End Sub
