VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1800
      Left            =   2100
      Top             =   2460
   End
   Begin VB.Frame Frame1 
      Height          =   3030
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   570
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   2100
         Width           =   480
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2850
         TabIndex        =   1
         Top             =   2370
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "MS Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         Caption         =   "Sourcecode Notebook"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   30
         TabIndex        =   3
         Top             =   150
         Width           =   4095
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "Sourcecode Notebook"
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    'dont want to have to click, should go away after 3 seconds
    Unload Me
End Sub
