VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain2 
   Caption         =   "Sourcecode Notebook"
   ClientHeight    =   9120
   ClientLeft      =   1320
   ClientTop       =   1950
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   14145
   Begin VB.Frame frameFilters 
      Caption         =   "Filter By:"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   4965
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   3165
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   10800
      Top             =   -300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0336
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0448
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":055A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":066C
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabControl 
      Height          =   8655
      Left            =   5040
      TabIndex        =   3
      Top             =   420
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   15266
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Code Window"
      TabPicture(0)   =   "frmMain2.frx":077E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "RichTextBox1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "frmMain2.frx":079A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   8000
         Left            =   30
         TabIndex        =   4
         Top             =   660
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   14102
         _Version        =   393217
         TextRTF         =   $"frmMain2.frx":07B6
      End
   End
   Begin MSComctlLib.ListView lstTitles 
      Height          =   8025
      Left            =   0
      TabIndex        =   5
      Top             =   1050
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   14155
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            'ToDo: Add 'New' button code.
            MsgBox "Add 'New' button code."
        Case "Open"
            'ToDo: Add 'Open' button code.
            MsgBox "Add 'Open' button code."
        Case "Save"
            'ToDo: Add 'Save' button code.
            MsgBox "Add 'Save' button code."
        Case "Copy"
            'ToDo: Add 'Copy' button code.
            MsgBox "Add 'Copy' button code."
        Case "Paste"
            'ToDo: Add 'Paste' button code.
            MsgBox "Add 'Paste' button code."
        Case "Print"
            'ToDo: Add 'Print' button code.
            MsgBox "Add 'Print' button code."
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            MsgBox "Add 'Delete' button code."
    End Select
End Sub

