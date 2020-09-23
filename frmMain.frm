VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Sourcecode Notebook"
   ClientHeight    =   9210
   ClientLeft      =   1635
   ClientTop       =   2055
   ClientWidth     =   11565
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   9210
   ScaleWidth      =   11565
   Begin VB.Frame frmSnippit 
      Caption         =   "Snippit"
      Height          =   6585
      Left            =   45
      TabIndex        =   2
      Top             =   2340
      Width           =   11445
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Width           =   8805
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmMain.frx":0C0C
         Left            =   8880
         List            =   "frmMain.frx":0C0E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   2355
      End
      Begin TabDlg.SSTab tabSnippit 
         Height          =   5685
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   750
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   10028
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   529
         TabMaxWidth     =   3528
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Source"
         TabPicture(0)   =   "frmMain.frx":0C10
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "rtbCodeWindow"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Notes"
         TabPicture(1)   =   "frmMain.frx":0C2C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "rtbNotes"
         Tab(1).ControlCount=   1
         Begin RichTextLib.RichTextBox rtbCodeWindow 
            Height          =   5295
            Left            =   30
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   9340
            _Version        =   393217
            ScrollBars      =   3
            RightMargin     =   1.00000e5
            TextRTF         =   $"frmMain.frx":0C48
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox rtbNotes 
            Height          =   5295
            Left            =   -74940
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   9340
            _Version        =   393217
            Enabled         =   -1  'True
            RightMargin     =   1.00000e5
            TextRTF         =   $"frmMain.frx":0D10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ListView lstTitles 
      Height          =   1905
      Left            =   60
      TabIndex        =   10
      Top             =   390
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   3360
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   11447
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code Type"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdoOpenDatabase 
      Left            =   10230
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "Open Database"
      Filter          =   "Access Database Files|*.mdb|All Files|*.*"
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   9300
      Top             =   1530
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DD8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EEA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FFC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":110E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1220
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1332
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1444
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1556
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      Left            =   8940
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   2354
   End
   Begin MSComctlLib.Toolbar toolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New snippit"
            Object.Tag             =   "new"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open database"
            Object.Tag             =   "open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save snippit"
            Object.Tag             =   "save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export"
            Object.Tag             =   "copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste snippit"
            Object.Tag             =   "paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete snippit"
            Object.Tag             =   "delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find"
            Object.Tag             =   "find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            Object.Tag             =   "print"
            ImageKey        =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   8925
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/23/99"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:03 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFilter 
      Caption         =   "Filter By:"
      Height          =   225
      Left            =   8940
      TabIndex        =   8
      Top             =   360
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Snippit"
      Index           =   2
      Begin VB.Menu mnuNew 
         Caption         =   "&New code snippit"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Export"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste Snippit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete current snippit"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save current snippit"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuType 
      Caption         =   "&Code Type"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Code Type"
      End
      Begin VB.Menu mnuModify 
         Caption         =   "&Modify / Delete Code Type"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuWebsite 
         Caption         =   "&Website"
      End
      Begin VB.Menu mnuPlanetSourceCode 
         Caption         =   "&Planet Source Code"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFilter_Click()
    LoadGridBox
End Sub

Private Sub Form_Load()
    'frmMain2.Show vbModal  'this was just another form design I was working on.
    frmSplash.Show vbModal, frmMain 'if you dont like splash screens just comment this out.
    gblConnectString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & AppPath & "sourcebook.mdb"
    gblNewCode = True 'start off with a clean slate
    
    LoadCodeTypes
    rtbNotes.OLEDropMode = 1  'setup for the drag drop in the code windows
    rtbCodeWindow.OLEDropMode = 1
        
End Sub

Private Sub LoadGridBox()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim cmdtext As String
    Dim lstItem As ListItem

    lstTitles.ListItems.Clear
    ' here we are building the SQL statement based upon the filter drop down
    If cmbFilter.Text = "No Filter" Then
        cmdtext = "SELECT id, title, codetype FROM source "
    Else
        cmdtext = "SELECT id, title, codetype FROM source WHERE codetype='" & StuffQuotes(cmbFilter.Text) & "' "
    End If
    
    'connect to the database and retrieve the code
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCon.CursorLocation = adUseClientBatch
    adoCmd.CommandText = cmdtext
    Set adoRS = adoCmd.Execute
    
    'loop through the recordset and add each item to the listview
    Do While Not adoRS.EOF
        Set lstItem = lstTitles.ListItems.Add(, , adoRS("title"))
        lstItem.Tag = adoRS("id") 'used for updating and deleting
        lstItem.SubItems(1) = CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    lstTitles_Click 'reset the list
    'make sure to clean up after ourselves
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub LoadCodeTypes()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    
    cmbType.Clear
    cmbFilter.Clear
    cmbFilter.AddItem "No Filter", 0 'no filter isnt in the db, so add it here so its on top
    
    'connect to the database and retrieve the valid code types
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "SELECT codetype FROM codetypes"
    Set adoRS = adoCmd.Execute
    
    'loop through the recordset and add them to the drop down
    Do While Not adoRS.EOF
        cmbType.AddItem CStr(adoRS("codetype"))
        cmbFilter.AddItem CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    
    'cleaning up the house
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing

    'reset lists to top item
    cmbFilter.ListIndex = 0
    cmbType.ListIndex = 0
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Function VerifyCode() As Boolean

    VerifyCode = True
    If txtTitle.Text = "" Then
        MsgBox "Please enter in a Title for the sourcecode snippit.", vbOKOnly, "Sourcecode Notebook"
        VerifyCode = False
        Exit Function
    End If
    If rtbCodeWindow.Text = "" Then
        MsgBox "You must enter in some Sourcecode to save a sourcecode snippit.", vbOKOnly, "Sourcecode Notebook"
        VerifyCode = False
        Exit Function
    End If

End Function

Private Sub Form_Resize()
    
    If Me.WindowState = vbNormal Then 'cant resize when min or maxed
        If Me.Height < 9000 Then Me.Height = 9000 'smallest window im allowing right now
        If Me.Width < 11000 Then Me.Width = 11000
        'this section was a major pain in the ........
        cmbFilter.Left = frmMain.Width - (cmbFilter.Width + 512)
        lblFilter.Left = cmbFilter.Left
        lstTitles.Width = frmMain.Width - (cmbFilter.Width + 580)
        lstTitles.Left = Screen.TwipsPerPixelX * 3
        frmSnippit.Left = lstTitles.Left
        frmSnippit.Width = frmMain.Width - (Screen.TwipsPerPixelX * 12)
        frmSnippit.Height = frmMain.Height - lstTitles.Height - 1200 - statBar.Height
        tabSnippit.Left = frmSnippit.Left + (Screen.TwipsPerPixelX * 3)
        tabSnippit.Width = frmSnippit.Width - (Screen.TwipsPerPixelX * 12)
        tabSnippit.Height = frmSnippit.Height - (Screen.TwipsPerPixelX * 55)
        txtTitle.Left = frmSnippit.Left
        cmbType.Left = frmMain.Width - (cmbType.Width + 512)
        txtTitle.Width = frmMain.Width - (cmbType.Width + 580)
        rtbCodeWindow.Left = tabSnippit.Left - (Screen.TwipsPerPixelX * 3)
        rtbCodeWindow.Width = tabSnippit.Width - (Screen.TwipsPerPixelX * 6)
        rtbCodeWindow.Height = tabSnippit.Height - (Screen.TwipsPerPixelX * 6) - tabSnippit.TabHeight
        rtbNotes.Left = tabSnippit.Left - (Screen.TwipsPerPixelX * 3)
        rtbNotes.Width = tabSnippit.Width - (Screen.TwipsPerPixelX * 6)
        rtbNotes.Height = tabSnippit.Height - (Screen.TwipsPerPixelX * 6) - tabSnippit.TabHeight
    End If
    
End Sub

Private Sub lstTitles_Click()
    
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim index As Long
            
    If lstTitles.ListItems.Count < 1 Then 'if there is nothing in the list yet
        Exit Sub
    End If
    'connect to the database and retrieve the selected items details
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    'here is one place where the tag comes in handy, selecting by title was
    'not the best idea, it would slow things down with large amount of snippits
    adoCmd.CommandText = "SELECT * FROM source WHERE id = " & lstTitles.SelectedItem.Tag
    Set adoRS = adoCmd.Execute
    
    'set up the code and notes windows, etc...
    rtbCodeWindow.Text = adoRS("code")
    txtTitle.Text = adoRS("title")
    rtbNotes.Text = adoRS("notes")
    'find the right code type in the drop down
    For index = 0 To cmbType.ListCount
        If Trim(cmbType.List(index)) = Trim(adoRS("codetype")) Then
            cmbType.ListIndex = index
            Exit For
        End If
    Next index
    
    'nope this aint a new piece of code
    gblNewCode = False
    
    'cleaning house
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub lstTitles_ColumnClick(ByVal ColumnHeader As MSComCtlLib.ColumnHeader)

    'some code I foundon PSC to easily sort the listview, kudos to whover posted this
    If lstTitles.SortKey <> ColumnHeader.index - 1 Then
        lstTitles.SortKey = ColumnHeader.index - 1
        lstTitles.SortOrder = lvwAscending
    Else
        lstTitles.SortOrder = IIf(lstTitles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If
    lstTitles.Sorted = True
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, frmMain  'plug, plug, plug, plug
End Sub

Private Sub mnuAdd_Click()
    
    On Error GoTo errHandler
    
    Dim retval As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    retval = InputBox("Please enter in the Code Type you wish to add:", "Add Code Type")
    If retval = "" Then
        Exit Sub
    End If
    
    'connect to the database and add in the new codetype
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "INSERT INTO codetypes (codetype) VALUES ('" & StuffQuotes(retval) & "')"
    adoCmd.Execute
            
    'clean up
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    LoadCodeTypes
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
        
End Sub

Private Sub mnuExit_Click()
    'seems kinda abrupt
    End
End Sub

Private Sub mnuFind_Click()
    Dim retval As String
    retval = InputBox("Find What?", "Find")
    If retval = "" Then
        Exit Sub
    End If
    Find (retval)
End Sub

Private Sub mnuModify_Click()
    frmCodeTypes.Show vbModal, frmMain
    'reload incase of changes
    LoadGridBox
    LoadCodeTypes
End Sub

Private Sub mnuPaste_Click()
    Clipboard.GetFormat vbCFText
    rtbCodeWindow.Text = Clipboard.GetText(vbCFText)
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText rtbCodeWindow.Text
End Sub

Private Sub mnuDelete_Click()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    'connnect to the database and delete the current selected snippit
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "DELETE FROM source WHERE id=" & lstTitles.SelectedItem.Tag
    adoCmd.Execute
            
    'cleanup the house
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    
    'reset the windows
    txtTitle.Text = ""
    rtbCodeWindow.Text = ""
    cmbType.ListIndex = 0
    LoadGridBox
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub mnuNew_Click()

    txtTitle.Text = ""
    cmbType.ListIndex = 0
    rtbCodeWindow.Text = ""
    rtbNotes.TextRTF = ""
    'hey it is new after all
    gblNewCode = True
    
End Sub

Private Sub mnuOpen_Click()

    Dim retval As String
    cdoOpenDatabase.ShowOpen
    retval = cdoOpenDatabase.FileName
    If retval = "" Then
        Exit Sub
    End If
    gblConnectString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & retval
    'this will get saved to the registry in a later version
    LoadGridBox  'reload the title list
    gblNewCode = True
    
End Sub

Private Sub mnuPlanetSourceCode_Click()
    'kudos to PSC
    Dim xRet As Long
    xRet = ShellExecute(0, vbNullString, "http://www.planet-source-code.com/PlanetSourceCode/", vbNullString, App.Path, 1)
End Sub

Private Sub mnuPrint_Click()
    If rtbCodeWindow.Text = "" Then 'oops
        MsgBox "There is nothing to print!", vbOKOnly, "Print"
        Exit Sub
    End If
    cdoOpenDatabase.ShowPrinter
    Screen.MousePointer = vbHourglass
    rtbCodeWindow.SelPrint (Printer.hDC)
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuSave_Click()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim retval As Boolean
    
    'connect to the database
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    If gblNewCode = False Then      'if we are working on an existing snippit
        retval = VerifyCode 'check to make sure the user dotted all i's and crossed all t's
        If retval = False Then
            Exit Sub
        End If
        'this really should be a stored procedure, but....
        adoCmd.CommandText = "UPDATE source SET title='" & StuffQuotes(txtTitle) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET code='" & StuffQuotes(rtbCodeWindow.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET codetype='" & StuffQuotes(cmbType.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET [datetime]='" & Now & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET notes='" & StuffQuotes(rtbNotes.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
    Else  'if its new
        retval = VerifyCode 'check to make sure the user dotted all i's and crossed all t's
        If retval = False Then
            Exit Sub
        End If
        adoCmd.CommandText = "INSERT INTO source ([datetime],title,codetype,code,notes) VALUES('" & Now & "', '" & StuffQuotes(txtTitle) & "', '" & StuffQuotes(cmbType.Text) & "', '" & StuffQuotes(rtbCodeWindow.Text) & "', '" & StuffQuotes(rtbNotes.TextRTF) & "')"
        adoCmd.Execute
        'we need the new identity created for it
        adoCmd.CommandText = "SELECT id FROM source WHERE title = '" & StuffQuotes(txtTitle) & "'"
        Set adoRS = adoCmd.Execute
        gblNewCode = False  'its no longer new
        adoRS.Close
        Set adoRS = Nothing
    End If
    
    'clean everything up
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    
    LoadGridBox  'reset the list
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub mnuWebsite_Click()
    'plug, plug, plug, plug
    Dim xRet As Long
    xRet = ShellExecute(0, vbNullString, "http://www.thejoyals.net/Sourcecode/", vbNullString, App.Path, 1)
End Sub

Private Sub rtbNotes_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

    'if anyone can tell me what the heck an effect of 3 is id appreciate it.
    'msdn documentation tells me there is no such thing
    If Data.GetFormat(vbCFText) Then 'if text
        If Effect = 3 Then
            rtbCodeWindow.Text = Data.GetData(vbCFText) 'set the window to the dragged in text
        Else
            rtbNotes.LoadFile Data.GetData(vbCFText), rtfText 'open the dragged in file
        End If
    End If
    If Data.GetFormat(vbCFFiles) Then 'if files from explorer
        rtbNotes.LoadFile Data.Files(1), rtfText  'open the file dragged from windows
    End If

End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)

    'no sense in recreating the wheel, so just call the menu item procedures
    If Button.Tag = "new" Then
        mnuNew_Click
    End If
    If Button.Tag = "delete" Then
        mnuDelete_Click
    End If
    If Button.Tag = "save" Then
        mnuSave_Click
    End If
    If Button.Tag = "paste" Then
        mnuPaste_Click
    End If
    If Button.Tag = "copy" Then
        mnuCopy_Click
    End If
    If Button.Tag = "open" Then
        mnuOpen_Click
    End If
    If Button.Tag = "find" Then
        mnuFind_Click
    End If
    If Button.Tag = "print" Then
        mnuPrint_Click
    End If
    
End Sub

Private Sub rtbCodeWindow_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

    'if anyone can tell me what the heck an effect of 3 is id appreciate it.
    'msdn documentation tells me there is no such thing
    If Data.GetFormat(vbCFText) Then  'if text
        If Effect = 3 Then
            rtbCodeWindow.Text = Data.GetData(vbCFText) 'set the window to the dragged in text
        Else
            rtbCodeWindow.LoadFile Data.GetData(vbCFText) 'open the dragged in file
        End If
    End If
    If Data.GetFormat(vbCFFiles) Then 'if files from explorer
        rtbCodeWindow.LoadFile Data.Files(1) 'open the file dragged from windows
    End If

End Sub

Private Sub Find(strSearch As String)
    
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim cmdtext As String
    
    'not the best find, but it does the job
    'this can get very slow with large numbers of snippits
    cmdtext = "SELECT title FROM source WHERE title like '%" & strSearch & "%'"
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = cmdtext
    Set adoRS = adoCmd.Execute
    
    'only take the first returned result, ignore any others
    If adoRS.EOF = False Then
        'find the item in the listview and select it
        lstTitles.SelectedItem = lstTitles.FindItem(adoRS("title"), , lvwPartial)
        lstTitles_Click 'load it into the code window
    Else
        'nothing matched
        MsgBox "Not Found", vbOKOnly, "Find"
    End If
    
    'clean up
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub
