VERSION 5.00
Begin VB.Form frmCodeTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify / Delete Code Types"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmCodeTypes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   435
      Left            =   1140
      TabIndex        =   6
      Top             =   3780
      Width           =   2295
   End
   Begin VB.Frame frameDelete 
      Caption         =   "Delete"
      Height          =   1665
      Left            =   120
      TabIndex        =   3
      Top             =   2010
      Width           =   4395
      Begin VB.ListBox lstDelete 
         Height          =   1035
         Left            =   750
         TabIndex        =   5
         Top             =   540
         Width           =   2955
      End
      Begin VB.Label lblDelete 
         Caption         =   "Select the Code Type you wish to delete:"
         Height          =   225
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   2925
      End
   End
   Begin VB.Frame frameModify 
      Caption         =   "Modify"
      Height          =   1665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      Begin VB.ListBox lstModify 
         Height          =   1035
         Left            =   660
         TabIndex        =   1
         Top             =   540
         Width           =   2955
      End
      Begin VB.Label lblModify 
         Caption         =   "Select the Code Type you wish to modify:"
         Height          =   225
         Left            =   660
         TabIndex        =   2
         Top             =   240
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmCodeTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset

    lstModify.Clear
    lstDelete.Clear
    'create the connection and execute the SQL
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "SELECT * FROM codetypes"
    Set adoRS = adoCmd.Execute
            
    Do While Not adoRS.EOF
        'add the recordsetset items into the listboxes
        lstModify.AddItem CStr(adoRS("codetype"))
        lstDelete.AddItem CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    
    'make sure we clean up!
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub lstDelete_Click()
    
    On Error GoTo errHandler
    
    Dim retval As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    'only 1 chance to say no!
    retval = MsgBox("Are you sure you wish to do this?  This will change all snippits that have this code type, to a <blank> code type", vbYesNo, "Delete Code Type")
    If retval = vbNo Then
        Exit Sub
    End If
    
    'connect to the database and delete the code type, the reset the source entries
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "DELETE FROM codetypes WHERE codetype='" & StuffQuotes(lstDelete.Text) & "'"
    adoCmd.Execute
    adoCmd.CommandText = "UPDATE source SET codetype='<blank>' WHERE codetype='" & StuffQuotes(lstDelete.Text) & "'"
    adoCmd.Execute
    
    'make sure we clean up after ourselves
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Form_Load  'reset everything
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub lstModify_Click()

    On Error GoTo errHandler
    
    Dim retval As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    retval = InputBox("Please enter in the new title for the Code Type", "Modify Code Type", CStr(lstModify.Text))
    If retval = "" Then
        Exit Sub
    End If
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "UPDATE codetypes SET codetype='" & retval & "' WHERE codetype='" & StuffQuotes(lstModify.Text) & "'"
    adoCmd.Execute
    adoCmd.CommandText = "UPDATE source SET codetype='" & retval & "' WHERE codetype='" & StuffQuotes(lstModify.Text) & "'"
    adoCmd.Execute
    
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Form_Load
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub
