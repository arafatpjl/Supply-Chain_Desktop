VERSION 5.00
Begin VB.Form frmNewReligion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " New Religion"
   ClientHeight    =   1230
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   5160
   Icon            =   "frmNewReligion.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   3585
      Begin VB.TextBox txtRelig 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   150
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label Label1 
         Caption         =   "Religion :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   2
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   1
      ToolTipText     =   "Create New Company"
      Top             =   180
      Width           =   1065
   End
End
Attribute VB_Name = "frmNewReligion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdOK_Click()

If Len(txtRelig) = 0 Then
  MsgBox "Enter name of Religion", vbInformation, cnstMsgInfo
  txtRelig.SetFocus
  Exit Sub
End If

Dim strRelig As String

If strMark = "RE" Then
  strRelig = frmReligion.flexDetails.Text
  If txtRelig.Text = strRelig Then
    Unload Me
    Exit Sub
  End If
End If

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

If strMark <> "RE" Or txtRelig.Text <> strRelig Then
''  strRecord = "SELECT Religion FROM new_Religion WHERE ((Religion='" & txtRelig.Text & "') AND (CompID=" & bytCompID & "));"
''  R.Open strRecord, MainConn, adOpenStatic
''  If R.RecordCount > 0 Then
''    MsgBox "Religion '" & txtRelig.Text & "' already exist." & vbCrLf & "Change the religion name.", vbInformation, cnstMsgInfo
''    Screen.MousePointer = vbDefault
''    txtRelig.SetFocus
''    Exit Sub
''  End If
''  R.Close
End If

If strMark <> "RE" Then
  lngLastID = 1
  strRecord = "SELECT Max(ReligID) AS LastID FROM new_Religion HAVING Max(ReligID) Is Not Null;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then lngLastID = lngLastID + R![LastID]
  R.Close
Else 'If strMark = "RE" Then
  strRecord = "SELECT ReligID FROM new_Religion WHERE ((Religion='" & strRelig & "') AND (CompID=" & bytCompID & "));"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then lngLastID = R![ReligID]
  R.Close
End If

MainConn.BeginTrans
  
If strMark <> "RE" Then
  MainComm.CommandText = "INSERT INTO new_Religion(ReligID,Religion,CompID)VALUES(" & lngLastID & ",'" & txtRelig.Text & "'," & bytCompID & ");"
  MainComm.Execute
Else 'If strMark = "RE" Then
  MainComm.CommandText = "UPDATE new_Religion SET Religion='" & txtRelig.Text & "' WHERE ReligID=" & lngLastID & ";"
  MainComm.Execute
End If

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  If Err.Number = -2147217873 Then
    MsgBox "Duplicate name of Religion", vbCritical, cnstMsgInfo
  Else
    MsgBox Err.Description, vbCritical, cnstMsgErDB
  End If
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  Err.Number = 0
  Exit Sub
End If

Set R = Nothing
Screen.MousePointer = vbDefault

If strMark = "RN" Then
  With frmReligion.flexDetails
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .Col = 0
    .Text = txtRelig.Text
    
    .Row = 0
    For intI = 1 To .Rows
      If intI Mod 2 = 0 Then
        .CellBackColor = &HF0F1F2
      Else
        .CellBackColor = &H80000005
      End If
      If .Row < (.Rows - 1) Then .Row = .Row + 1
    Next intI
    If .Rows = 1 And boUserEdit = True Then
      frmReligion.cmdEdit.Enabled = True
      frmReligion.cmdDelete.Enabled = True
    End If
  End With
  frmReligion.txtTotal.Text = CLng(frmReligion.txtTotal.Text) + 1
  txtRelig.SetFocus
ElseIf strMark = "RE" Then
  With frmReligion.flexDetails
    .Col = 0
    .Text = txtRelig.Text
  End With
  Unload Me
ElseIf strMark = "RELIG" Then
  Call addReligion(frmNewEmpInfo.cboReligion)
  frmNewEmpInfo.cboReligion.Text = txtRelig.Text
  Unload Me
  frmNewEmpInfo.cboReligion.SetFocus
End If
txtRelig.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmNewReligion = Nothing
End Sub
Private Sub txtRelig_GotFocus()
  txtGotFocus txtRelig
End Sub
Private Sub txtRelig_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdOK.SetFocus
End Sub
Private Sub txtRelig_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtRelig_LostFocus()
  txtRelig.Text = Trim(txtRelig.Text)
End Sub
