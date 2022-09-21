VERSION 5.00
Begin VB.Form frmNewSubDept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " New Sub Department"
   ClientHeight    =   1860
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   5160
   Icon            =   "frmNewSubDept.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   3585
      Begin VB.ComboBox cboDept 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1170
         Width           =   3285
      End
      Begin VB.TextBox txtSubDept 
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
      Begin VB.Label Label2 
         Caption         =   "Under Department :"
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
         TabIndex        =   6
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Sub Department :"
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
         TabIndex        =   5
         Top             =   210
         Width           =   1575
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
      TabIndex        =   3
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
      TabIndex        =   2
      ToolTipText     =   "Create New Company"
      Top             =   180
      Width           =   1065
   End
End
Attribute VB_Name = "frmNewSubDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdOK.SetFocus
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdOK_Click()

If Len(txtSubDept) = 0 Then
  MsgBox "Enter name of sub department", vbInformation, cnstMsgInfo
  txtSubDept.SetFocus
  Exit Sub
End If
If Len(cboDept) = 0 Then
  MsgBox "Select under department", vbInformation, cnstMsgInfo
  cboDept.SetFocus
  Exit Sub
End If

Dim strSubDept As String, strDept As String, intDeptID As Integer

If strMark = "SE" Then
  With frmSubDept.flexDetails
    .Col = 0
    strSubDept = .Text
    .Col = 1
    strDept = .Text
    .Col = 0
  End With
  If txtSubDept.Text = strSubDept And cboDept.Text = strDept Then
    Unload Me
    Exit Sub
  End If
End If
intDeptID = findDeptID(cboDept.Text)

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

If strMark <> "SE" Or txtSubDept.Text <> strSubDept Then
''  strRecord = "SELECT SubDept FROM new_Sub_Dept WHERE ((SubDept='" & txtSubDept.Text & "') AND (DeptID=" & intDeptID & ") AND (CompID=" & bytCompID & "));"
''  R.Open strRecord, MainConn, adOpenStatic
''  If R.RecordCount > 0 Then
''    MsgBox "Sub Department '" & txtSubDept.Text & "' already exist under selected department." & vbCrLf & "Change the sub department name.", vbInformation, cnstMsgInfo
''    Screen.MousePointer = vbDefault
''    txtSubDept.SetFocus
''    Exit Sub
''  End If
''  R.Close
End If

If strMark <> "SE" Then
  lngLastID = 1
  strRecord = "SELECT Max(SubDeptID) AS LastID FROM new_Sub_Dept HAVING Max(SubDeptID) Is Not Null;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then lngLastID = lngLastID + R![LastID]
  R.Close
Else 'If strMark = "SE" Then
  strRecord = "SELECT new_Sub_Dept.SubDeptID FROM new_Dept_Name INNER JOIN new_Sub_Dept ON new_Dept_Name.DeptID = new_Sub_Dept.DeptID" _
            & " Where (((new_Sub_Dept.SubDept)='" & strSubDept & "') And ((new_Dept_Name.DeptName)='" & strDept & "')" _
            & " And ((new_Sub_Dept.CompID)=" & bytCompID & ")) GROUP BY new_Sub_Dept.SubDeptID;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then lngLastID = R![SubDeptID]
  R.Close
End If

MainConn.BeginTrans
  
If strMark <> "SE" Then
  MainComm.CommandText = "INSERT INTO new_Sub_Dept(SubDeptID,SubDept,DeptID,CompID)VALUES(" & lngLastID & ",'" & txtSubDept.Text & "'," & intDeptID & "," & bytCompID & ");"
  MainComm.Execute
Else 'If strMark = "SE" Then
  MainComm.CommandText = "UPDATE new_Sub_Dept SET SubDept='" & txtSubDept.Text & "',DeptID=" & intDeptID & " WHERE SubDeptID=" & lngLastID & ";"
  MainComm.Execute
End If

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  If Err.Number = -2147217873 Then
    MsgBox "Duplicate name of Sub Department", vbCritical, cnstMsgInfo
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

If strMark = "SN" Then
  With frmSubDept.flexDetails
    .Rows = .Rows + 1
    .Row = .Rows - 1
    .Col = 0
    .Text = txtSubDept.Text
    .Col = 1
    .Text = cboDept.Text
    
    .Row = 0
    For intI = 1 To .Rows
      If intI Mod 2 = 0 Then
        .Col = 0
        .CellBackColor = &HF0F1F2
        .Col = 1
        .CellBackColor = &HF0F1F2
      Else
        .Col = 0
        .CellBackColor = &H80000005
        .Col = 1
        .CellBackColor = &H80000005
      End If
      .Col = 0
      If .Row < (.Rows - 1) Then .Row = .Row + 1
    Next intI
    If .Rows = 1 And boUserEdit = True Then
      frmSubDept.cmdEdit.Enabled = True
      frmSubDept.cmdDelete.Enabled = True
    End If
  End With
  frmSubDept.txtTotal.Text = CLng(frmSubDept.txtTotal.Text) + 1
  txtSubDept.SetFocus
ElseIf strMark = "SE" Then
  With frmSubDept.flexDetails
    .Col = 0
    .Text = txtSubDept.Text
    .Col = 1
    .Text = cboDept.Text
    .Col = 0
  End With
  Unload Me
ElseIf strMark = "SDEPT" Then
  Call addSubDept(frmNewEmpInfo.cboSubDept, CInt(frmNewEmpInfo.cboDept.Tag))
  frmNewEmpInfo.cboSubDept.Text = txtSubDept.Text
  Unload Me
  frmNewEmpInfo.cboSubDept.SetFocus
End If
txtSubDept.Text = ""
End Sub
Private Sub Form_Load()
  addDeptName cboDept
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmNewSubDept = Nothing
End Sub
Private Sub txtSubDept_GotFocus()
  txtGotFocus txtSubDept
End Sub
Private Sub txtSubDept_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cboDept.SetFocus
End Sub
Private Sub txtSubDept_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtSubDept_LostFocus()
  txtSubDept.Text = Trim(txtSubDept.Text)
End Sub
