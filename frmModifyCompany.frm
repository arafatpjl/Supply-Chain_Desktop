VERSION 5.00
Begin VB.Form frmModifyCompany 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   2685
   ClientTop       =   2865
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmModifyCompany.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTINNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   1350
      MaxLength       =   30
      TabIndex        =   2
      Top             =   2955
      Width           =   2730
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "Clo&se"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   915
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   4
      Top             =   3780
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2175
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   3
      Top             =   3795
      Width           =   945
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   1095
      Left            =   180
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1785
      Width           =   3885
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   180
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1095
      Width           =   3855
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   525
      Index           =   0
      Left            =   -120
      TabIndex        =   13
      Top             =   3720
      Width           =   4305
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   -120
      TabIndex        =   12
      Top             =   3690
      Width           =   4290
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   525
      Width           =   4290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Company"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   75
      Width           =   1875
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   8
      Top             =   2985
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TIN No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   2985
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   1545
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   855
      Width           =   1455
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   525
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4305
   End
End
Attribute VB_Name = "frmModifyCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdSave_Click()

If Len(txtName) = 0 Then
  MsgBox "Enter Name of Company", vbInformation, cnstMsgInfo
  txtName.SetFocus
  Exit Sub
End If
If Len(txtAddress) = 0 Then
  MsgBox "Enter Company Address", vbInformation, cnstMsgInfo
  txtAddress.SetFocus
  Exit Sub
End If

If MsgBox("Are You sure to Update", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

MainComm.CommandText = "UPDATE Company_Information SET compName='" & txtName.Text & "',compAdd='" & txtAddress.Text & "',CompTIN='" & txtTINNo.Text & "' WHERE compID=" & bytCompID & ";"
MainComm.Execute

MainConn.CommitTrans

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  err.Number = 0
  Exit Sub
End If

MsgBox "Successfully Updated", vbInformation, cnstMsgInfo

frmMDIMain.Caption = txtName.Text
Unload Me

End Sub
Private Sub Form_Load()
  Set R = New ADODB.Recordset
  strRecord = "SELECT compName,compAdd,CompTIN FROM Company_Information WHERE compID=" & bytCompID & ";"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    txtName.Text = R![CompName]
    txtAddress.Text = R![CompAdd]
    If Not IsNull(R![CompTIN]) Then txtTINNo.Text = R![CompTIN]
  End If
  R.Close
  Set R = Nothing
  If boUserEdit = False Then cmdSave.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmModifyCompany = Nothing
End Sub

Private Sub Label5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, 161, 2, 0
End Sub

Private Sub txtAddress_GotFocus()
  'SendKeys "^{Home}+^{End}"
  txtGotFocus txtAddress
End Sub
Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtTINNo.SetFocus
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13 Or KeyAscii = 39) Then KeyAscii = 0
End Sub
Private Sub txtAddress_LostFocus()
  txtAddress.Text = Trim(txtAddress.Text)
End Sub
Private Sub txtName_GotFocus()
  txtGotFocus txtName
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtAddress.SetFocus
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtName_LostFocus()
  txtName.Text = Trim(txtName.Text)
End Sub
Private Sub txtTINNo_GotFocus()
  txtGotFocus txtTINNo
End Sub
Private Sub txtTINNo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then cmdSave.SetFocus
End Sub
Private Sub txtTINNo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtTINNo_LostFocus()
  If Len(txtTINNo) = 0 Then Exit Sub
  txtTINNo.Text = Trim(txtTINNo.Text)
End Sub
