VERSION 5.00
Begin VB.Form frmNewPFNo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   4260
   ClientTop       =   2880
   ClientWidth     =   5805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNewPFNo.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4500
      TabIndex        =   11
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3195
      TabIndex        =   10
      ToolTipText     =   "Save/Update Record"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1905
      TabIndex        =   9
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.TextBox txtEditPFNo 
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
      Height          =   315
      Left            =   1320
      MaxLength       =   100
      TabIndex        =   2
      Top             =   2460
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cboPFNo 
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
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtPFNo 
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
      Height          =   315
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2130
      Width           =   3135
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   -15
      TabIndex        =   8
      Top             =   3405
      Width           =   5835
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -15
      TabIndex        =   7
      Top             =   3360
      Width           =   6075
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
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   855
      Width           =   6135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create File No."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1890
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File  No.  :"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1830
      Width           =   1305
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmNewPFNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboPFNo_Click()
  If Len(cboPFNo) = 0 Then Exit Sub
  txtEditPFNo.Text = cboPFNo.Text
End Sub
Private Sub cboPFNo_LostFocus()
  cboPFNo.Text = Trim(UCase(cboPFNo.Text))
  If Len(cboPFNo) > 20 Then
    MsgBox "Maximum size is '20'", vbInformation, cnstMsgInfo
    cboPFNo.SetFocus
  End If
End Sub

Private Sub cmdClose_Click()
    Set frmNewPFNo = Nothing
    Unload Me
End Sub

Private Sub cmdEdit_Click()

If cmdEdit.Caption = "&Delete" Then

  If Len(cboPFNo) = 0 Then
    MsgBox "Select PF No.", vbInformation, cnstMsgInfo
    cboPFNo.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  On Error GoTo ErrorHandler
  
  Set MainComm = New ADODB.Command
  MainComm.ActiveConnection = MainConn
  
  MainComm.CommandText = "DELETE  FROM new_PF_No WHERE PFNo='" & cboPFNo.Text & "';"
  MainComm.Execute
  
  addPFNo cboPFNo
  cboPFNo.SetFocus
Else
  txtPFNo.Visible = False
  cboPFNo.Top = txtPFNo.Top
  cboPFNo.Visible = True
  cmdSave.Caption = "&Update"
  cmdEdit.Caption = "&Delete"
  Exit Sub
End If

ErrorHandler:
If err.Number <> 0 Then
  If err.Number = -2147467259 Then
    MsgBox "This record cannot be deleted", vbCritical, cnstMsgErDB
  Else
    MsgBox err.Description, vbCritical, cnstMsgErDB
  End If
End If

End Sub
Private Sub cmdSave_Click()

Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

If cmdSave.Caption = "&Save" Then
  If Len(txtPFNo) = 0 Then
    MsgBox "Type PF No.", vbInformation, cnstMsgInfo
    txtPFNo.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  MainConn.BeginTrans
    
  MainComm.CommandText = "INSERT INTO new_PF_No VALUES('" & txtPFNo.Text & "');"
  MainComm.Execute

  MainConn.CommitTrans
  
  txtPFNo.Text = ""
  txtPFNo.SetFocus
Else
  If Len(cboPFNo) = 0 Then
    MsgBox "Select PF No.", vbInformation, cnstMsgInfo
    cboPFNo.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

  If cboPFNo.Text <> txtEditPFNo.Text Then
    strRecord = "SELECT PFNo FROM new_PF_No WHERE PFNo='" & cboPFNo.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      MsgBox "Duplicate PF No", vbInformation, cnstMsgInfo
      cboPFNo.SetFocus
      Exit Sub
    End If
    R.Close
  End If
  MainConn.BeginTrans
  
  MainComm.CommandText = "UPDATE new_PF_No SET PFNo='" & cboPFNo.Text & "' WHERE PFNo='" & txtEditPFNo.Text & "';"
  MainComm.Execute

  MainConn.CommitTrans
  cboPFNo.SetFocus
End If
addPFNo cboPFNo
End Sub
Private Sub Form_Load()
  addPFNo cboPFNo
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmNewPFNo = Nothing
End Sub
Private Sub txtPFNo_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtPFNo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtPFNo_LostFocus()
  If Len(txtPFNo) = 0 Then Exit Sub
  txtPFNo.Text = Trim(UCase(txtPFNo.Text))
  Set R = New ADODB.Recordset
  strRecord = "SELECT PFNo FROM new_PF_No WHERE PFNo='" & txtPFNo.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    MsgBox "Duplicate PF No.", vbInformation, cnstMsgInfo
    txtPFNo.SetFocus
  End If
  R.Close
  Set R = Nothing
End Sub
