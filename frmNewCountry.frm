VERSION 5.00
Begin VB.Form frmNewCountry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3480
   ClientLeft      =   4275
   ClientTop       =   2895
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNewCountry.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5835
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
      Left            =   4560
      TabIndex        =   8
      ToolTipText     =   "Edit/Delete Record"
      Top             =   2805
      Width           =   1185
   End
   Begin VB.TextBox txtEditCountry 
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
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cboCountry 
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
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1155
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   1950
      TabIndex        =   3
      ToolTipText     =   "Edit/Delete Record"
      Top             =   2805
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
      Left            =   3255
      TabIndex        =   2
      ToolTipText     =   "Save/Update Record"
      Top             =   2805
      Width           =   1185
   End
   Begin VB.TextBox txtCountry 
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
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1770
      Width           =   3135
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   5835
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   0
      TabIndex        =   9
      Top             =   2595
      Width           =   5895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country Information"
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
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2700
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
      Left            =   -240
      TabIndex        =   7
      Top             =   855
      Width           =   6135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name of Country  :"
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
      Height          =   210
      Left            =   1440
      TabIndex        =   5
      Top             =   1500
      Width           =   1545
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmNewCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCountry_Click()
  If Len(cboCountry) = 0 Then Exit Sub
  txtEditCountry.Text = cboCountry.Text
End Sub
Private Sub cboCountry_LostFocus()
  cboCountry.Text = Trim(UCase(cboCountry.Text))
  If Len(cboCountry) > 30 Then
    MsgBox "Maximum size is '30'", vbInformation, cnstMsgInfo
    cboCountry.SetFocus
  End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()

If cmdEdit.Caption = "&Delete" Then

  If Len(cboCountry) = 0 Then
    MsgBox "Select name of Country", vbInformation, cnstMsgInfo
    cboCountry.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  On Error GoTo ErrorHandler
  
  Set MainComm = New ADODB.Command
  MainComm.ActiveConnection = MainConn
  
  MainComm.CommandText = "DELETE * FROM new_Country WHERE CountryName='" & cboCountry.Text & "';"
  MainComm.Execute
  
  addCountry cboCountry
  cboCountry.SetFocus
Else
  txtCountry.Visible = False
  cboCountry.Top = txtCountry.Top
  cboCountry.Visible = True
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

  If Len(txtCountry) = 0 Then
    MsgBox "Type name of Country", vbInformation, cnstMsgInfo
    txtCountry.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  MainConn.BeginTrans
    
  MainComm.CommandText = "INSERT INTO new_Country VALUES('" & txtCountry.Text & "');"
  MainComm.Execute

  MainConn.CommitTrans
  
  txtCountry.Text = ""
  txtCountry.SetFocus

Else

  If Len(cboCountry) = 0 Then
    MsgBox "Select name of Country", vbInformation, cnstMsgInfo
    cboCountry.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

  If cboCountry.Text <> txtEditCountry.Text Then
    strRecord = "SELECT CountryName FROM new_Country WHERE CountryName='" & cboCountry.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      MsgBox "Duplicate name of Country", vbInformation, cnstMsgInfo
      cboCountry.SetFocus
      Exit Sub
    End If
    R.Close
  End If
  
  MainConn.BeginTrans
  
  MainComm.CommandText = "UPDATE new_Country SET CountryName='" & cboCountry.Text & "' WHERE CountryName='" & txtEditCountry.Text & "';"
  MainComm.Execute

  MainConn.CommitTrans
  
  cboCountry.SetFocus

End If
addCountry cboCountry
End Sub
Private Sub Form_Load()
  addCountry cboCountry
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmNewCountry = Nothing
End Sub

Private Sub txtCountry_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtCountry_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtCountry_LostFocus()
  If Len(txtCountry) = 0 Then Exit Sub
  txtCountry.Text = Trim(UCase(txtCountry.Text))
  Set R = New ADODB.Recordset
  strRecord = "SELECT CountryName FROM new_Country WHERE CountryName='" & txtCountry.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    MsgBox "Duplicate name of Country", vbInformation, cnstMsgInfo
    txtCountry.SetFocus
  End If
  R.Close
  Set R = Nothing
End Sub
