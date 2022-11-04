VERSION 5.00
Begin VB.Form frmUserLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4290
   ClientLeft      =   3405
   ClientTop       =   3015
   ClientWidth     =   5835
   ControlBox      =   0   'False
   Icon            =   "frmUserLogIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2534.673
   ScaleMode       =   0  'User
   ScaleWidth      =   5478.749
   Begin VB.CommandButton cmdModPass 
      BackColor       =   &H00C0C000&
      Caption         =   "&Change Password"
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
      Left            =   255
      TabIndex        =   13
      Top             =   3615
      Width           =   1185
   End
   Begin VB.TextBox txtUserName 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1350
      TabIndex        =   0
      Top             =   1515
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Login"
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
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3630
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3150
      TabIndex        =   3
      Top             =   3630
      Width           =   1185
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1350
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2460
      Width           =   3015
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   3465
      Width           =   5835
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -150
      TabIndex        =   10
      Top             =   3420
      Width           =   6015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login"
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
      TabIndex        =   9
      Top             =   240
      Width           =   1395
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
      TabIndex        =   8
      Top             =   840
      Width           =   5850
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
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   0
      Left            =   2340
      TabIndex        =   7
      Top             =   1275
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   1350
      TabIndex        =   6
      Top             =   1275
      Width           =   1035
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
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   3
      Left            =   2340
      TabIndex        =   5
      Top             =   2190
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Index           =   0
      Left            =   1350
      TabIndex        =   4
      Top             =   2190
      Width           =   1035
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdModPass_Click()
    frmCreateModifyUser.Show 1
End Sub

Private Sub Form_GotFocus()
    txtUserName.SetFocus
End Sub

Private Sub Label5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, 161, 2, 0
End Sub

Private Sub txtPassword_GotFocus()
    Call txtGotFocus(txtPassword)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.Text = Trim(txtPassword.Text)
End Sub

Private Sub txtUserName_GotFocus()
    Call txtGotFocus(txtUserName)
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
    If Len(txtUserName) = 0 Then
        MsgBox "Type User Name", vbInformation, cnstMsgInfo
        txtUserName.SetFocus
    Exit Sub
    End If
    If Len(txtPassword) = 0 Then
        MsgBox "Type Password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
        Exit Sub
    End If
    Call ConnectDatabase
    Set R = New ADODB.Recordset
        strRecord = "SELECT UserID,UserStatus FROM InventoryPJL.dbo.Sys_User_Name WHERE (UserName='" & txtUserName.Text & "') AND (UserPWord='" & EncryptIt(txtPassword.Text, 11) & "') "
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount = 0 Then
        MsgBox "Incorrect User Name or Password", vbInformation, "Login Error"
        txtUserName.SetFocus
        Exit Sub
    Else
    bytUserID = R!Userid
    strUserStatus = R!UserStatus
    strUser = txtUserName.Text
    
'    Call prcMenuVisibleLogin
    End If
    R.Close
    ''-======================
'    Set RS = New ADODB.Recordset
'                strRecord = "SELECT Sys_Default_Permission.PermName FROM Sys_User_Permission INNER JOIN Sys_Default_Permission" _
'                    & " ON Sys_User_Permission.PermID = Sys_Default_Permission.PermID WHERE Sys_User_Permission.UserID=" & bytUserID & ";"
'                R.Open strRecord, MainConn, adOpenStatic
'                If R.RecordCount > 0 Then
'                    boUserEdit = False
'                    Do Until R.EOF
'                        If UCase(R![PermName]) = "DATA ENTRY" Then frmMDIMain.mnuDataEntry.Visible = True
'                        If UCase(R![PermName]) = "DATA EDIT" Then boUserEdit = True
'                        If UCase(R![PermName]) = "VIEW REPORT" Then frmMDIMain.mnuReport.Visible = True
'                    R.MoveNext
'                    Loop
'                End If
'                R.Close
    If strUserStatus <> "A" Then
        Call prcMenuInvisibleAll
        Call prcMenuVisibleCompany
    End If
    frmMDIMain.Show
    frmMDIMain.lblCompName = ""
    frmMDIMain.lblCompName = strComNameTxt
    ''=======================
'            frmMDIMain.mnuHouseKeeping.Visible = True
'            frmMDIMain.mnuDataEntry.Visible = True
'            frmMDIMain.mnuReport.Visible = True
'            frmMDIMain.mnuOption.Visible = True
'            frmMDIMain.mnuCloseComp.Visible = True
'            frmMDIMain.mnuNextExit.Visible = True

   
    
    Set R = Nothing
    
    Unload Me
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUserLogin = Nothing
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtUserName_LostFocus()
    txtUserName.Text = Trim(txtUserName.Text)
End Sub
