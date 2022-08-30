VERSION 5.00
Begin VB.Form frmCreateUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7695
   ClientLeft      =   2205
   ClientTop       =   2040
   ClientWidth     =   7635
   ControlBox      =   0   'False
   Icon            =   "frmCreateUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Working Company"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   1485
      Left            =   825
      TabIndex        =   10
      Top             =   2925
      Width           =   5835
      Begin VB.ListBox List3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         ItemData        =   "frmCreateUser.frx":0442
         Left            =   150
         List            =   "frmCreateUser.frx":0444
         TabIndex        =   32
         Top             =   510
         Width           =   2475
      End
      Begin VB.CommandButton cmdRemoveC 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   31
         ToolTipText     =   "Remove "
         Top             =   900
         Width           =   420
      End
      Begin VB.CommandButton cmdAddC 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   30
         ToolTipText     =   "Add"
         Top             =   570
         Width           =   420
      End
      Begin VB.ListBox List4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         ItemData        =   "frmCreateUser.frx":0446
         Left            =   3210
         List            =   "frmCreateUser.frx":0448
         TabIndex        =   29
         Top             =   510
         Width           =   2475
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitted Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   255
         Index           =   0
         Left            =   3210
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
      Enabled         =   0   'False
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
      Left            =   6180
      TabIndex        =   9
      Top             =   7020
      Width           =   1185
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
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
      Left            =   2625
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1275
      Width           =   2805
   End
   Begin VB.ComboBox cboUser 
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
      ItemData        =   "frmCreateUser.frx":044A
      Left            =   2625
      List            =   "frmCreateUser.frx":044C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1005
      Visible         =   0   'False
      Width           =   2805
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
      Left            =   3615
      TabIndex        =   7
      Top             =   7020
      Width           =   1185
   End
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
      Left            =   4905
      TabIndex        =   8
      Top             =   7020
      Width           =   1185
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00C0C000&
      Caption         =   "&Modify"
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
      Left            =   2325
      TabIndex        =   2
      Top             =   7020
      Width           =   1185
   End
   Begin VB.TextBox txtVPassword 
      Alignment       =   2  'Center
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
      Left            =   2625
      MaxLength       =   15
      TabIndex        =   4
      Top             =   2055
      Width           =   2805
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
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
      Left            =   2625
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1665
      Width           =   2805
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   2625
      TabIndex        =   20
      Top             =   2415
      Width           =   2805
      Begin VB.OptionButton optAdmin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Administrator"
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
         Left            =   1140
         TabIndex        =   6
         Top             =   180
         Width           =   1485
      End
      Begin VB.OptionButton optUser 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User"
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
         Top             =   180
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Working Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   2265
      Left            =   825
      TabIndex        =   11
      Top             =   4395
      Width           =   5835
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1425
         ItemData        =   "frmCreateUser.frx":044E
         Left            =   3210
         List            =   "frmCreateUser.frx":0450
         TabIndex        =   17
         Top             =   510
         Width           =   2475
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   16
         ToolTipText     =   "Add"
         Top             =   600
         Width           =   420
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   15
         ToolTipText     =   "Add All"
         Top             =   930
         Width           =   420
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   14
         ToolTipText     =   "Remove "
         Top             =   1380
         Width           =   420
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "<>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   13
         ToolTipText     =   "Remove All"
         Top             =   1710
         Width           =   420
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1425
         ItemData        =   "frmCreateUser.frx":0452
         Left            =   150
         List            =   "frmCreateUser.frx":0454
         TabIndex        =   12
         Top             =   510
         Width           =   2475
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Permitted Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   3210
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   0
      TabIndex        =   39
      Top             =   6795
      Width           =   7755
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   15
      TabIndex        =   38
      Top             =   6840
      Width           =   7635
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create or Modify User"
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
      TabIndex        =   36
      Top             =   225
      Width           =   2820
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   35
      Top             =   840
      Width           =   7875
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   2385
      TabIndex        =   28
      Top             =   2565
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Category"
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
      Left            =   855
      TabIndex        =   27
      Top             =   2565
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   855
      TabIndex        =   26
      Top             =   1275
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   2
      Left            =   2385
      TabIndex        =   25
      Top             =   1275
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Verify Password"
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
      Left            =   855
      TabIndex        =   24
      Top             =   2055
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Password"
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
      Left            =   855
      TabIndex        =   23
      Top             =   1665
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   2385
      TabIndex        =   22
      Top             =   2055
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   3
      Left            =   2385
      TabIndex        =   21
      Top             =   1665
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   7635
   End
End
Attribute VB_Name = "frmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboUser_Click()
    If Len(cboUser) = 0 Then Exit Sub

    Set R = New ADODB.Recordset
    strRecord = "SELECT UserID,UserPWord,UserStatus from InventoryPJL.dbo.Sys_User_Name WHERE UserName='" & EncryptIt(cboUser.Text, 11) & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        txtUser.Text = EncryptIt(cboUser.Text, 11)
        bytUserID = R![Userid]
        txtPassword.Text = DecryptIt(R![UserPWord], 11)
        txtVPassword.Text = DecryptIt(R![UserPWord], 11)
        If DecryptIt(R![UserStatus], 11) = "A" Then
            optAdmin.Value = True
            cboUser.Tag = "A"
        Else
            optUser.Value = True
            cboUser.Tag = "U"
        End If
    End If
    R.Close

    List4.Clear
    List2.Clear

    If optAdmin.Value = True Then
        List4.Clear
        List3.ListIndex = -1
        
        For inti = 1 To List3.ListCount
            List3.ListIndex = List3.ListIndex + 1
            List4.AddItem List3.Text
        Next inti

        List1.ListIndex = -1
        For inti = 1 To List1.ListCount
            List1.ListIndex = List1.ListIndex + 1
            List2.AddItem List1.Text
        Next inti
    Else
        strRecord = "SELECT Company_Information.compName FROM Sys_User_Company INNER JOIN Company_Information" _
            & " ON Sys_User_Company.CompID = Company_Information.compID WHERE Sys_User_Company.UserID=" & bytUserID & ";"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            Do Until R.EOF
                List4.AddItem R![CompName]
            R.MoveNext
            Loop
        End If
        R.Close

        strRecord = "SELECT Sys_Default_Permission.PermName FROM Sys_User_Permission INNER JOIN Sys_Default_Permission" _
            & " ON Sys_User_Permission.PermID = Sys_Default_Permission.PermID WHERE Sys_User_Permission.UserID=" & bytUserID & ";"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            Do Until R.EOF
                List2.AddItem R![PermName]
            R.MoveNext
            Loop
        End If
        R.Close
    End If
    Set R = Nothing
    
    List1.ListIndex = -1
    List2.ListIndex = -1
    List3.ListIndex = -1
    List4.ListIndex = -1
End Sub

Private Sub cboUser_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAdd_Click()
    If List1.Text <> "" Then List2.AddItem List1.Text
End Sub

Private Sub cmdAddAll_Click()
    List2.Clear
    List1.ListIndex = -1
    For inti = 1 To List1.ListCount
        List1.ListIndex = List1.ListIndex + 1
        List2.AddItem List1.Text
    Next inti
End Sub

Private Sub cmdAddC_Click()
    If List3.Text <> "" Then List4.AddItem List3.Text
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Len(cboUser) = 0 Then
        MsgBox "Select User Name", vbInformation, cnstMsgInfo
        cboUser.SetFocus
    Exit Sub
    End If

    If MsgBox("Are you sure to delete this User?", vbInformation + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn

        MainConn.BeginTrans
            MainComm.CommandText = "DELETE Sys_User_Company FROM Sys_User_Company INNER JOIN InventoryPJL.dbo.Sys_User_Name" _
                & " ON Sys_User_Company.UserID = InventoryPJL.dbo.Sys_User_Name.UserID WHERE InventoryPJL.dbo.Sys_User_Name.UserName='" & txtUser.Text & "';"
            MainComm.Execute

            MainComm.CommandText = "DELETE Sys_User_Permission FROM Sys_User_Permission INNER JOIN InventoryPJL.dbo.Sys_User_Name" _
                & " ON Sys_User_Permission.UserID = InventoryPJL.dbo.Sys_User_Name.UserID WHERE InventoryPJL.dbo.Sys_User_Name.UserName='" & txtUser.Text & "';"
            MainComm.Execute

            MainComm.CommandText = "DELETE from InventoryPJL.dbo.Sys_User_Name WHERE UserName='" & txtUser.Text & "';"
            MainComm.Execute
        MainConn.CommitTrans

ErrorHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        err.Number = 0
    Exit Sub
    End If

    MsgBox "Successfully Deleted", vbInformation, cnstMsgInfo
    Screen.MousePointer = vbDefault

    txtUser.Text = ""
    txtPassword.Text = ""
    txtVPassword.Text = ""
    List1.ListIndex = -1
    List3.ListIndex = -1
    List4.Clear
    List2.Clear
    optUser.Value = True
    
    Call addUserName(cboUser)
    cboUser.SetFocus
End Sub

Private Sub cmdModify_Click()
    txtUser.Visible = False
    cboUser.Top = txtUser.Top
    cboUser.Visible = True
    txtPassword.Text = ""
    txtVPassword.Text = ""
    
    List4.Clear
    List2.Clear
    optUser.Value = True
    cboUser.ListIndex = -1
    cboUser.SetFocus
    cmdSave.Caption = "&Update"
    cmdDelete.Enabled = True
End Sub

Private Sub cmdRemove_Click()
    If List2.ListIndex >= 0 Then List2.RemoveItem List2.ListIndex
End Sub

Private Sub cmdRemoveAll_Click()
    List2.Clear
End Sub

Private Sub cmdRemoveC_Click()
    If List4.ListIndex >= 0 Then List4.RemoveItem List4.ListIndex
End Sub

Private Sub cmdSave_Click()
    Dim strUser As String, strPWord As String, strStatus As String, boModify As Boolean
    
    If cmdSave.Caption = "&Update" Then boModify = True Else boModify = False
    If boModify = False Then
        If Len(txtUser) = 0 Then
            MsgBox "Enter User Name", vbInformation, cnstMsgInfo
            txtUser.SetFocus
        Exit Sub
        End If
        
        strUser = EncryptIt(txtUser.Text, 11)
    Else
        If Len(cboUser) = 0 Then
            MsgBox "Select User Name", vbInformation, cnstMsgInfo
            cboUser.SetFocus
        Exit Sub
        End If

        strUser = EncryptIt(cboUser.Text, 11)
    End If
    If Len(txtPassword) = 0 Then
        MsgBox "Enter User Password", vbInformation, cnstMsgInfo
        txtPassword.SetFocus
    Exit Sub
    End If
    
    If Len(txtVPassword) = 0 Then
        MsgBox "Enter Verify Password", vbInformation, cnstMsgInfo
        txtVPassword.SetFocus
    Exit Sub
    End If

    If txtPassword.Text <> txtVPassword.Text Then
        MsgBox "User password does not match with Verify Password", vbInformation, cnstMsgInfo
        txtVPassword.SetFocus
    Exit Sub
    End If

    If optAdmin.Value = False And List2.ListCount = 0 Then
        MsgBox "Permission not found", vbInformation, cnstMsgInfo
    Exit Sub
    End If

    If optAdmin.Value = True Then strStatus = EncryptIt("A", 11)
        strPWord = EncryptIt(txtPassword.Text, 11)
    
        Set R = New ADODB.Recordset
        strRecord = "SELECT UserStatus from InventoryPJL.dbo.Sys_User_Name WHERE UserStatus='" & strStatus & "';"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            MsgBox "User category 'Administrator' not allowed", vbInformation, cnstMsgInfo
        Exit Sub
        End If
        R.Close
    
        If boModify = True Then
            If optUser.Value = True And cboUser.Tag = "A" Then
                MsgBox "'Administrator' not replace with 'User'", vbInformation, cnstMsgInfo
            Exit Sub
            End If
        Else
            strRecord = "SELECT UserName from InventoryPJL.dbo.Sys_User_Name WHERE UserName='" & strUser & "';"
            R.Open strRecord, MainConn, adOpenStatic
            If R.RecordCount > 0 Then
                MsgBox "Duplicate User Name", vbInformation, cnstMsgInfo
                txtUser.SetFocus
            Exit Sub
            End If
            R.Close
        End If

        If boModify = False Then
            strRecord = "SELECT UserPWord from InventoryPJL.dbo.Sys_User_Name WHERE UserPWord='" & strPWord & "';"
        Else
            strRecord = "SELECT UserPWord from InventoryPJL.dbo.Sys_User_Name WHERE ((UserPWord='" & strPWord & "') AND (UserName<>'" & strUser & "'));"
        End If
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            MsgBox "Duplicate password", vbInformation, cnstMsgInfo
            txtPassword.SetFocus
        Exit Sub
        End If
        R.Close

        If boModify = False Then
            lngLastID = 1
            strRecord = "SELECT Max(UserID) AS LastID from InventoryPJL.dbo.Sys_User_Name HAVING Max(UserID) Is Not Null;"
            R.Open strRecord, MainConn, adOpenStatic
            If R.RecordCount > 0 Then lngLastID = lngLastID + R![LastID]
                R.Close
            Else
                strRecord = "SELECT UserID AS LastID from InventoryPJL.dbo.Sys_User_Name WHERE UserName='" & strUser & "';"
                R.Open strRecord, MainConn, adOpenStatic
                If R.RecordCount > 0 Then lngLastID = R![LastID]
                R.Close
            End If
        
            If optUser.Value = True Then strStatus = EncryptIt(CStr(lngLastID), 11)

            On Error GoTo ErrorHandler
            Screen.MousePointer = vbHourglass

            Set MainComm = New ADODB.Command
            MainComm.ActiveConnection = MainConn

            MainConn.BeginTrans
                If boModify = False Then
                    MainComm.CommandText = "INSERT INTO InventoryPJL.dbo.Sys_User_Name (UserID,UserName,UserPWord,UserStatus)" _
                        & " VALUES(" & lngLastID & ",'" & strUser & "','" & strPWord & "','" & strStatus & "');"
                    MainComm.Execute
                Else
                    MainComm.CommandText = "UPDATE InventoryPJL.dbo.Sys_User_Name SET UserPWord='" & strPWord & "' WHERE UserID=" & lngLastID & ";"
                    MainComm.Execute
                    
                    MainComm.CommandText = "DELETE FROM Sys_User_Company WHERE UserID=" & lngLastID & ";"
                    MainComm.Execute
                    
                    MainComm.CommandText = "DELETE FROM Sys_User_Permission WHERE UserID=" & lngLastID & ";"
                    MainComm.Execute
                End If
        
                If optUser.Value = True Then
                    List4.ListIndex = -1
                    For inti = 1 To List4.ListCount
                        List4.ListIndex = List4.ListIndex + 1
                        
'                        MainComm.CommandText = "INSERT INTO Sys_User_Company (UserID,CompID)VALUES(" & lngLastID & "," & bytCompID(List4.Text) & ");"
                        MainComm.Execute
                    Next inti

                    List2.ListIndex = -1
                    For inti = 1 To List2.ListCount
                        List2.ListIndex = List2.ListIndex + 1
                        
                        MainComm.CommandText = "INSERT INTO Sys_User_Permission (UserID,PermID)VALUES(" & lngLastID & "," & findPermID(List2.Text) & ");"
                        MainComm.Execute
                    Next inti
                End If
            MainConn.CommitTrans

ErrorHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        err.Number = 0
    Exit Sub
    End If

    txtPassword.Text = ""
    txtVPassword.Text = ""
    optUser.Value = True
    List4.Clear
    List2.Clear
    List3.ListIndex = -1
    List1.ListIndex = -1

    Screen.MousePointer = vbDefault
    Set R = Nothing

    If boModify = False Then
        MsgBox "Successfully Saved", vbInformation, cnstMsgInfo
        txtUser.Text = ""
        txtUser.SetFocus
    Else
        MsgBox "Successfully Updated", vbInformation, cnstMsgInfo
        Call addUserName(cboUser)
        cboUser.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    
    Set R = New ADODB.Recordset
    strRecord = "SELECT compName FROM Company_Information;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            List3.AddItem R![CompName]
        R.MoveNext
        Loop
    End If
    R.Close

    strRecord = "SELECT PermName FROM Sys_Default_Permission;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            List1.AddItem R![PermName]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
    
    Call addUserName(cboUser)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCreateUser = Nothing
End Sub

Private Sub txtPassword_GotFocus()
    Call txtGotFocus(txtPassword)
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.Text = Trim(txtPassword.Text)
End Sub

Private Sub txtUser_GotFocus()
    Call txtGotFocus(txtUser)
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtUser_LostFocus()
    txtUser.Text = Trim(txtUser.Text)
End Sub

Private Sub txtVPassword_GotFocus()
    Call txtGotFocus(txtVPassword)
End Sub

Private Sub txtVPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtVPassword_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtVPassword_LostFocus()
    txtVPassword.Text = Trim(txtVPassword.Text)
End Sub
