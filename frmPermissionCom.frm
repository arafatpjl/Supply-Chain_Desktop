VERSION 5.00
Begin VB.Form frmPermissionCom 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6149
   ClientLeft      =   39
   ClientTop       =   -65
   ClientWidth     =   7800
   ControlBox      =   0   'False
   Icon            =   "frmPermissionCom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6149
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstCompany 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3380
      Left            =   360
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1590
      Width           =   7035
   End
   Begin VB.ComboBox cmbUserName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   990
      Width           =   4545
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      DisabledPicture =   "frmPermissionCom.frx":12FA
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4920
      Picture         =   "frmPermissionCom.frx":1BC1
      TabIndex        =   2
      Top             =   5475
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      DisabledPicture =   "frmPermissionCom.frx":4155
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6360
      Picture         =   "frmPermissionCom.frx":4A3C
      TabIndex        =   1
      Top             =   5475
      Width           =   1185
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   5295
      Width           =   7995
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   0
      TabIndex        =   10
      Top             =   5250
      Width           =   7875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Company Permission"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.51
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   105
      TabIndex        =   9
      Top             =   90
      Width           =   1230
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   810
      Width           =   7995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   1050
      Width           =   1605
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPermissionCom.frx":70D1
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   1455
      TabIndex        =   6
      Top             =   90
      Width           =   5445
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Company List :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1350
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E26101&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmPermissionCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim strCompany

Private Sub cmbUserName_Click()
    cmbUserName.Tag = 0
    If cmbUserName.ListIndex >= 0 Then
        cmbUserName.Tag = cmbUserName.ItemData(cmbUserName.ListIndex)
        Call prcCheckCompanyPermission
    End If
End Sub

Private Sub cmbUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Form_Load()
On Error GoTo x
    Call prcMakeCenter(Me)
    Call prcAddUserName(cmbUserName)
    Call prcFillList
    Call prcSettingForm
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
    If fncBlank = True Then Exit Sub
    Call fncSaveCompanyPermission
        
    MsgBox "Update successfully.", vbInformation, App.Title
Exit Sub
x:
con.RollbackTrans
MsgBox err.Description, vbCritical, AppTitle
End Sub

Private Sub prcFillList()
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    
    Mysql = "Select CompId, CompName from Company_Information Order By CompId;"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    
    lstCompany.Clear
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            lstCompany.AddItem RS("CompName").Value
            lstCompany.ItemData(lstCompany.NewIndex) = RS("CompID").Value
        RS.MoveNext
        Loop
    End If
    
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPermissionCom = Nothing
End Sub

Private Sub lstCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub prcSettingForm()
    cmbUserName.ListIndex = -1
    lstCompany.Enabled = True
    cmbUserName.Enabled = True
    
''    Dim i As Integer
''    For i = 0 To lstCompany.ListCount - 1
''        lstCompany.Selected(i) = False
''    Next
End Sub

Private Function fncBlank() As Boolean
    fncBlank = True
        Dim blnCompany As Boolean
        For i = 0 To lstCompany.ListCount - 1
            If lstCompany.Selected(i) = True Then
                blnCompany = True
            End If
        Next
        If blnCompany = False Then
            MsgBox "You should provide check mark for company.", vbExclamation, AppTitle
            lstCompany.SetFocus
        Exit Function
        End If
    fncBlank = False
End Function

Private Function fncSaveCompanyPermission()
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn

    Mysql = " Delete From Sys_User_Company Where UserID = " & cmbUserName.Tag & ";"
    MainComm.CommandText = Mysql
    MainComm.Execute
    
    Mysql = ""
    For i = 0 To lstCompany.ListCount - 1
        lstCompany.Tag = 0
        lstCompany.Tag = lstCompany.ItemData(i)
        
        If lstCompany.Selected(i) = True Then
            Mysql = " Insert Into Sys_User_Company (UserID, Compid) Values (" & cmbUserName.Tag & ", " & lstCompany.Tag & ") ;"
            MainComm.CommandText = Mysql
            MainComm.Execute
        End If
    Next
    Set MainComm = Nothing
End Function

Private Sub prcCheckCompanyPermission()
    Dim i As Integer
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    
    For i = 0 To lstCompany.ListCount - 1
        lstCompany.Selected(i) = False
    Next
    
    Mysql = "Select Compid, CompName from viewUserCompany Where UserID = " & cmbUserName.Tag & ";"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            lstCompany.Tag = 0
            
            For i = 0 To lstCompany.ListCount - 1
                lstCompany.Tag = lstCompany.ItemData(i)
                If RS("Compid") = lstCompany.Tag Then lstCompany.Selected(i) = True
            Next
        RS.MoveNext
        Loop
    End If
    Set MainComm = Nothing
End Sub

