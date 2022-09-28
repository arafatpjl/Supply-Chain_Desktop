VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPermissionMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7735
   ClientLeft      =   39
   ClientTop       =   -65
   ClientWidth     =   7631
   ControlBox      =   0   'False
   Icon            =   "frmPermissionMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7735
   ScaleWidth      =   7631
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   90
      TabIndex        =   6
      Top             =   1440
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   10615
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Main Menu"
      TabPicture(0)   =   "frmPermissionMenu.frx":12FA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSaveMM"
      Tab(0).Control(1)=   "lstUserMM"
      Tab(0).Control(2)=   "Label3"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Sub Menu"
      TabPicture(1)   =   "frmPermissionMenu.frx":1316
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmbMS_MM"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lstUserMS"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdSaveMS"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Sub Sub Menu"
      TabPicture(2)   =   "frmPermissionMenu.frx":1332
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(3)=   "cmbMSS_MM"
      Tab(2).Control(4)=   "cmdSaveMSS"
      Tab(2).Control(5)=   "cmbMSS_MS"
      Tab(2).Control(6)=   "lstUserMSS"
      Tab(2).ControlCount=   7
      Begin VB.ListBox lstUserMSS 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4004
         ItemData        =   "frmPermissionMenu.frx":134E
         Left            =   -74820
         List            =   "frmPermissionMenu.frx":1355
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   1320
         Width           =   7035
      End
      Begin VB.ComboBox cmbMSS_MS 
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
         Left            =   -72900
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   780
         Width           =   5115
      End
      Begin VB.CommandButton cmdSaveMSS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         DisabledPicture =   "frmPermissionMenu.frx":1365
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69360
         Picture         =   "frmPermissionMenu.frx":1C2C
         TabIndex        =   17
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         DisabledPicture =   "frmPermissionMenu.frx":41C0
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5640
         Picture         =   "frmPermissionMenu.frx":4A87
         TabIndex        =   16
         Top             =   5520
         Width           =   1575
      End
      Begin VB.ListBox lstUserMS 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4420
         ItemData        =   "frmPermissionMenu.frx":701B
         Left            =   180
         List            =   "frmPermissionMenu.frx":701D
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   930
         Width           =   7035
      End
      Begin VB.ComboBox cmbMSS_MM 
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
         Left            =   -72900
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   390
         Width           =   5115
      End
      Begin VB.CommandButton cmdSaveMM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         DisabledPicture =   "frmPermissionMenu.frx":701F
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69360
         Picture         =   "frmPermissionMenu.frx":78E6
         TabIndex        =   11
         Top             =   5520
         Width           =   1575
      End
      Begin VB.ListBox lstUserMM 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4628
         ItemData        =   "frmPermissionMenu.frx":9E7A
         Left            =   -74820
         List            =   "frmPermissionMenu.frx":9E7C
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   690
         Width           =   7035
      End
      Begin VB.ComboBox cmbMS_MM 
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
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   420
         Width           =   5145
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Sub Menu List :"
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
         Left            =   -74820
         TabIndex        =   21
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Main Menu :"
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
         Left            =   -74790
         TabIndex        =   19
         Top             =   810
         Width           =   1800
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Menu List :"
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
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Main Menu :"
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
         Left            =   -74790
         TabIndex        =   13
         Top             =   450
         Width           =   1800
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Menu List :"
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
         Left            =   -74790
         TabIndex        =   10
         Top             =   450
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Main Menu :"
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
         Left            =   210
         TabIndex        =   8
         Top             =   450
         Width           =   1800
      End
   End
   Begin VB.ComboBox cmbUserName 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1710
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   990
      Width           =   3945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      DisabledPicture =   "frmPermissionMenu.frx":9E7E
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6120
      Picture         =   "frmPermissionMenu.frx":A745
      TabIndex        =   1
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Menu Permission"
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
      Left            =   120
      TabIndex        =   22
      Top             =   90
      Width           =   1230
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   5
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
      Left            =   360
      TabIndex        =   4
      Top             =   1020
      Width           =   1605
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmPermissionMenu.frx":CCD9
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
      Left            =   1440
      TabIndex        =   3
      Top             =   90
      Width           =   5970
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
      TabIndex        =   2
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmPermissionMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim strCompany

Private Sub cmbMS_MM_Click()
    cmbMS_MM.Tag = 0
    If cmbMS_MM.ListIndex >= 0 Then
        cmbMS_MM.Tag = cmbMS_MM.ItemData(cmbMS_MM.ListIndex)
        
        Call prcFillListMS(lstUserMS, cmbMS_MM.Tag)
        Call prcCheckPermMS(lstUserMS, cmbUserName.Tag, cmbMS_MM.Tag)
    End If
End Sub

Private Sub cmbMSS_MM_Click()
    lstUserMSS.Clear
    cmbMSS_MM.Tag = 0
    If cmbMSS_MM.ListIndex >= 0 Then
        cmbMSS_MM.Tag = cmbMSS_MM.ItemData(cmbMSS_MM.ListIndex)
        
        Call prcAddMS(cmbMSS_MS, cmbUserName.Tag, cmbMSS_MM.Tag)
    End If
End Sub

Private Sub cmbMSS_MS_Click()
    cmbMSS_MS.Tag = 0
    If cmbMSS_MS.ListIndex >= 0 Then
        cmbMSS_MS.Tag = cmbMSS_MS.ItemData(cmbMSS_MS.ListIndex)
        
        Call prcFillListMSS(lstUserMSS, cmbMSS_MM.Tag, cmbMSS_MS.Tag)
        Call prcCheckPermMSS(lstUserMSS, cmbUserName.Tag, cmbMSS_MM.Tag, cmbMSS_MS.Tag)
    End If
End Sub

Private Sub cmbUserName_Click()
    lstUserMS.Clear
    lstUserMSS.Clear
    
    cmbUserName.Tag = 0
    If cmbUserName.ListIndex >= 0 Then
        cmbUserName.Tag = cmbUserName.ItemData(cmbUserName.ListIndex)
        
        Call prcAddMM(cmbMS_MM, cmbUserName.Tag)
        Call prcAddMM(cmbMSS_MM, cmbUserName.Tag)
        Call prcCheckPermMM(lstUserMM, cmbUserName.Tag)
    End If
End Sub

Private Sub cmbUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdSaveMM_Click()
On Error GoTo x
    Dim MenuSel As Integer
    
    If Len(cmbUserName.Text) = 0 Then
        MsgBox "Select User Name.", vbCritical
        cmbUserName.SetFocus
    Exit Sub
    End If

    MenuSel = 0
    For i = 0 To lstUserMM.ListCount - 1
        If lstUserMM.Selected(i) = True Then MenuSel = MenuSel + 1
    Next
'    If MenuSel = 0 Then
'        MsgBox "Select a Menu Name.", vbCritical
'        lstUserMM.SetFocus
'    Exit Sub
'    End If
    
    If MsgBox("Do you want to save main menu permission.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn

    Mysql = " Delete From Sys_User_MM Where UserID = " & cmbUserName.Tag & ";"
    MainComm.CommandText = Mysql
    MainComm.Execute
    
'    Mysql = "SET NOCOUNT ON Delete From Sys_User_MS Where UserID = " & cmbUserName.Tag & " "
'    MainComm.CommandText = Mysql
'    MainComm.Execute
'
'    Mysql = "SET NOCOUNT ON Delete From Sys_User_MSS Where UserID = " & cmbUserName.Tag & ""
'    MainComm.CommandText = Mysql
'    MainComm.Execute
    
    Mysql = ""
    For i = 0 To lstUserMM.ListCount - 1
        lstUserMM.Tag = 0
        lstUserMM.Tag = lstUserMM.ItemData(i)
        
        If lstUserMM.Selected(i) = True Then
            Mysql = " Insert Into Sys_User_MM (UserID, Menu_Main_Id) Values (" & cmbUserName.Tag & ", " & lstUserMM.Tag & ") ;"
            MainComm.CommandText = Mysql
            MainComm.Execute
        End If
    Next
    Set MainComm = Nothing
    MsgBox "Sucessfully Save Main Menu Permission.", vbInformation
    
Exit Sub
x:
MsgBox err.Description, vbCritical
End Sub

Private Sub cmdSaveMS_Click()
On Error GoTo x
    Dim MenuSel As Integer
    
    If Len(cmbUserName.Text) = 0 Then
        MsgBox "Select User Name.", vbCritical
        cmbUserName.SetFocus
    Exit Sub
    End If
    If Len(cmbMS_MM.Text) = 0 Then
        MsgBox "Select Main Menu.", vbCritical
        cmbMS_MM.SetFocus
    Exit Sub
    End If

    MenuSel = 0
    For i = 0 To lstUserMS.ListCount - 1
        If lstUserMS.Selected(i) = True Then MenuSel = MenuSel + 1
    Next
'    If MenuSel = 0 Then
'        MsgBox "Select a Sub-Menu Name.", vbCritical
'        lstUserMS.SetFocus
'    Exit Sub
'    End If
    
    If MsgBox("Do you want to save Sub-Main menu permission.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn

    Mysql = "SET NOCOUNT ON Delete From Sys_User_MS Where UserID = " & cmbUserName.Tag & " And Menu_Main_Id = " & cmbMS_MM.Tag & ";"
'      Mysql = "SET NOCOUNT ON Delete From Sys_User_MS Where UserID = " & cmbUserName.Tag & " ;"
    MainComm.CommandText = Mysql
    MainComm.Execute
    
    Mysql = ""
    For i = 0 To lstUserMS.ListCount - 1
        lstUserMS.Tag = 0
        lstUserMS.Tag = lstUserMS.ItemData(i)
        
        If lstUserMS.Selected(i) = True Then
            Mysql = " Insert Into Sys_User_MS (UserID, Menu_Sub_Id, Menu_Main_Id) Values (" & cmbUserName.Tag & ", " & lstUserMS.Tag & ", " & cmbMS_MM.Tag & ") ;"
            MainComm.CommandText = Mysql
            MainComm.Execute
        End If
    Next
    Set MainComm = Nothing
    MsgBox "Sucessfully Save Main Menu Permission.", vbInformation
Exit Sub
x:
MsgBox err.Description, vbCritical
End Sub

Private Sub cmdSaveMSS_Click()
On Error GoTo x
    Dim MenuSel As Integer
    
    If Len(cmbUserName.Text) = 0 Then
        MsgBox "Select User Name.", vbCritical
        cmbUserName.SetFocus
    Exit Sub
    End If
    If Len(cmbMSS_MM.Text) = 0 Then
        MsgBox "Select Main Menu.", vbCritical
        cmbMSS_MM.SetFocus
    Exit Sub
    End If
    If Len(cmbMSS_MS.Text) = 0 Then
        MsgBox "Select Sub-Menu.", vbCritical
        cmbMSS_MM.SetFocus
    Exit Sub
    End If
    
    MenuSel = 0
    For i = 0 To lstUserMSS.ListCount - 1
        If lstUserMSS.Selected(i) = True Then MenuSel = MenuSel + 1
    Next
'    If MenuSel = 0 Then
'        MsgBox "Select a Sub-Sub-Menu Name.", vbCritical
'        lstUserMSS.SetFocus
'    Exit Sub
'    End If
    
    If MsgBox("Do you want to save Sub-Sub-Main menu permission.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn

    Mysql = "SET NOCOUNT ON Delete From Sys_User_MSS Where UserID = " & cmbUserName.Tag & " And Menu_Main_Id = " & cmbMSS_MM.Tag & " And Menu_Sub_Id = " & cmbMSS_MS.Tag & ";"
'    Mysql = "SET NOCOUNT ON Delete From Sys_User_MSS Where UserID = " & cmbUserName.Tag & " ;"
    MainComm.CommandText = Mysql
    MainComm.Execute
    
    Mysql = ""
    For i = 0 To lstUserMSS.ListCount - 1
        lstUserMSS.Tag = 0
        lstUserMSS.Tag = lstUserMSS.ItemData(i)
        
        If lstUserMSS.Selected(i) = True Then
            Mysql = " Insert Into Sys_User_MSS (UserID, Menu_Sub_Sub_ID, Menu_Sub_Id, Menu_Main_Id) Values (" & cmbUserName.Tag & ", " & lstUserMSS.Tag & ", " & cmbMSS_MS.Tag & ",  " & cmbMSS_MM.Tag & ") ;"
            MainComm.CommandText = Mysql
            MainComm.Execute
        End If
    Next
    Set MainComm = Nothing
    MsgBox "Sucessfully Save Sub-Sub-Main Menu Permission.", vbInformation
Exit Sub
x:
MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()
On Error GoTo x
    Call prcMakeCenter(Me)
    
    Call prcAddUserName(cmbUserName)
    Call prcFillListMM(lstUserMM)
    
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

Private Sub Form_Unload(Cancel As Integer)
    Set frmPermissionCom = Nothing
End Sub

Private Sub lstUserMM_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub prcSettingForm()
    cmbUserName.ListIndex = -1
    lstUserMM.Enabled = True
    cmbUserName.Enabled = True

End Sub

Private Function fncBlank() As Boolean
    fncBlank = True
        Dim blnCompany As Boolean
        For i = 0 To lstUserMM.ListCount - 1
            If lstUserMM.Selected(i) = True Then
                blnCompany = True
            End If
        Next
        If blnCompany = False Then
            MsgBox "You should provide check mark for company.", vbExclamation, AppTitle
            lstUserMM.SetFocus
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
    For i = 0 To lstUserMM.ListCount - 1
        lstUserMM.Tag = 0
        lstUserMM.Tag = lstUserMM.ItemData(i)
        
        If lstUserMM.Selected(i) = True Then
            Mysql = " Insert Into Sys_User_Company (UserID, Compid) Values (" & cmbUserName.Tag & ", " & lstUserMM.Tag & ") ;"
            MainComm.CommandText = Mysql
            MainComm.Execute
        End If
    Next
    Set MainComm = Nothing
End Function
