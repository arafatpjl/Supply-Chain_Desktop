VERSION 5.00
Begin VB.Form frmLockData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3930
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6690
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLockData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIncrement 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Increment"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3360
      TabIndex        =   7
      Top             =   2790
      Width           =   2145
   End
   Begin VB.CheckBox chkASalary 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advance Salary"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3360
      TabIndex        =   5
      Top             =   2070
      Width           =   2325
   End
   Begin VB.CheckBox chkAAttendance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advance Attendance"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   2325
   End
   Begin VB.CheckBox chkAttendance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attendance"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   990
      TabIndex        =   1
      Top             =   1680
      Width           =   2145
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   405
      Left            =   5220
      TabIndex        =   8
      Top             =   3420
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   405
      Left            =   240
      TabIndex        =   9
      Top             =   3420
      Width           =   1305
   End
   Begin VB.CheckBox chkOvertime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monthly Overtime"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   990
      TabIndex        =   3
      Top             =   2430
      Width           =   2145
   End
   Begin VB.CheckBox chkFBonus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Festival Bonus"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3360
      TabIndex        =   6
      Top             =   2430
      Width           =   2325
   End
   Begin VB.CheckBox chkSalary 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monthly Salary"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   990
      TabIndex        =   2
      Top             =   2070
      Width           =   2145
   End
   Begin VB.ComboBox cmbMonth 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1050
      Width           =   2925
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
      TabIndex        =   14
      Top             =   840
      Width           =   6825
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Data Editing"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   330
      TabIndex        =   13
      Top             =   240
      Width           =   5025
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
      Left            =   0
      TabIndex        =   11
      Top             =   3330
      Width           =   6825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Month :"
      ForeColor       =   &H00004080&
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   10
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   795
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   3330
      Width           =   6795
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6795
   End
End
Attribute VB_Name = "frmLockData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAAttendance_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub chkASalary_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub chkAttendance_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub chkFBonus_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub chkIncrement_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub chkOvertime_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub chkSalary_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbMonth_Click()
    cmbMonth.Tag = 0
    If cmbMonth.ListIndex >= 0 Then cmbMonth.Tag = cmbMonth.ItemData(cmbMonth.ListIndex)
    
    Call prcDisplayData
End Sub

Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo x
    If MsgBox("Do you want to update data?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    If fncBlank = True Then Exit Sub
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    MainConn.CommandTimeout = 500
    
    MainConn.BeginTrans
        'Insert Record If Locking Row Doesn't Exist
        Mysql = "Exec prcLockData " & bytCompID & ", " & cmbMonth.Tag & ""
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        'Updating Lock Data
        Mysql = "Update Emp_Lock_Data Set Attendance = " & chkAAttendance.Value & ", AAttendance = " & chkAAttendance.Value & ", Salary = " & chkSalary.Value & ", "
        Mysql = Mysql + "ASalary = " & chkASalary.Value & ", Overtime = " & chkOvertime.Value & ", FBonus = " & chkFBonus.Value & ", Increment = " & chkIncrement.Value & ""
        Mysql = Mysql + " Where CompId = " & bytCompID & " and Monthid = " & cmbMonth.Tag & ""
        MainComm.CommandText = Mysql
        MainComm.Execute
    MainConn.CommitTrans
    MsgBox "Data Update Successfully.", vbInformation
    Set MainComm = Nothing
Exit Sub
x:
MainConn.RollbackTrans
Set MainComm = nthing
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    
'    Call addMonthName(cmbMonth)
    cmbMonth.ListIndex = Month(Date) - 1
End Sub

Private Sub prcDisplayData()
On Error GoTo x
    ''clear check box
    chkAttendance.Value = 0
    chkAAttendance.Value = 0
    chkSalary.Value = 0
    chkASalary.Value = 0
    chkOvertime.Value = 0
    chkFBonus.Value = 0
    chkIncrement.Value = 0
    
    If Len(cmbMonth.Text) = 0 Then Exit Sub
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    MainConn.CommandTimeout = 500
    
    Set R = New Recordset
        Mysql = "Select * From Emp_Lock_Data Where ((compid = " & bytCompID & ") and (MonthId = " & cmbMonth.Tag & "))"
        MainComm.CommandText = Mysql
        Set R = MainComm.Execute
                
        blnDataExist = False
        If R.EOF = False And R.BOF = False Then
            chkAttendance.Value = R("Attendance").Value
            chkAAttendance.Value = R("AAttendance").Value
            chkSalary.Value = R("Salary").Value
            chkASalary.Value = R("ASalary").Value
            chkOvertime.Value = R("Overtime").Value
            chkFBonus.Value = R("FBonus").Value
            chkIncrement.Value = R("Increment").Value
        End If
    Set R = Nothing
    Set MainComm = Nothing
Exit Sub
x:
Set R = Nothing
Set MainComm = Nothing
MsgBox Err.Description, vbCritical, App.Title
End Sub

Private Function fncBlank() As Boolean
fncBlank = True
    If Len(cmbMonth.Text) = 0 Then
        MsgBox "You should provide month.", vbCritical, App.Title
        cmbMonth.SetFocus
    Exit Function
    End If
fncBlank = False
End Function

Private Sub Label5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, 161, 2, 0
End Sub
