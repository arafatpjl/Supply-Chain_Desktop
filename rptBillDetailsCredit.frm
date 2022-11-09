VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form rptBillDetailsCredit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6123
   ClientLeft      =   13
   ClientTop       =   -91
   ClientWidth     =   7397
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6123
   ScaleWidth      =   7397
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtToDate 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1830
      TabIndex        =   14
      Top             =   4740
      Width           =   1575
   End
   Begin VB.TextBox txtFromDate 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1830
      TabIndex        =   13
      Top             =   4350
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   525
      Left            =   5760
      TabIndex        =   16
      Top             =   5400
      Width           =   1185
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   525
      Left            =   4320
      TabIndex        =   15
      Top             =   5400
      Width           =   1185
   End
   Begin VB.ComboBox cmbMRRNO1 
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6630
      TabIndex        =   12
      Text            =   "cmbMRRNO"
      Top             =   3960
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox cmbPartyName1 
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6630
      TabIndex        =   11
      Text            =   "cmbPartyName"
      Top             =   3600
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox cmbHead 
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
      ForeColor       =   &H00000080&
      Height          =   286
      Left            =   1830
      TabIndex        =   10
      Text            =   "cmbHead"
      Top             =   3240
      Width           =   4225
   End
   Begin VB.ComboBox cmbBillNO1 
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6390
      TabIndex        =   9
      Text            =   "cmbBillNO"
      Top             =   2880
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ComboBox cmbIDNO1 
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6630
      TabIndex        =   8
      Text            =   "cmbIDNO1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Report Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   105
      TabIndex        =   17
      Top             =   1020
      Width           =   7275
      Begin VB.OptionButton optIDDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID Date Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   4680
         TabIndex        =   1
         Top             =   325
         Width           =   2025
      End
      Begin VB.OptionButton optBillDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bill Date Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   949
         Width           =   2025
      End
      Begin VB.OptionButton optMRRDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR Date Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   2640
         TabIndex        =   7
         Top             =   949
         Width           =   2415
      End
      Begin VB.OptionButton optMRRNO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR NO Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   4680
         TabIndex        =   6
         Top             =   650
         Width           =   2415
      End
      Begin VB.OptionButton optBillNO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bill No Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   650
         Width           =   2025
      End
      Begin VB.OptionButton optHead 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A/C Head Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   325
         Width           =   2415
      End
      Begin VB.OptionButton optParty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Party Name Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   650
         Width           =   2055
      End
      Begin VB.OptionButton optIDNO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID No Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   480
         TabIndex        =   0
         Top             =   325
         Width           =   2025
      End
   End
   Begin MSDataListLib.DataCombo cmbIDNO 
      Height          =   315
      Left            =   1800
      TabIndex        =   31
      Top             =   2520
      Width           =   4245
      _ExtentX        =   7836
      _ExtentY        =   527
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbBillNO 
      Height          =   315
      Left            =   1800
      TabIndex        =   29
      Top             =   2880
      Width           =   4245
      _ExtentX        =   7836
      _ExtentY        =   527
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPartyName 
      Height          =   315
      Left            =   1800
      TabIndex        =   30
      Top             =   3600
      Width           =   4245
      _ExtentX        =   7836
      _ExtentY        =   527
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbMRRNO 
      Height          =   315
      Left            =   1800
      TabIndex        =   32
      Top             =   3960
      Width           =   4245
      _ExtentX        =   7836
      _ExtentY        =   527
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: Bill Details Information Report (Credit)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.27
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
      TabIndex        =   28
      Top             =   0
      Width           =   6240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8040
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   26
      Top             =   5280
      Width           =   8040
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Date  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   25
      Top             =   4770
      Width           =   1425
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Date  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   6
      Left            =   330
      TabIndex        =   24
      Top             =   4380
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Index           =   5
      Left            =   0
      TabIndex        =   23
      Top             =   5160
      Width           =   8055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MRR NO  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   4
      Left            =   330
      TabIndex        =   22
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   3
      Left            =   330
      TabIndex        =   21
      Top             =   3630
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Head  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   2
      Left            =   330
      TabIndex        =   20
      Top             =   3270
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill NO  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   1
      Left            =   330
      TabIndex        =   19
      Top             =   2910
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID No  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   255
      Index           =   0
      Left            =   330
      TabIndex        =   18
      Top             =   2550
      Width           =   1425
   End
End
Attribute VB_Name = "rptBillDetailsCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstHead As New Recordset
Dim rstParty As New Recordset

Private Sub cmbBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbHead_Click()
On Error GoTo x
    If cmbHead.ListIndex >= 0 Then
        rstHead.MoveFirst
        rstHead.Move cmbHead.ListIndex
        cmbHead.Tag = rstHead("HeadID")
    End If
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbHead_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbIDNo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

'Private Sub cmbPartyName_Click()
'On Error GoTo x
'    If cmbPartyName.ListIndex >= 0 Then
'        rstParty.MoveFirst
'        rstParty.Move cmbPartyName.ListIndex
'        cmbPartyName.Tag = rstParty("PartyID")
'    End If
'Exit Sub
'x:
'MsgBox err.Description, vbCritical, App.Title
'End Sub

Private Sub cmbPartyName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdPreview_Click()
On Error GoTo x
    If fncBlank = True Then Exit Sub
'    Screen.MousePointer = 11
    'Call ShowMessage("Generating Report .....................", Me)
    Call ShowReports
    'Call HideMessage
'    Screen.MousePointer = 0
Exit Sub
x:
Screen.MousePointer = 0
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
''On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call prcAddIDNO
    Call prcAddBillNO
    Call prcAddHead
    Call prcAddParty
    'Call prcAddSupplier(cmbPartyName)
    Call prcAddMRRNO
'    Call prcConReport(CrystalReport1)
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub
Public Sub prcAddParty()

Set rstParty = New ADODB.Recordset
  
     Mysql = "Select Distinct PartyID, PartyName  from Party ORDER BY PARTYNAME;" 'where M.CompID = 1
              rstParty.Open Mysql, MainConn, adOpenStatic
    
         Set cmbPartyName.RowSource = rstParty
             cmbPartyName.ListField = "PartyName"
             cmbPartyName.BoundColumn = "PartyID"

        Set rstParty = Nothing
'    If rstParty.EOF = False And rstParty.BOF = False Then
'        cmbPartyName.Clear
'        rstParty.MoveFirst
'        Do While Not rstParty.EOF
'            cmbPartyName.AddItem rstParty("PartyName")
'            rstParty.MoveNext
'        Loop
'    Else
'        cmbPartyName.Clear
'    End If
    
    
'    Mysql = "Select Distinct SupplierID,SuppName From new_Supplier Order By SuppName"
''    Mysql = "Select Distinct PartyID, PartyName  from Party ORDER BY PARTYNAME;" 'where M.CompID = 1
'  rstParty.Open Mysql, MainConn, adOpenStatic
'    If rstParty.EOF = False And rstParty.BOF = False Then
'        cmbPartyName.Clear
'        rstParty.MoveFirst
'        Do While Not rstParty.EOF
'            cmbPartyName.AddItem rstParty("SuppName")
'            rstParty.MoveNext
'        Loop
'    Else
'        cmbPartyName.Clear
'    End If
End Sub
Public Sub prcAddIDNO()
    Dim tmprst As New Recordset
    Mysql = "Select Distinct IDNO From BillMain Where CompID = " & bytCompID & " and Status ='C' ORDER BY IDNO;"
   tmprst.Open Mysql, MainConn, adOpenStatic
    
    
'    If tmprst.EOF = False And tmprst.BOF = False Then
'        cmbIDNO.Clear
'        tmprst.MoveFirst
'        Do While Not tmprst.EOF
'            cmbIDNO.AddItem tmprst("IDNO")
'            tmprst.MoveNext
'        Loop
'    Else
'        cmbIDNO.Clear
'    End If

             Set cmbIDNO.RowSource = tmprst
             cmbIDNO.ListField = "IDNO"
             cmbIDNO.BoundColumn = "IDNO"
    
    
    Set tmprst = Nothing
End Sub
Public Sub prcAddBillNO()
    Dim tmprst As New Recordset
    
    
    Mysql = "Select Distinct B.BillNO  From BillDetails As B Inner Join BillMain AS M on B.BillID = M.BillID Where M.CompID = " & bytCompID & " And M.Status='C' ORDER BY B.BILLNO;"
    tmprst.Open Mysql, MainConn, adOpenStatic
    
'
'    If tmprst.EOF = False And tmprst.BOF = False Then
'        cmbBillNO.Clear
'        tmprst.MoveFirst
'        Do While Not tmprst.EOF
'            cmbBillNO.AddItem tmprst("BILLNO")
'            tmprst.MoveNext
'        Loop
'    Else
'        cmbBillNO.Clear
'    End If

            Set cmbBillNo.RowSource = tmprst
            cmbBillNo.ListField = "BillNO"
             cmbBillNo.BoundColumn = "BillNO"
    Set tmprst = Nothing
End Sub

Public Sub prcAddHead()
  cmbHead.Clear
  Set R = New ADODB.Recordset
  strRecord = "Select Distinct HeadID, HeadName From AccountsHead ORDER BY HEADNAME;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    Do Until R.EOF
      cmbHead.AddItem R![HeadName]
      R.MoveNext
    Loop
  End If
  R.Close
  Set R = Nothing

''    Mysql = "Select Distinct HeadID, HeadName  From AccountsHead ORDER BY HEADNAME ;"   'where M.CompID = 1"
''    rstHead.Open Mysql, MainConn, adOpenStatic
''    If rstHead.EOF = False And rstHead.BOF = False Then
''        cmbHead.Clear
''        rstHead.MoveFirst
''        Do While Not rstHead.EOF
''            cmbHead.AddItem rstHead("HeadName")
''            rstHead.MoveNext
''        Loop
''    Else
''        cmbHead.Clear
''    End If
End Sub

'Public Sub prcAddParty()
'    Mysql = "Select Distinct SupplierID, SupplierName  From Supplier_Information ORDER BY SupplierID;"    'where M.CompID = 1
'    com.CommandText = Mysql
'    Set rstParty = com.Execute
'    If rstParty.EOF = False And rstParty.BOF = False Then
'        cmbPartyName.Clear
'        rstParty.MoveFirst
'        Do While Not rstParty.EOF
'            'cmbPartyName.AddItem rstParty("PartyName")
'            cmbPartyName.AddItem rstParty("SupplierID") & " " & rstParty("SupplierName")
'            rstParty.MoveNext
'        Loop
'    Else
'        cmbPartyName.Clear
'    End If
'End Sub
Public Sub prcAddMRRNO()
Dim tmprst As New Recordset
    Mysql = "Select Distinct B.MRRNO  From BillDetails As B Inner Join BillMain AS M on B.BillID = M.BillID Where M.CompID = " & bytCompID & " And M.Status = 'C' ORDER BY B.MRRNO;"
    tmprst.Open Mysql, MainConn, adOpenStatic
    
'    If tmprst.EOF = False And tmprst.BOF = False Then
'        cmbMRRNO.Clear
'        tmprst.MoveFirst
'        Do While Not tmprst.EOF
'            cmbMRRNO.AddItem tmprst("MRRNO")
'            tmprst.MoveNext
'        Loop
'    Else
'        cmbMRRNO.Clear
'    End If
     Set cmbMRRNo.RowSource = tmprst
                   cmbMRRNo.ListField = "MRRNO"
                   cmbMRRNo.BoundColumn = "MRRNO"

    Set tmprst = Nothing
End Sub

Private Sub optBillDate_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = False
    cmbHead.Enabled = False
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optBillDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optBillNO_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = True
    cmbHead.Enabled = False
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optBillNO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optHead_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = False
    cmbHead.Enabled = True
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optHead_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optIDDate_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = False
    cmbHead.Enabled = False
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optIDDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optIDNO_Click()
On Error Resume Next
    cmbIDNO.Enabled = True
    cmbBillNo.Enabled = False
    cmbHead.Enabled = False
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optIDNO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optMRRDate_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = False
    cmbHead.Enabled = False
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optMRRDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optMRRNO_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = False
    cmbHead.Enabled = False
    cmbPartyName.Enabled = False
    cmbMRRNo.Enabled = True
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optParty_Click()
On Error Resume Next
    cmbIDNO.Enabled = False
    cmbBillNo.Enabled = False
    cmbHead.Enabled = False
    cmbPartyName.Enabled = True
    cmbMRRNo.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optParty_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFromDate_GotFocus()
On Error Resume Next
    txtFromDate.SelStart = 0
    txtFromDate.SelLength = Len(txtFromDate.Text)
End Sub

Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtFromDate_LostFocus()
On Error Resume Next
    Call FormatDate(txtFromDate)
End Sub

Private Sub txtToDate_GotFocus()
On Error Resume Next
    txtToDate.SelStart = 0
    txtToDate.SelLength = Len(txtToDate.Text)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtToDate_LostFocus()
On Error Resume Next
    Call FormatDate(txtToDate)
End Sub

Public Function fncBlank() As Boolean
fncBlank = True
    If optIDNO.Value = True Then
        If cmbIDNO.Text = "" Then
            MsgBox "You should provide ID NO.", vbCritical, App.Title
            cmbIDNO.SetFocus
            Exit Function
        End If
    ElseIf optBillNO.Value = True Then
        If cmbBillNo.Text = "" Then
            MsgBox "You should provide Bill NO.", vbCritical, App.Title
            cmbBillNo.SetFocus
            Exit Function
        End If
    ElseIf optBillDate.Value = True Or optIDDate.Value = True Then
        If Trim(txtFromDate.Text) = "" Then
            MsgBox "You should provide from date.", vbCritical, App.Title
            txtFromDate.SetFocus
            Exit Function
        End If
        If Trim(txtToDate.Text) = "" Then
            MsgBox "You should provide to date.", vbCritical, App.Title
            txtToDate.SetFocus
            Exit Function
        End If
    ElseIf optHead.Value = True Then
        If cmbHead.Text = "" Then
            MsgBox "You should provide A/C Head Name.", vbCritical, App.Title
            cmbHead.SetFocus
            Exit Function
        End If
        If Trim(txtFromDate.Text) = "" Then
            MsgBox "You should provide from date.", vbCritical, App.Title
            txtFromDate.SetFocus
            Exit Function
        End If
        If Trim(txtToDate.Text) = "" Then
            MsgBox "You should provide to date.", vbCritical, App.Title
            txtToDate.SetFocus
            Exit Function
        End If
    ElseIf optParty.Value = True Then
        If cmbPartyName.Text = "" Then
            MsgBox "You should provide party name.", vbCritical, App.Title
            cmbPartyName.SetFocus
            Exit Function
        End If
        If Trim(txtFromDate.Text) = "" Then
            MsgBox "You should provide from date.", vbCritical, App.Title
            txtFromDate.SetFocus
            Exit Function
        End If
        If Trim(txtToDate.Text) = "" Then
            MsgBox "You should provide to date.", vbCritical, App.Title
            txtToDate.SetFocus
            Exit Function
        End If
    ElseIf optMRRNO.Value = True Then
        If cmbMRRNo.Text = "" Then
            MsgBox "You should provide MRR NO.", vbCritical, App.Title
            cmbMRRNo.SetFocus
            Exit Function
        End If
    ElseIf optMRRDate.Value = True Then
        If Trim(txtFromDate.Text) = "" Then
            MsgBox "You should provide from date.", vbCritical, App.Title
            txtFromDate.SetFocus
            Exit Function
        End If
        If Trim(txtToDate.Text) = "" Then
            MsgBox "You should provide to date.", vbCritical, App.Title
            txtToDate.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function
Public Function findPartyID(parItem As String) As Integer
    Set R = New ADODB.Recordset
    strRecord = "Select Distinct PartyID, PartyName  from Party WHERE PartyName='" & parItem & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findPartyID = R![PartyID]
    R.Close
    Set R = Nothing
End Function

Public Sub ShowReports()
Dim strDateRange As String
Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    strDateRange = "From Date : " & Trim(txtFromDate.Text) & " To : " & Trim(txtToDate.Text)

        If optIDNO.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsIDNo.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE IDNO='" & cmbIDNO.Text & "' And Status='C' And CompID = " & bytCompID & " Order By RowNO"
        
        ElseIf optIDDate.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsIDDate.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE IDDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' And Status='C' And CompID = " & bytCompID & " Order By IDDate,RowNO"
        
        ElseIf optBillNO.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsBillNO.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE BillNO ='" & Trim(cmbBillNo.Text) & "' And Status='C' And CompID = " & bytCompID & " Order By RowNO"
        
        ElseIf optBillDate.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsBillDate.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE BillDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' And Status='C' And CompID = " & bytCompID & " Order By BillDate,RowNO"
        
        ElseIf optHead.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsHead.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE HeadID = " & cmbHead.Tag & " And BillDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' And Status='C' And CompID = " & bytCompID & " Order By BillDate,BillNO,RowNO"
        
        ElseIf optParty.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsParty.rpt")
       ' strRecord = "SELECT * FROM vieBillDetails WHERE BillDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' And PartyID = " & cmbPartyName.Tag & " And Status='C' And CompID = " & bytCompID & " Order By BillDate,BillNO,RowNO"
         strRecord = "SELECT * FROM vieBillDetails WHERE BillDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' And PartyID = " & findPartyID(cmbPartyName) & " And Status='C' And CompID = " & bytCompID & " Order By BillDate,BillNO,RowNO"

        ElseIf optMRRNO.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsMRRNO.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE MRRNo ='" & Trim(cmbMRRNo.Text) & "' And Status='C' And CompID = " & bytCompID & " Order By MRRNO,RowNO"
        
        ElseIf optMRRDate.Value = True Then
        Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsMRRDate.rpt")
        strRecord = "SELECT * FROM vieBillDetails WHERE MRRDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' And Status='C' And CompID = " & bytCompID & " Order By MRRNO"
        
        End If
    Set R = New ADODB.Recordset
    R.Open strRecord, MainConn, adOpenStatic
    
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = REPORT
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    R.Close
    
End Sub
