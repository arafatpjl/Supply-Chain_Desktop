VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmChequeInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8190
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChequeInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSL 
      Height          =   315
      Left            =   10560
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtPrefix 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Top             =   1740
      Width           =   615
   End
   Begin VB.ComboBox cmbIssueChq 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2520
      Width           =   2205
   End
   Begin VB.CommandButton cmdAddBank 
      BackColor       =   &H00808080&
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
      Left            =   3540
      Picture         =   "frmChequeInfo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Add New Supplier"
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.ComboBox cmbBookNO 
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1050
      Width           =   3315
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9510
      TabIndex        =   15
      Top             =   3172
      Width           =   1035
   End
   Begin VB.CheckBox chkCancel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANCEL :"
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
      Left            =   8280
      TabIndex        =   14
      Top             =   3210
      Width           =   1095
   End
   Begin VB.TextBox txtRemarks 
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
      Height          =   315
      Left            =   4140
      TabIndex        =   13
      Top             =   3210
      Width           =   4005
   End
   Begin VB.TextBox txtEncashmentDate 
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
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   3210
      Width           =   1635
   End
   Begin VB.TextBox txtPaymentParticulars 
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
      Height          =   315
      Left            =   90
      TabIndex        =   11
      Top             =   3210
      Width           =   2205
   End
   Begin VB.TextBox txtBeneficiary 
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
      Height          =   315
      Left            =   5970
      TabIndex        =   10
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox txtCHQValue 
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
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      Top             =   2520
      Width           =   1725
   End
   Begin VB.TextBox txtIssueDate 
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
      Height          =   315
      Left            =   2370
      TabIndex        =   8
      Top             =   2520
      Width           =   1665
   End
   Begin VB.TextBox txtIssueCHQNO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EDECE9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtTotalPage 
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1635
   End
   Begin VB.TextBox txtCHQNOTo 
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
      Height          =   315
      Left            =   6630
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1740
      Width           =   2415
   End
   Begin VB.TextBox txtCHQNOFrom 
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
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   1740
      Width           =   2055
   End
   Begin VB.ComboBox cmbBank 
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
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   3435
   End
   Begin VB.CommandButton cmdAddNew 
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
      Left            =   960
      TabIndex        =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00EDECE9&
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
      Height          =   525
      Left            =   7800
      TabIndex        =   18
      Top             =   7440
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00EDECE9&
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
      Left            =   6360
      TabIndex        =   20
      Top             =   7440
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00EDECE9&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      CausesValidation=   0   'False
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
      Left            =   9120
      TabIndex        =   22
      Top             =   7440
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3555
      Left            =   60
      TabIndex        =   16
      ToolTipText     =   "Select && Press Delete To Remove a Row"
      Top             =   3630
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   6271
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque Information "
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
      Height          =   510
      Index           =   2
      Left            =   240
      TabIndex        =   39
      Top             =   120
      Width           =   5040
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   12120
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   37
      Top             =   7320
      Width           =   12240
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Prefix"
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
      Height          =   195
      Left            =   3060
      TabIndex        =   36
      Top             =   1470
      Width           =   705
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   2130
      Width           =   10935
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   -30
      TabIndex        =   32
      Top             =   7230
      Width           =   10815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "REMARKS :"
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
      Height          =   225
      Left            =   4170
      TabIndex        =   31
      Top             =   2970
      Width           =   1035
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF HONOUR. :"
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
      Height          =   225
      Left            =   2400
      TabIndex        =   30
      Top             =   2970
      Width           =   1545
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT PARTICULARS. :"
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
      Height          =   225
      Left            =   90
      TabIndex        =   29
      Top             =   2970
      Width           =   2145
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "BENEFICIARY :"
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
      Height          =   225
      Left            =   5970
      TabIndex        =   28
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CHQ VALUE :"
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
      Height          =   225
      Left            =   4170
      TabIndex        =   27
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUE DATE :"
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
      Height          =   225
      Left            =   2400
      TabIndex        =   26
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUE CHQ# :"
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
      Height          =   225
      Left            =   90
      TabIndex        =   25
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PAGE"
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
      Height          =   225
      Left            =   90
      TabIndex        =   24
      Top             =   1470
      Width           =   1305
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CHQUE BOOK NO"
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
      Height          =   225
      Left            =   5790
      TabIndex        =   23
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No  From"
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
      Height          =   225
      Left            =   4110
      TabIndex        =   21
      Top             =   1470
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No To"
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
      Height          =   195
      Left            =   6630
      TabIndex        =   19
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHQ ISSUE BANK "
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
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   810
      Width           =   1545
   End
End
Attribute VB_Name = "frmChequeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim boEdit As Boolean
Dim bytRowNo As Integer
Dim intGSLNO As Integer
Dim intGMaxSLNO As Long

Private Sub chkCancel_Click()
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
    Else
        chkCancel.ForeColor = &H80000012
    End If
End Sub

Private Sub chkCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub cmbBank_Click()
On Error GoTo x
    cmbBank.Tag = 0
    cmbBookNo.Clear:    cmbBookNo.Text = ""
    
    If cmbBank.ListIndex >= 0 Then
        cmbBank.Tag = cmbBank.ItemData(cmbBank.ListIndex)
        Call prcAddBookNo
    End If
    
    txtCHQNOFrom.Text = ""
    txtCHQNOTo.Text = ""
    txtTotalPage.Text = ""

    flexDetails.Rows = 1
    
    Call prcBlank(1)
    cmbBank.ToolTipText = cmbBank.Text
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbBank_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbBookNo_Click()
On Error GoTo x
    cmbBookNo.ToolTipText = cmbBookNo.Text
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbBookNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbBookNO_LostFocus()
On Error GoTo x
Dim dataUpdate As Integer
    Call prcSetConnectionActive
    If Len(cmbBank.Text) = 0 Then Exit Sub
    If Len(cmbBookNo.Text) = 0 Then Exit Sub
    
    txtCHQNOFrom.Text = ""
    txtCHQNOTo.Text = ""
    txtTotalPage.Text = 0
    txtPrefix.Text = ""
    flexDetails.Rows = 1
        
   ''150111''''   cmdDelete.Enabled = False
'    cmdSave.Enabled = False
    'Set R = New ADODB.Recordset
    Set R = New ADODB.Recordset
    strRecord = "SELECT IDNo, COMPID, BankID, BankName, ChqBookNo, ChqStartNO, ChqEndNO, ChequeNo, IssueDate, ChequeValue, Beneficiary, PaymentParticulars, EncashmentDate, Remarks, Status,ChequeLot,Prefix FROM QueryChequeInfo WHERE BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' AND COMPID = " & bytCompID & " order by ChequeNo;"
    R.Open strRecord, MainConn
    
    dataUpdate = 0
    flexDetails.Tag = 0
    If R.EOF = False And R.BOF = False Then
        dataUpdate = 1
''150111''''        If canDelete = 1 Then cmdDelete.Enabled = True
'        cmdSave.Enabled = True
        
        flexDetails.Tag = R.Fields("IDNO").Value
        txtPrefix.Text = R.Fields("Prefix").Value & ""
        txtCHQNOFrom.Text = R.Fields("CHQStartNO").Value
        txtCHQNOTo.Text = R.Fields("CHQEndNo").Value
        txtTotalPage.Text = R.Fields("ChequeLot").Value
        With flexDetails
            .Rows = 1
            Do Until R.EOF
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .Col = 0: .Text = R("ChequeNO") & ""
                .Col = 1
                    If R("IssueDate") = "12:00:00 AM" Then
                        .Text = ""
                    Else
                        .Text = R("IssueDate").Value & ""
                    End If
                .Col = 2: .Text = FormatNumber(R("ChequeValue"), 0) & ""
                .Col = 3: .Text = R("Beneficiary").Value & ""
                .Col = 4: .Text = R("PaymentParticulars").Value & ""
                .Col = 5
                    If R("EncashmentDate") = "12:00:00 AM" Then
                        .Text = ""
                    Else
                        .Text = R("EncashmentDate").Value & ""
                    End If
                .Col = 6: .Text = R("Remarks").Value & ""
                .Col = 7: .Text = R("Status").Value
                '.Col = 8: .Text = IIf(Not IsNull(R("SLNO")), R("SLNO"), "")
        R.MoveNext
        Loop
        End With
        Set R = New ADODB.Recordset
             strRecord = "SELECT UseCheque,Prefix FROM Cheque_Master_Table WHERE BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' AND COMPID = " & bytCompID & ";"
           R.Open strRecord, MainConn
        strUseCheque = R("UseCheque").Value
        cmbIssueChq.Clear
        For i = 1 To Len(strUseCheque)
            strusecheque2 = Mid(strUseCheque, i, 1)
            If strusecheque2 = 0 Then
                cmbIssueChq.AddItem R("Prefix").Value & txtCHQNOFrom + (i - 1)
            End If
        Next i
    Else
        Call prcBlank(1)
    End If
    
    'Call prcTotalPage
    
    Set R = Nothing
    
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAdd_Click()
On Error GoTo x
    'If Len(txtIssueCHQNO.Text) = 0 Then
    If cmdAdd.Caption = "&Add" Then
        If Len(cmbIssueChq.Text) = 0 Then
            MsgBox "You should provide cheque no.", vbCritical, App.Title
            cmbIssueChq.SetFocus
            Exit Sub
        End If
    End If
    If Len(txtIssueDate.Text) = 0 Then
        MsgBox "You should provide issue date.", vbCritical, App.Title
        txtIssueDate.SetFocus
        Exit Sub
    End If
    If Len(txtBeneficiary.Text) = 0 Then
        MsgBox "You should provide beneficiary name.", vbCritical, App.Title
        txtBeneficiary.SetFocus
        Exit Sub
    End If
   
    '=======Edit Row ========='
'    If fncFindData = True Then
'            If cmdAdd.Caption = "&Replace" Then
'            Call prcSetConnectionActive
'    Mysql = "Update Cheque_Child_Table  SET DeleteRow = 2,Pcname='" & getComName & "',username='" & strUser & "',EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' WHERE IDNO='" & flexDetails.Tag & "' ;"
'    MainConn.Execute (Mysql)
'
'            End If
'    End If
    '========================-=-=++--'
   
   
   
    If cmdAdd.Caption = "&Add" Then bytRowNo = 0

    Dim strItem As String
     If cmdAdd.Caption = "&Add" Then
    With flexDetails
        If .Rows > 1 Then
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0:   strItem = .Text
                'If strItem = txtIssueCHQNO.Text And inti <> bytRowNo Then
                If strItem = cmbIssueChq.Text And inti <> bytRowNo Then
                    MsgBox "Adding Duplicate Cheque No", vbInformation, cnstMsgInfo
                    cmbIssueChq.SetFocus
                    Exit Sub
                End If
            Next inti
        End If
    End With
    End If
    
'    If FINDFANCDATA = True Then
'    findslnoMAX (flexDetails.Tag)
'    End If
'
    With flexDetails
        If boEdit = False Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0:      .Text = cmbIssueChq.Text
        Else
            .Row = bytRowNo
        End If
        
        '.Col = 0:      .Text = txtIssueCHQNO.Text
        .Col = 1:      .Text = txtIssueDate.Text
        .Col = 2:      .Text = txtCHQValue.Text
        .Col = 3:      .Text = txtBeneficiary.Text
        .Col = 4:      .Text = txtPaymentParticulars.Text
        .Col = 5:      .Text = txtEncashmentDate.Text
        .Col = 6:      .Text = txtRemarks.Text
        .Col = 7:      .Text = chkCancel.Value

    End With

     
    cmbIssueChq.Visible = True
    cmdAdd.Caption = "&Add"
    
    prcBlank (2)
    txtIssueCHQNO.SetFocus
    Exit Sub
x:
'    MainConn.RollbackTrans
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo x
    If cmdAddNew.Caption = "&Edit" Then
        cmdAddNew.Caption = "Add &New"
        cmdSave.Caption = "&Update"
        cmdDelete.Enabled = True
    Else
        cmdAddNew.Caption = "&Edit"
        cmdSave.Caption = "&Save"
        cmdDelete.Enabled = False
    End If
    
    cmbBank.SetFocus
    flexDetails.Rows = 1
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAddBank_Click()
    frmNewBank.Show 1
    Call prcAddBank(cmbBank)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo x
    Call prcSetConnectionActive
    If fncBlank(1) = True Then Exit Sub
    If MsgBox("Are you confirm to delete ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        MainConn.BeginTrans
'            strRecord = "Delete From Cheque_Master_Table Where IDNO = " & flexDetails.Tag & ";"
'            MainConn.Execute (strRecord)
            
            strRecord = "UPDATE Cheque_Master_Table SET Deleterow=1,Username='" & strUser & "',Pcname='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where IDNO = " & flexDetails.Tag & " and Deleterow=0;"
            MainConn.Execute (strRecord)
            
            strRecord = "UPDATE Cheque_Child_Table SET Deleterow=1,Username='" & strUser & "',Pcname='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where IDNO = " & flexDetails.Tag & " and Deleterow=0;"
            MainConn.Execute (strRecord)
            
            Call prcBlank(1)
            MsgBox "Deletion complete successfully. ", vbInformation, App.Title
            Call prcAddBookNo
        MainConn.CommitTrans
    End If
    Exit Sub
x:
    con.RollbackTrans
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdClose_Click()
    canInsert = 0
    canUpdate = 0
    canDelete = 0
    canPreview = 0

    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
Dim strCHQNO As String, strDate As String, curCHQValue As Currency
Dim strBeneficiary As String, strPayParticulars As String, strEncashmentDate As String
Dim intStatus As Integer, strRemarks As String
Dim intIDNO As Integer
Dim strUseCheque As String
'Dim StrUser As String
'Dim strPCName As String
    'StrUser = frmMaster.StatusBar1.Panels(1).Text
    'strPCName = GetComputerName
    Call prcSetConnectionActive
    If fncBlank(1) = True Then Exit Sub
    
    If fncFindData = False Then
        If MsgBox("Do you want to Save?.", vbInformation + vbYesNo, App.Title) = vbNo Then Exit Sub
    Else
        If MsgBox("Do you want to Update.", vbInformation + vbYesNo, App.Title) = vbNo Then Exit Sub
    End If
    strUseCheque = String(Val(txtTotalPage.Text), "0")
'    If fncBlank(1) = True Then Exit Sub
    If fncFindData = False Then
        intIDNO = fncNewID
      
        MainConn.BeginTrans
                Mysql = " Insert Into Cheque_Master_Table (IDNO, BankID, CHQBookNO, ChqStartNO, ChqEndNO, Compid,UserName,PCName,ChequeLot,UseCheque,Prefix,entrydate,entrytime)"
                Mysql = Mysql + " Values (" & intIDNO & ", " & Trim(cmbBank.Tag) & ", '" & cmbBookNo & "', '" & txtCHQNOFrom.Text & "', '" & txtCHQNOTo.Text & "', " & bytCompID & ",'" & strUser & "','" & getComName & "'," & txtTotalPage & ",'" & strUseCheque & "','" & txtPrefix.Text & "','" & Date & "','" & Time & "');"
                MainConn.Execute (Mysql)
                Call prcSaveGriddata(intIDNO)
        MainConn.CommitTrans
    Else
        intIDNO = flexDetails.Tag
        MainConn.BeginTrans
             strRecord = " Update Cheque_Child_Table set Deleterow=2,Username='" & strUser & "',Pcname='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where IDNO = " & intIDNO & " AND  DELETEROW=0;"
        
'             strRecord = " Delete FROM Cheque_Child_Table Where IDNO = " & intIDNO & " AND  DELETEROW=0;"
             MainConn.Execute (strRecord)
             strRecord = " Update Cheque_Master_Table Set BankID=" & cmbBank.Tag & ", CHQBookNO='" & cmbBookNo.Text & "', ChqStartNO='" & txtCHQNOFrom.Text & "', ChqEndNO='" & txtCHQNOTo.Text & "',UserName = '" & strUser & "', PCName = '" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "',ChequeLot = " & txtTotalPage.Text & ",UseCheque = '" & strUseCheque & "',Prefix = '" & txtPrefix.Text & "'"
             strRecord = strRecord + " Where Compid = " & bytCompID & " and IDNO=" & intIDNO & " and Deleterow=0;"
             MainConn.Execute (strRecord)
            Call prcSaveGriddata(intIDNO)
       MainConn.CommitTrans
    End If
    Call prcBlank(1)
'    cmdSave.Enabled = False
    MsgBox "Data updated successfully.", vbInformation, App.Title
    Exit Sub
x:
     MainConn.RollbackTrans
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub flexDetails_DblClick()
On Error Resume Next
    flexDetails_KeyDown 13, 0
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With flexDetails
            If .RowSel = 0 Then Exit Sub
            .Col = 0:    txtIssueCHQNO.Text = .Text
            .Col = 1:    txtIssueDate.Text = .Text
            .Col = 2:    txtCHQValue.Text = .Text
            .Col = 3:    txtBeneficiary.Text = .Text
            .Col = 4:    txtPaymentParticulars.Text = .Text
            .Col = 5:    txtEncashmentDate.Text = .Text
            .Col = 6:    txtRemarks.Text = .Text
            .Col = 7:    chkCancel.Value = .Text
'            .Col = 8:    txtSL = .Text
            bytRowNo = .Row
        End With
         boEdit = True
        cmbIssueChq.Visible = False
        cmdAdd.Caption = "&Replace"
        txtIssueCHQNO.SetFocus
    End If

    If KeyCode = vbKeyDelete Then
        With flexDetails

            If .RowSel = 0 Then
                Beep
            Else
                If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
                If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
            End If

            
            
            
            
        End With
        cmbIssueChq.Visible = True
        cmdAdd.Caption = "&Add"
    End If
End Sub


Private Sub Form_Load()
On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    Call prcGridSetting
    Call prcAddBank(cmbBank)
    
    Set R = New ADODB.Recordset
    strRecord = "Select DelStatus,UpStatus from InventoryPJL.dbo.Sys_User_Name where UserName='" & strUser & "'"
    R.Open strRecord, MainConn, adOpenStatic
    If R!DelStatus = 1 Then
       cmdDelete.Enabled = True
    End If

''''12-jan-11''''''    Call prcButtonPermission(Me.Name)
   ''''15-jan-11''''''   If canInsert = 0 Then cmdSave.Enabled = False Else cmdSave.Enabled = True
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub txtBeneficiary_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBeneficiary_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtBeneficiary.ToolTipText = txtBeneficiary.Text
End Sub

Private Sub txtCHQNOFrom_GotFocus()
    Call txtGotFocus(txtCHQNOFrom)
End Sub

Private Sub txtCHQNOFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCHQNOFrom_LostFocus()
    Call prcTotalPage
End Sub

Private Sub txtCHQNOTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCHQNOTo_LostFocus()
    Call prcTotalPage
End Sub

Private Sub txtCHQValue_GotFocus()
    txtCHQValue.SelStart = 0
    txtCHQValue.SelLength = Len(txtCHQValue.Text)
End Sub

Private Sub txtCHQValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCHQValue_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtCHQValue_LostFocus()
    If Len(txtCHQValue.Text) > 0 Then txtCHQValue.Text = FormatNumber(txtCHQValue.Text, 0)
End Sub

Private Sub txtEncashmentDate_GotFocus()
    txtEncashmentDate.SelStart = 0
    txtEncashmentDate.SelLength = Len(txtEncashmentDate.Text)
End Sub

Private Sub txtEncashmentDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtEncashmentDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtEncashmentDate_LostFocus()
Dim dtIssueDate As Date
Dim dtEncashmentDate As Date
    Call FormatDate(txtEncashmentDate)
    If txtEncashmentDate.Text <> "" Then
        dtIssueDate = txtIssueDate.Text
        dtEncashmentDate = txtEncashmentDate.Text
        If dtIssueDate > dtEncashmentDate Then
            MsgBox "Honor Date should be greater than Issue Date ", vbCritical
        End If
    End If
End Sub

Private Sub txtIssueCHQNO_GotFocus()
    txtIssueCHQNO.SelStart = 0
    txtIssueCHQNO.SelLength = Len(txtIssueCHQNO.Text)
End Sub

Private Sub txtIssueCHQNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtIssueDate_GotFocus()
    txtIssueDate.SelStart = 0
    txtIssueDate.SelLength = Len(txtIssueDate.Text)
End Sub

Private Sub txtIssueDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtIssueDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtIssueDate_LostFocus()
    Call FormatDate(txtIssueDate)
End Sub

Private Sub txtPaymentParticulars_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPaymentParticulars_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtPaymentParticulars.ToolTipText = txtPaymentParticulars.Text
End Sub

Private Sub txtPrefix_GotFocus()
    Call txtGotFocus(txtPrefix)
End Sub

Private Sub txtPrefix_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPrefix_LostFocus()
    'txtCHQNOFrom.Text = txtPrefix.Text & txtCHQNOFrom.Text
    'txtCHQNOTo.Text = txtPrefix.Text & txtCHQNOTo.Text
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtRemarks_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtRemarks.ToolTipText = txtRemarks.Text
End Sub

Private Sub txtTotalPage_Click()
'With flexDetails
'If .Row > 0 Then
'intGMaxSLNO = findslnoMAX(flexDetails.Tag)
'End If
'End With

End Sub

Private Sub txtTotalPage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub prcAddBookNo()
On Error GoTo x
    Call prcSetConnectionActive
    Set R = New ADODB.Recordset
        strRecord = "Select chqBookNO From Cheque_Master_Table Where BankID = " & cmbBank.Tag & "  AND COMPID = " & bytCompID & " and Deleterow=0 Order By chqBookNO;"
   R.Open strRecord, MainConn, adOpenStatic
        
        cmbBookNo.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmbBookNo.AddItem R("CHQBOOKNO")
            R.MoveNext
            Loop
        End If
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub prcBlank(Param As Integer)
''    If param = 1 Then
''        txtCHQNOFrom.Text = ""
''        txtCHQNOTo.Text = ""
''        txtTotalPage.Text = ""
''
''        flexDetails.Rows = 1
''    End If
    
    txtIssueCHQNO.Text = ""
    txtIssueDate.Text = ""
    txtCHQValue.Text = ""
    txtBeneficiary.Text = ""
    txtPaymentParticulars.Text = ""
    txtEncashmentDate.Text = ""
    txtRemarks.Text = ""
    chkCancel.Value = 0
    flexDetails.Rows = 1
End Sub

Private Function fncBlank(Param As Integer) As Boolean
    fncBlank = True
        If Param = 1 Then
            If Trim(cmbBank.Text) = "" Then
                MsgBox "You should provide issue bank name", vbCritical, App.Title
                cmbBank.SetFocus
                Exit Function
            End If
            If Trim(cmbBookNo.Text) = "" Then
                MsgBox "You should provide issue book no", vbCritical, App.Title
                cmbBookNo.SetFocus
                Exit Function
            End If
            If flexDetails.Rows = 1 Then
                MsgBox "Data not found in the grid to save", vbCritical, App.Title
                cmdSave.SetFocus
                Exit Function
            End If
        End If
    fncBlank = False
End Function
Public Sub prcGridSetting()
    With flexDetails
        .Rows = 1
        .Cols = 8
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignCenterCenter
        '.ColAlignment(8) = flexAlignCenterCenter
        
        .Col = 0:   .Text = "ISSUED CHQ #"
        .Col = 1:   .Text = "ISSUED Date"
        .Col = 2:   .Text = "CHQ VALUE"
        .Col = 3:   .Text = "BENEFICIARY"
        .Col = 4:   .Text = "PAYMENT PARTICULARS"
        .Col = 5:   .Text = "DATE OF HONOUR"
        .Col = 6:   .Text = "REMARKS"
        .Col = 7:   .Text = "STATUS"
        '.Col = 8: .Text = "SLNO"
        
        
        .ColWidth(0) = 1800:     .ColWidth(1) = 1200
        .ColWidth(2) = 1200:     .ColWidth(3) = 3000
        .ColWidth(4) = 2500:     .ColWidth(5) = 2500
        .ColWidth(6) = 2500:     .ColWidth(7) = 800
        '.ColWidth(8) = 900
    End With
End Sub

Private Function fncFindData() As Boolean
    Set R = New ADODB.Recordset
        strRecord = "Select Count(*) AS RC From Cheque_Master_Table Where BankID = " & cmbBank.Tag & " and CHQBookNO = '" & cmbBookNo.Text & "' and Compid = " & bytCompID & ";"
       R.Open strRecord, MainConn, adOpenStatic
        
        If R("RC") > 0 Then
            fncFindData = True
        Else
            fncFindData = False
        End If
    Set R = Nothing
End Function

Private Function fncFindID() As Integer
    Set R = New ADODB.Recordset
        Mysql = "Select IDNO from Cheque_Master_Table Where BankID = " & cmbBank.Tag & " and ChqBookNo = '" & Trim(cmbBookNo.Text) & "' and Compid = " & bytCompID & ";"
        com.CommandText = Mysql
        Set R = com.Execute
        
        If R.EOF = False And R.BOF = False Then
            fncFindID = R("IDNO")
        End If
    Set R = Nothing
End Function

Private Function fncNewID() As Integer
    Set R = New ADODB.Recordset
        Mysql = "Select max(IDNO) as maxID from Cheque_Master_Table;"
    R.Open Mysql, MainConn, adOpenStatic
        
        If R.EOF = False And R.BOF = False Then
            If IsNull(R("maxID")) = False Then
                fncNewID = R("maxid") + 1
            Else
                fncNewID = 1
            End If
        Else
            fncNewID = 1
        End If
    Set R = Nothing
End Function

Private Sub prcTotalPage()
Dim intFrom As Double, intTo As Double
    If Len(txtTotalPage.Text) = 0 Then Exit Sub
    If Len(txtCHQNOFrom.Text) = 0 Then Exit Sub
    'If Len(txtCHQNOTo.Text) = 0 Then Exit Sub
    'intFrom = txtCHQNOFrom.Text
    'intTo = txtCHQNOTo.Text
    'txtTotalPage.Text = CStr((intTo - intFrom) + 1)
    txtCHQNOTo.Text = (Val(txtCHQNOFrom.Text) + Val(txtTotalPage.Text)) - 1
    cmbBookNo.Text = txtPrefix.Text + txtCHQNOFrom.Text + "-" + txtPrefix.Text + txtCHQNOTo.Text
    ''''''''''
    Set R = New ADODB.Recordset
    strRecord = "Select chqBookNO From Cheque_Master_Table Where chqStartNO ='" & txtCHQNOFrom.Text & "' And BankID = " & cmbBank.Tag & " AND COMPID = " & bytCompID & ";"
    R.Open strRecord, MainConn
    If R.EOF = False And R.BOF = False Then
    Else
        cmbIssueChq.Clear
        For i = 0 To Val(txtTotalPage.Text) - 1
            cmbIssueChq.AddItem txtPrefix.Text & Val(txtCHQNOFrom.Text) + i
        Next i
    End If
End Sub

Private Sub prcSaveGriddata(intIDNO As Integer)
'Dim StrUser As String
'Dim strPCName As String
    'StrUser = frmMaster.StatusBar1.Panels(1).Text
    'strPCName = GetComputerName
    
    'tkc5 = Len(txtTotalPage.Text)
    strUseCheque = String(Val(txtTotalPage.Text), "0")
    With flexDetails
        For inti = 1 To .Rows - 1
            .Row = inti
            .Col = 0: strCHQNO = .Text
            .Col = 1:
                If Len(Trim(.Text)) = 0 Then
                    strDate = Format("0", cnstDtFrmtI)
                Else
                    strDate = .Text
                End If
            .Col = 2
                If .Text = "" Then
                    curCHQValue = 0
                Else
                    curCHQValue = CDbl(.Text)
                End If
            .Col = 3:   strBeneficiary = Trim(.Text)
            .Col = 4:   strPayParticulars = Trim(.Text)
            .Col = 5:
                If Len(Trim(.Text)) = 0 Then
                    strEncashmentDate = ""
                Else
                    strEncashmentDate = .Text
                End If
            .Col = 6:   strRemarks = .Text
            .Col = 7:   strStatus = .Text
            '.Col = 8: intSLNO = .Text
            intIssueTicket = (Val(strCHQNO) - Val(txtCHQNOFrom.Text)) + 1
            strUseCheque = Mid(strUseCheque, 1, intIssueTicket - 1) & "1" & Mid(strUseCheque, intIssueTicket + 1)
            If Len(strEncashmentDate) = 0 Then
                Mysql = " Insert  Into Cheque_Child_Table (IDNO, ChequeNO, IssueDate, ChequeValue, Beneficiary, PaymentParticulars, Remarks, Status,UserName,PCName,entrydate,entrytime)"
                Mysql = Mysql + " values (" & intIDNO & ", '" & strCHQNO & "', '" & strDate & "', " & curCHQValue & ", '" & strBeneficiary & "', '" & strPayParticulars & "', '" & strRemarks & "', '" & strStatus & "','" & strUser & "','" & getComName & "','" & Date & "','" & Time & "');"
            Else
                Mysql = " Insert  Into Cheque_Child_Table (IDNO, ChequeNO, IssueDate, ChequeValue, Beneficiary, PaymentParticulars, EncashmentDate, Remarks, Status,UserName,PCName,entrydate,entrytime)"
                Mysql = Mysql + " values (" & intIDNO & ", '" & strCHQNO & "', '" & strDate & "', " & curCHQValue & ", '" & strBeneficiary & "', '" & strPayParticulars & "', '" & strEncashmentDate & "', '" & strRemarks & "', '" & strStatus & "','" & strUser & "','" & getComName & "','" & Date & "','" & Time & "');"
            End If
            
            MainConn.Execute (Mysql)
            '''''''''
            strRecord = " Update Cheque_Master_Table Set UseCheque = '" & strUseCheque & "'"
            strRecord = strRecord + " Where Compid = " & bytCompID & " and IDNO=" & intIDNO & ";"
            MainConn.Execute (strRecord)
            ''''''''
        Next inti
            
    End With
End Sub
