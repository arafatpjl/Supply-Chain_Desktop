VERSION 5.00
Begin VB.Form frmrptChequeInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5941
   ClientLeft      =   13
   ClientTop       =   13
   ClientWidth     =   7098
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
   ForeColor       =   &H8000000E&
   Icon            =   "frmrptChequeInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5941
   ScaleWidth      =   7098
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbCompanyName 
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2550
      Width           =   4485
   End
   Begin VB.TextBox txtToDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Top             =   4470
      Width           =   1965
   End
   Begin VB.TextBox txtFromDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1110
      TabIndex        =   8
      Top             =   4470
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SELECT CRITERIA"
      Height          =   1215
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   6795
      Begin VB.OptionButton optTransactionWise 
         BackColor       =   &H80000004&
         Caption         =   "TRANSACTION WISE"
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   1995
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H80000004&
         Caption         =   "DATE WISE"
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   1890
         TabIndex        =   1
         Top             =   210
         Width           =   1485
      End
      Begin VB.OptionButton optChqPay 
         BackColor       =   &H80000004&
         Caption         =   "CHQ PAYMENT"
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   4080
         TabIndex        =   3
         Top             =   600
         Width           =   1635
      End
      Begin VB.OptionButton optChqCancel 
         BackColor       =   &H80000004&
         Caption         =   "CHQ CANCEL"
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton optChqTran 
         BackColor       =   &H80000004&
         Caption         =   "CHQ TRANSFER"
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1665
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H80000004&
         Caption         =   "&ALL"
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optCash 
         BackColor       =   &H80000004&
         Caption         =   "&CASH DRAWING"
         ForeColor       =   &H00E26101&
         Height          =   195
         Left            =   3360
         TabIndex        =   2
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   525
      Left            =   5400
      TabIndex        =   11
      Top             =   5160
      Width           =   1185
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   525
      Left            =   3960
      TabIndex        =   10
      Top             =   5160
      Width           =   1185
   End
   Begin VB.ComboBox cmbBookNo 
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1110
      TabIndex        =   7
      Top             =   3870
      Width           =   4485
   End
   Begin VB.ComboBox cmbBank 
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   4485
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: Cheque Information Report"
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
      Left            =   0
      TabIndex        =   23
      Top             =   120
      Width           =   4665
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   7800
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   21
      Top             =   5040
      Width           =   7800
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Name"
      ForeColor       =   &H00E26101&
      Height          =   255
      Left            =   1140
      TabIndex        =   20
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TO ISSUE DATE "
      ForeColor       =   &H00E26101&
      Height          =   225
      Left            =   3600
      TabIndex        =   17
      Top             =   4230
      Width           =   1395
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM ISSUE DATE "
      ForeColor       =   &H00E26101&
      Height          =   225
      Left            =   1110
      TabIndex        =   16
      Top             =   4230
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   45
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   4980
      Width           =   7815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHQ BOOK # :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   1
      Left            =   1110
      TabIndex        =   13
      Top             =   3630
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUING BANK :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   0
      Left            =   1110
      TabIndex        =   12
      Top             =   2970
      Width           =   1245
   End
End
Attribute VB_Name = "frmrptChequeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBank_Click()
    On Error GoTo x
    cmbBank.Tag = 0
    cmbBookNo.Clear:    cmbBookNo.Text = ""
    If cmbCompanyName.Text = "" Then
        MsgBox "Plese! Select Company Name", vbCritical
        cmbCompanyName.SetFocus
        Exit Sub
    End If
    If cmbBank.ListIndex >= 0 Then
        cmbBank.Tag = cmbBank.ItemData(cmbBank.ListIndex)
        Call prcAddBookNo
    End If
    
    cmbBank.ToolTipText = cmbBank.Text
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbBank_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbBookNo_Click()
    cmbBookNo.ToolTipText = cmbBookNo.Text
End Sub

Private Sub cmbBookNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbCompanyName_Click()
    cmbCompanyName.Tag = cmbCompanyName.ItemData(cmbCompanyName.ListIndex)
    cmbBookNo.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
On Error GoTo x
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set RS = New ADODB.Recordset

        If optAll.Value = True Then
            If fncBlank(1) = True Then Exit Sub
            
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfo.rpt")
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' Order By ChequeNo;"

        ElseIf optDate.Value = True Then
            If fncBlank(1) = True Then Exit Sub
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfoDate.rpt")

            If Len(cmbBookNo.Text) = 0 Then
                strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " and IssueDate Between '" & txtFromDate.Text & "' and '" & txtToDate.Text & "' Order By ChequeNo;"
            Else
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' and IssueDate Between '" & txtFromDate.Text & "' and '" & txtToDate.Text & "'  Order By ChequeNo;"
            End If

        ElseIf optCash.Value = True Then
            If fncBlank(1) = True Then Exit Sub
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfo.rpt")
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' and LEFT(Beneficiary,4) = 'CASH' Order By ChequeNo;"
       
        ElseIf optChqPay.Value = True Then
            If fncBlank(1) = True Then Exit Sub
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfo.rpt")
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' AND ((LEFT(BENEFICIARY,4) <> 'CASH') OR (LEFT(BENEFICIARY,8) <> 'TRANSFER')) Order By ChequeNo;"
                    
        ElseIf optChqTran.Value = True Then
            If fncBlank(1) = True Then Exit Sub
            
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfo.rpt")
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "'  and LEFT(Beneficiary, 8) = 'TRANSFER' Order By ChequeNo;"
        
        ElseIf optChqCancel.Value = True Then
            If fncBlank(1) = True Then Exit Sub
            
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfo.rpt")
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' AND Status = 1 Order By ChequeNo;"
            
        ElseIf optTransactionWise.Value = True Then
            
            Set Report = Appl.OpenReport(findReportPath & "Cheque\rptChequeInfoTransactionWise.rpt")
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " AND BANKID = " & cmbBank.Tag & " AND CHQBOOKNO = '" & cmbBookNo.Text & "' AND Status = 1 Order By ChequeNo;"
        
            If Len(cmbCompanyName.Text) <> 0 And Len(cmbBank.Text) <> 0 And Len(cmbBookNo.Text) <> 0 Then
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & cmbCompanyName.Tag & " AND BANKID = " & cmbBank.Tag & " AND chqBookNo = '" & cmbBookNo.Text & "' And IssueDate Between '" & txtFromDate.Text & "' and '" & txtToDate.Text & "' Order By COMPID,BANKID,chqBookNo,ChequeNo;"

            ElseIf Len(cmbCompanyName.Text) <> 0 And Len(cmbBank.Text) <> 0 And Len(cmbBookNo.Text) = 0 Then
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & cmbCompanyName.Tag & " AND BANKID = " & cmbBank.Tag & " And IssueDate Between '" & txtFromDate.Text & "' and '" & txtToDate.Text & "' Order By ChequeNo;"

            ElseIf Len(cmbCompanyName.Text) <> 0 And Len(cmbBank.Text) = 0 And Len(cmbBookNo.Text) = 0 Then
            strRecord = "SELECT * FROM vieChequeInfo WHERE COMPID = " & bytCompID & " And IssueDate Between '" & txtFromDate.Text & "' And '" & txtToDate.Text & "' Order By COMPID,BANKID,chqBookNo,ChequeNo;"

            Else
            strRecord = "SELECT * FROM vieChequeInfo WHERE IssueDate Between '" & txtFromDate.Text & "' and '" & txtToDate.Text & "' Order By COMPID,BANKID,chqBookNo,ChequeNo;"

            End If
            
            Report.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, "DD-MMM-YYYY") & "'"
            Report.FormulaFields(4).Text = "'" & Format(txtToDate.Text, "DD-MMM-YYYY") & "'"
           
        End If
            
    RS.Open strRecord, MainConn, adOpenStatic
    Report.DiscardSavedData
    Report.Database.SetDataSource RS
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    RS.Close

Exit Sub
x:
'Screen.MousePointer = 0
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call prcLoadCompany(cmbCompanyName)
    Call prcAddBank(cmbBank)
'    Call prcConReport(CrystalReport1)
    Call prcButtonPermission(Me.Name)
'''    If canPreview = 0 Then cmdPreview.Enabled = False
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub prcAddBookNo()
    On Error GoTo x
    Set R = New ADODB.Recordset
        'strRecord = "Select chqBookNO From Cheque_Master_Table Where BankID = " & cmbBank.Tag & ";"
        strRecord = "Select chqBookNO From Cheque_Master_Table Where CompID = " & cmbCompanyName.Tag & " And BankID = " & cmbBank.Tag & ";"
'        com.CommandText = strRecord
'        Set R = com.Execute
        R.Open strRecord, MainConn, adOpenStatic
        cmbBookNo.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmbBookNo.AddItem R("CHQBOOKNO")
            R.MoveNext
            Loop
        End If
    Set R = Nothing
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Public Function fncBlank(intParam As Integer) As Boolean
fncBlank = True
    If intParam = 0 Then
        If Len(cmbCompanyName.Text) = 0 Then
            MsgBox "You should provide Company name.", vbCritical, App.Title
            cmbCompanyName.SetFocus
            Exit Function
        End If
        If Len(cmbBank.Text) = 0 Then
            MsgBox "You should provide Bank name.", vbCritical, App.Title
            cmbBank.SetFocus
            Exit Function
        End If
        If optDate.Value = False Then
            If cmbBookNo.Text = "" Then
                MsgBox "You should provide Cheque Book NO.", vbCritical, App.Title
                cmbBookNo.SetFocus
                Exit Function
            End If
        End If
    Else
    
    End If
fncBlank = False
End Function
    
Private Sub optAll_Click()
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optAll_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optCash_Click()
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optCash_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optChqCancel_Click()
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optChqCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optChqPay_Click()
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optChqPay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optChqTran_Click()
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optChqTran_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optDate_Click()
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optTransactionWise_Click()
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub txtFromDate_GotFocus()
    txtFromDate.SelStart = 0
    txtFromDate.SelLength = Len(txtFromDate.Text)
End Sub

Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFromDate_LostFocus()
    Call FormatDate(txtFromDate)
End Sub

Private Sub txtToDate_GotFocus()
    txtToDate.SelStart = 0
    txtToDate.SelLength = Len(txtToDate.Text)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtToDate_LostFocus()
    Call FormatDate(txtToDate)
End Sub
