VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrptTransactionWiseReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Wise Report"
   ClientHeight    =   5252
   ClientLeft      =   39
   ClientTop       =   429
   ClientWidth     =   6786
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmrptTransactionWiseReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5252
   ScaleWidth      =   6786
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4225
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4605
      Width           =   1155
   End
   Begin VB.ComboBox cmbStatementNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2985
      Width           =   1185
   End
   Begin VB.ComboBox cmbPreDate 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2520
      Width           =   2595
   End
   Begin VB.CommandButton cmdFindDate 
      BackColor       =   &H00C0C000&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   585
   End
   Begin VB.ComboBox cmbCashAndCredit 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmrptTransactionWiseReport.frx":000C
      Left            =   2580
      List            =   "frmrptTransactionWiseReport.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   2595
   End
   Begin VB.CheckBox chkAllCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Company"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2565
      TabIndex        =   7
      Top             =   3750
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker DTPickerToDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   6
      Top             =   1995
      Width           =   1935
      _ExtentX        =   3403
      _ExtentY        =   671
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.5094
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124583937
      CurrentDate     =   39284
   End
   Begin MSComCtl2.DTPicker DTPickerFromDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   5
      Top             =   1530
      Width           =   1935
      _ExtentX        =   3403
      _ExtentY        =   671
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.5094
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124583937
      CurrentDate     =   39284
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4590
      Width           =   1245
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5505
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4590
      Width           =   1125
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3330
      Top             =   4500
      _ExtentX        =   671
      _ExtentY        =   671
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   4440
      Width           =   6795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: Transaction Wise "
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
      TabIndex        =   16
      Top             =   105
      Width           =   3480
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   17
      Top             =   -120
      Width           =   6795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   825
      TabIndex        =   15
      Top             =   2565
      Width           =   1620
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement No.  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1065
      TabIndex        =   13
      Top             =   3015
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   9
      Top             =   1095
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   39
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   4407
      Width           =   6799
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Index           =   2
      Left            =   645
      TabIndex        =   3
      Top             =   1620
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Index           =   1
      Left            =   630
      TabIndex        =   2
      Top             =   2055
      Width           =   1785
   End
End
Attribute VB_Name = "frmrptTransactionWiseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPreDate_Click()
    If cmbPreDate.Text = "" Then
    Else
        Set R = New Recordset
            sql1 = "Select CompanyID, Tran_Type, Tran_Date, StatementNo From Statement_of_Cash_and_Credit Group By CompanyID,Tran_Type,Tran_Date,StatementNo HAVING CompanyID= " & bytCompID & " And Tran_Type='" & cmbCashAndCredit.Text & "' And Tran_Date='" & cmbPreDate.Text & "' Order By Tran_Date"
            R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            cmbStatementNo.Clear
            Do While Not R.EOF
                cmbStatementNo.AddItem R("StatementNo")
                R.MoveNext
            Loop
        Set R = Nothing
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFindDate_Click()
    If cmbCashAndCredit.Text = "" Then
        MsgBox "Please type Transaction Type.", vbCritical, App.Title
    Else
        Set R = New Recordset
            sql1 = "Select CompanyID, Tran_Type, Tran_Date From Statement_of_Cash_and_Credit Group By CompanyID,Tran_Type,Tran_Date HAVING CompanyID= " & bytCompID & " And Tran_Type='" & cmbCashAndCredit.Text & "' And Tran_Date>='" & DTPickerFromDate.Value & "'And Tran_Date<='" & DTPickerToDate.Value & "' Order By Tran_Date"
            R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            cmbPreDate.Clear
            cmbStatementNo.Clear
            Do While Not R.EOF
                cmbPreDate.AddItem R("Tran_Date")
                R.MoveNext
            Loop
        Set R = Nothing
    End If
End Sub

Private Sub cmdPreview_Click()
'    MousePointer = 11
    Dim SQLQuery As String
        If chkAllCompany.Value = 0 Then
           
             If cmbCashAndCredit.Text = "All Type" Then
                If cmbPreDate.Text = "" Then
                    SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                        & " From viewStatement_Cash_Credit " _
                        & " Where CompanyID= " & bytCompID & " AND Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' " _
                        & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                     
                Else
                    If cmbStatementNo.Text = "" Then
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where companyID= " & intCompanyID & " AND Tran_Date ='" & cmbPreDate.Text & "' " _
                            & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                    Else
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where CompanyID= " & intCompanyID & " AND Tran_Date ='" & cmbPreDate.Text & "' And StatementNo = " & cmbStatementNo.Text & " " _
                            & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                    End If
                End If
            Else
                If cmbPreDate.Text = "" Then
                    SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where CompanyID= " & intCompanyID & " AND Tran_Type = '" & cmbCashAndCredit.Text & "' AND Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' " _
                            & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                Else
                    If cmbStatementNo.Text = "" Then
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where CompanyID= " & intCompanyID & " AND Tran_Type = '" & cmbCashAndCredit.Text & "' AND Tran_Date = '" & cmbPreDate.Text & "' " _
                            & " ORDER BY Tran_Type, StatementNo,Tran_Date;"
                    Else
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where CompanyID= " & intCompanyID & " AND Tran_Type = '" & cmbCashAndCredit.Text & "' AND Tran_Date = '" & cmbPreDate.Text & "' And StatementNo = " & cmbStatementNo.Text & "" _
                            & " ORDER BY Tran_Type, StatementNo,Tran_Date;"
                    End If
                End If
            End If
           
            Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptTransactionWise.rpt")
            Report.FormulaFields(2).Text = "'Date From : " & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & "  To " & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & "'"

        ElseIf chkAllCompany.Value = 1 Then

            If cmbCashAndCredit.Text = "All Type" Then
                If cmbPreDate.Text = "" Then
                    SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                        & " From viewStatement_Cash_Credit " _
                        & " Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' " _
                        & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                Else
                    If cmbStatementNo.Text = "" Then
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where Tran_Date ='" & cmbPreDate.Text & "' " _
                            & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                    Else
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where Tran_Date ='" & cmbPreDate.Text & "' And StatementNo = " & cmbStatementNo.Text & " " _
                            & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                    End If
                End If
            Else
                If cmbPreDate.Text = "" Then
                    SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where Tran_Type = '" & cmbCashAndCredit.Text & "' AND Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' " _
                            & " ORDER BY Tran_Type,StatementNo,Tran_Date;"
                Else
                    If cmbStatementNo.Text = "" Then
                        strRecord = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where Tran_Type = '" & cmbCashAndCredit.Text & "' AND Tran_Date = '" & cmbPreDate.Text & "' " _
                            & " ORDER BY Tran_Type, StatementNo,Tran_Date;"
                    Else
                        SQLQuery = "SELECT CompanyID, CompName, CompAdd, Tr_No, Tran_Type, Tran_Date, NameOfTheParty, BillNo, BillDate, MRRNo, MRRDate, Amount, Remark, StatementNo, SupplierName, Chalan_Date, Chalan_No " _
                            & " From viewStatement_Cash_Credit " _
                            & " Where Tran_Type = '" & cmbCashAndCredit.Text & "' AND Tran_Date = '" & cmbPreDate.Text & "' And StatementNo = " & cmbStatementNo.Text & "" _
                            & " ORDER BY Tran_Type, StatementNo,Tran_Date;"
                    End If
                End If
            End If

            Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptTransactionWise.rpt")
            Report.FormulaFields(1).Text = "'All Company'"
            Report.FormulaFields(2).Text = "'Date From : " & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & "  To " & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & "'"

        End If
        
    Set R = New ADODB.Recordset
    R.Open SQLQuery, MainConn, adOpenStatic
    
    Report.DiscardSavedData
    Report.Database.SetDataSource R
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    
    
    

    R.Close
    'End With
 '   MousePointer = 0
    Exit Sub
x:
  '  MousePointer = 0
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub Command1_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call prcConReport(CrystalReport1)
    DTPickerFromDate.Value = Date
    DTPickerToDate.Value = Date
    cmbCashAndCredit.Clear
    cmbPreDate.Clear
    cmbStatementNo.Clear
    cmbCashAndCredit.AddItem "All Type"
    cmbCashAndCredit.AddItem "Cash"
    cmbCashAndCredit.AddItem "Credit"
    cmbCashAndCredit.AddItem "Cheque"
    'cmbCashAndCredit.AddItem "Cheque/Cash"
    cmbCashAndCredit.AddItem "Cash Cheque"
        
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub
