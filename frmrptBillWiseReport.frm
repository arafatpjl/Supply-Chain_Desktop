VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrptBillWiseReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Wise Report"
   ClientHeight    =   4940
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
   Icon            =   "frmrptBillWiseReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4940
   ScaleWidth      =   6786
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
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
      TabIndex        =   10
      Top             =   2970
      Width           =   3375
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
      Left            =   2550
      TabIndex        =   9
      Top             =   3600
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker DTPickerToDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   1830
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
      Format          =   232980481
      CurrentDate     =   39284
   End
   Begin MSComCtl2.DTPicker DTPickerFromDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   1260
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
      Format          =   232980481
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
      Left            =   169
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4230
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
      Left            =   5447
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4230
      Width           =   1105
   End
   Begin VB.ComboBox cmbBillNo 
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
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2490
      Width           =   3375
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3330
      Top             =   4260
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
      TabIndex        =   14
      Top             =   4080
      Width           =   6795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: Bill Wise "
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
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Bill No :"
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
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
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
      Height          =   45
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   4500
      Width           =   6795
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
      Height          =   285
      Index           =   2
      Left            =   630
      TabIndex        =   5
      Top             =   1350
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
      Height          =   285
      Index           =   1
      Left            =   630
      TabIndex        =   4
      Top             =   1890
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Bill No :"
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
      Height          =   285
      Index           =   0
      Left            =   630
      TabIndex        =   3
      Top             =   2490
      Width           =   1785
   End
End
Attribute VB_Name = "frmrptBillWiseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbBillNo_Click()
    txtBillNo.Text = cmbBillNo.Text
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    'MousePointer = 11
    
    If txtBillNo.Text = "" Then
        MsgBox "You should provide BILL No.", vbCritical, App.Title
        txtBillNo.SetFocus
        Exit Sub
    End If
    
        
    If chkAllCompany.Value = 0 Then
    Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptBillWise.rpt")
    strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And BILLNo = '" & txtBillNo.Text & "';"
    
    Else
     Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptBillWise.rpt")
    strRecord = "Select * From viewStatement_Cash_Credit Where BillNo = '" & txtBillNo.Text & "';"
    Report.FormulaFields(1).Text = "'All Company'"
    
    End If
        
    Set R = New ADODB.Recordset
    R.Open strRecord, MainConn, adOpenStatic
    
    Report.DiscardSavedData
    Report.Database.SetDataSource R
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = Report
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    R.Close
    
    'MousePointer = 0
    Exit Sub
x:
    'MousePointer = 0
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub DTPickerFromDate_Change()
    Set R = New Recordset
        sql1 = "Select BillNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By BillNo Order By BillNo"
        R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
        While Not R.EOF
            cmbBillNo.AddItem R.Fields("BILLNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub DTPickerFromDate_Click()
    Set R = New Recordset
        sql1 = "Select BillNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By BillNo Order By BillNo"
        R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
        While Not R.EOF
            cmbBillNo.AddItem R.Fields("BillNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub DTPickerToDate_Change()
     Set R = New Recordset
        sql1 = "Select BillNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By BillNo Order By BillNo"
        R.Open sql1, MainConn, adOpenStatic
        While Not R.EOF
            cmbBillNo.AddItem R.Fields("BILLNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub DTPickerToDate_Click()
     Set R = New Recordset
        sql1 = "Select BillNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By BillNo Order By BillNo"
        R.Open sql1, MainConn, adOpenStatic
        While Not R.EOF
            cmbBillNo.AddItem R.Fields("BillNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call prcConReport(CrystalReport1)
    DTPickerFromDate.Value = Date
    DTPickerToDate.Value = Date
    
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
