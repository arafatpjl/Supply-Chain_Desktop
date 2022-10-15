VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrptMonthlyStatement 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Statement"
   ClientHeight    =   5187
   ClientLeft      =   39
   ClientTop       =   429
   ClientWidth     =   6786
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
   Icon            =   "frmrptMonthlyStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5187
   ScaleWidth      =   6786
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbQueryType 
      Height          =   315
      Left            =   2565
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1290
      Width           =   3405
   End
   Begin VB.CheckBox chkAllCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Company"
      Height          =   225
      Left            =   2550
      TabIndex        =   9
      Top             =   3645
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker DTPickerToDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   2895
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
      Format          =   288423937
      CurrentDate     =   39284
   End
   Begin MSComCtl2.DTPicker DTPickerFromDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   2355
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
      Format          =   288423937
      CurrentDate     =   39284
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Close"
      Height          =   465
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Preview"
      Height          =   465
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1125
   End
   Begin VB.ComboBox cmbStatementType 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1830
      Width           =   3375
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1920
      Top             =   4440
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
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   4320
      Width           =   7800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: Monthly Statement"
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
      TabIndex        =   12
      Top             =   120
      Width           =   3660
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7800
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Query Type  :"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   945
      TabIndex        =   11
      Top             =   1335
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   4140
      Width           =   6795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From Date  :"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   2
      Left            =   630
      TabIndex        =   5
      Top             =   2445
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   630
      TabIndex        =   4
      Top             =   2955
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Type  :"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   0
      Left            =   630
      TabIndex        =   3
      Top             =   1830
      Width           =   1785
   End
End
Attribute VB_Name = "frmrptMonthlyStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    MousePointer = 11
    
    If cmbStatementType.Text = "" Then
        MsgBox "You should provide Statement Type.", vbCritical, App.Title
        cmbStatementType.SetFocus
        Exit Sub
    End If
            
        If chkAllCompany.Value = 0 Then
                Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptMonthlyStatement.rpt")
                strRecord = "Select * From viewMonthlyStatement Where CompanyID = " & bytCompID & " And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "';"
                Report.FormulaFields(4).Text = "'" & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & "'"
                Report.FormulaFields(5).Text = "'" & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & "'"

        Else
                Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptMonthlyStatement.rpt")
                strRecord = "Select * From viewMonthlyStatement Where Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "';"
                Report.FormulaFields(6).Text = "'All Company'"
                Report.FormulaFields(4).Text = "'" & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & "'"
                Report.FormulaFields(5).Text = "'" & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & "'"
    
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
        
'    End With
    MousePointer = 0
    Exit Sub
x:
    MousePointer = 0
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Call prcConReport(CrystalReport1)
    cmbStatementType.AddItem "Cash"
    cmbStatementType.AddItem "Cheque"
    cmbStatementType.AddItem "Credit"
    
    Call prcLoadQueryType
    
    DTPickerFromDate.Value = Date
    DTPickerToDate.Value = Date
    
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub prcLoadQueryType()
    cmbQueryType.Clear
    cmbQueryType.AddItem "Transaction Date"
    cmbQueryType.AddItem "Bill Date"
End Sub

Private Sub cmbStatementType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
