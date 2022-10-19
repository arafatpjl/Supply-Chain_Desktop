VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrptMRRWiseReport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MRR Wise Report"
   ClientHeight    =   4680
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
   Icon            =   "frmrptMRRWiseReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6786
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMRRNo 
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
      Top             =   2730
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
      Top             =   3360
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker DTPickerToDate 
      Height          =   375
      Left            =   2580
      TabIndex        =   8
      Top             =   1590
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
      Top             =   1020
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4110
      Width           =   1125
   End
   Begin VB.ComboBox cmbMRRNo 
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
      Top             =   2250
      Width           =   3375
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3330
      Top             =   4020
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
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   3960
      Width           =   6795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: MRR Wise "
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
      TabIndex        =   13
      Top             =   120
      Width           =   2625
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "MRR No :"
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
      Top             =   2760
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
      Top             =   3900
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
      Top             =   1110
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
      Top             =   1650
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select MRR No :"
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
      Top             =   2250
      Width           =   1785
   End
End
Attribute VB_Name = "frmrptMRRWiseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbMRRNo_Click()
    txtMRRNo.Text = cmbMRRNo.Text
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
'    MousePointer = 11
    
    If txtMRRNo.Text = "" Then
        MsgBox "You should provide MRR No.", vbCritical, App.Title
        txtMRRNo.SetFocus
        Exit Sub
    End If
    If chkAllCompany.Value = 0 Then
        Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptMRRWise.rpt")
        strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And MRRNo = '" & txtMRRNo.Text & "';"

    Else
        Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptMRRWise.rpt")
        strRecord = "Select * From viewStatement_Cash_Credit Where MRRNo = '" & txtMRRNo.Text & "';"
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
    
 '   MousePointer = 0
    Exit Sub
x:
  '  MousePointer = 0
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub DTPickerFromDate_Change()
    Set R = New Recordset
        sql1 = "Select MRRNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By MRRNo Order By MRRNo"
        R.Open sql1, MainConn, adOpenStatic
        While Not R.EOF
            cmbMRRNo.AddItem R.Fields("MRRNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub DTPickerFromDate_Click()
    Set R = New Recordset
        sql1 = "Select MRRNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By MRRNo Order By MRRNo"
        R.Open sql1, MainConn, adOpenStatic
        While Not R.EOF
            cmbMRRNo.AddItem R.Fields("MRRNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub DTPickerToDate_Change()
     Set R = New Recordset
        sql1 = "Select MRRNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By MRRNo Order By MRRNo"
        R.Open sql1, MainConn, adOpenStatic
        While Not R.EOF
            cmbMRRNo.AddItem R.Fields("MRRNO")
            R.MoveNext
        Wend
    Set R = Nothing
End Sub

Private Sub DTPickerToDate_Click()
     Set R = New Recordset
        sql1 = "Select MRRNO From Statement_of_Cash_and_Credit Where Tran_Date Between '" & DTPickerFromDate.Value & "' And '" & DTPickerToDate.Value & "' Group By MRRNo Order By MRRNo"
        R.Open sql1, MainConn, adOpenStatic
        While Not R.EOF
            cmbMRRNo.AddItem R.Fields("MRRNO")
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

Private Sub cmbMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
