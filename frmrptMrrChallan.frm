VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmrptMrrChallan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7635
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmrptMrrChallan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGroup 
      BackColor       =   &H00E26101&
      Caption         =   "GROUP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1768
      TabIndex        =   34
      Top             =   6292
      Visible         =   0   'False
      Width           =   949
   End
   Begin VB.ComboBox cboitem 
      Height          =   315
      Left            =   360
      TabIndex        =   32
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATE"
      ForeColor       =   &H00E26101&
      Height          =   1170
      Left            =   3840
      TabIndex        =   27
      Top             =   4680
      Width           =   3735
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1800
         TabIndex        =   31
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox txtFromDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TO DATE :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FROM DATE :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.ComboBox cmbthread 
      Height          =   315
      Left            =   4200
      TabIndex        =   23
      Text            =   "Thread"
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CheckBox ChkOther 
      Caption         =   "Other"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox vmrrid 
      Height          =   375
      Left            =   255
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbMrrID 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   405
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5055
      Width           =   3195
   End
   Begin VB.ComboBox cmbChallanNo1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox cmbSupplier1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6840
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  : :  Select Criteria  : :  "
      ForeColor       =   &H00E26101&
      Height          =   1635
      Left            =   375
      TabIndex        =   10
      Top             =   960
      Width           =   6825
      Begin VB.OptionButton OptSupItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier && Item"
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   300
         TabIndex        =   26
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Opthread 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO (Thread)"
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   300
         TabIndex        =   21
         Top             =   1200
         Width           =   2415
      End
      Begin VB.OptionButton optChallanDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Challan Date Wise"
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   3600
         TabIndex        =   6
         Top             =   1260
         Width           =   2835
      End
      Begin VB.OptionButton optMrrNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR No Wise"
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   2625
      End
      Begin VB.OptionButton optSupplierMrrDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier && MRR Date Wise"
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   300
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton optChallanNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Challan No Wise"
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   3600
         TabIndex        =   9
         Top             =   900
         Width           =   2835
      End
      Begin VB.OptionButton optMRRDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR Date Wise"
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   540
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Report"
      Height          =   525
      Left            =   4800
      TabIndex        =   4
      Top             =   6240
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   525
      Left            =   6180
      TabIndex        =   5
      Top             =   6225
      Width           =   1185
   End
   Begin MSDataListLib.DataCombo cmbChallanNo 
      Height          =   315
      Left            =   4200
      TabIndex        =   24
      Top             =   3480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   315
      Left            =   360
      TabIndex        =   25
      Top             =   3480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ITEM NAME :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Left            =   360
      TabIndex        =   33
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PO(THREAD) NO  :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Left            =   4080
      TabIndex        =   22
      Top             =   3840
      Width           =   2115
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   6045
      Width           =   7635
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -120
      TabIndex        =   17
      Top             =   6000
      Width           =   7755
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   7875
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: MRR Information "
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
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   225
      Width           =   3570
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MRR NO  :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Left            =   405
      TabIndex        =   13
      Top             =   4785
      Width           =   2115
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHALLAN NO :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   1
      Left            =   4125
      TabIndex        =   12
      Top             =   3165
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUPPLIER NAME :"
      ForeColor       =   &H00E26101&
      Height          =   285
      Left            =   405
      TabIndex        =   11
      Top             =   3165
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7635
   End
End
Attribute VB_Name = "frmrptMrrChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BLOCK_SIZE = 10000

Private Sub ChkOther_Click()
If Len(cmbSupplier) = 0 Then Exit Sub
If ChkOther.Value = 1 Then
    Call prcAddSuppMrrIDOC(cmbMRRID, cmbSupplier.Text)
ElseIf ChkOther.Value = 0 Then
    Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
End If
End Sub

Private Sub cmbMrrID_Click()
   If Len(cmbMRRID) = 0 Then Exit Sub
    cmbMRRID.Tag = cmbMRRID.ItemData(cmbMRRID.ListIndex)
    Call prcAddSuppMrrno(vmrrid, cmbMRRID.Text, cmbSupplier.Text)
    
End Sub

Private Sub cmbMrrID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbSupplier_Change()
If Len(cmbSupplier) = 0 Then Exit Sub
    Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
    
    Call addSuppItem(cboItem, cmbSupplier.Text)
    
End Sub

'Private Sub cmbSupplier_Click()
'    If Len(cmbSupplier) = 0 Then Exit Sub
'    Call prcAddSuppMrrID(cmbMrrID, cmbSupplier.Text)
'End Sub
Private Sub cmbSupplierName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub cmbChallanNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdPreview_Click()
    Call prcShowReport("PREVIEW")
End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Form_Load()
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call addSupplierNameD(cmbSupplier)
    Call prcAddChallanNo(cmbChallanNo)
    addThreadPONo cmbthread
   
    ''If bytCompID = 1 Or bytCompID = 5 Or bytCompID = 7 Or bytCompID = 13 Then chkGroup.Visible = True
    ''If bytCompID = 1 Then chkGroup.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmrptMRRInformation = Nothing
End Sub

Private Sub optChallanDate_Click()
    cmbSupplier.Enabled = False
   cmbthread.Enabled = False
    cmbChallanNo.Enabled = False
    cmbMRRID.Enabled = False
    cboItem.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
    chkGroup.Visible = False
End Sub

Private Sub optChallanDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optPFNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Opthread_Click()

'    cmbSupplier.Enabled = False
 cmbthread.Enabled = True

'    cmbChallanNo.Enabled = False
chkGroup.Visible = False
End Sub

Private Sub OptSupItem_Click()

 cmbSupplier.Enabled = True
   
    cmbChallanNo.Enabled = False
    cboItem.Enabled = True
    cmbthread.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
  If bytCompID = 1 Then chkGroup.Visible = True
End Sub

Private Sub optSupplierMrrDate_Click()
    cmbSupplier.Enabled = True
    cmbthread.Enabled = False
    cmbChallanNo.Enabled = False
    cboItem.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
    chkGroup.Visible = False
End Sub

Private Sub optSupplierMrrDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optChallanNo_Click()
    cmbSupplier.Enabled = False
    cmbthread.Enabled = False
    cmbChallanNo.Enabled = True
    cmbMRRID.Enabled = False
    cboItem.Enabled = False
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
    chkGroup.Visible = False
End Sub

Private Sub optChallanNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optLCNoMrrDate_Click()
    cmbSupplier.Enabled = True
   
    cmbChallanNo.Enabled = False
    cmbMRRID.Enabled = False
    
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
End Sub

Private Sub optLCNoMrrDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optLCNoWise_Click()
    cmbSupplier.Enabled = True
    
    cmbChallanNo.Enabled = False
    cmbMRRID.Enabled = False
    
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
End Sub

Private Sub optLCNoWise_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optMRRDate_Click()
    cmbSupplier.Enabled = False
  cmbthread.Enabled = False
    cmbChallanNo.Enabled = False
    cmbMRRID.Enabled = False
    cboItem.Enabled = False
    txtFromDate.Enabled = True
    txtToDate.Enabled = True
    chkGroup.Visible = False
End Sub

Private Sub optMRRDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub optMRRNO_Click()
    cmbSupplier.Enabled = True
    
    cmbChallanNo.Enabled = False
    cmbMRRID.Enabled = True
    cboItem.Enabled = False
    cmbthread.Enabled = False
    txtFromDate.Enabled = False
    txtToDate.Enabled = False
    chkGroup.Visible = False
End Sub

Private Sub optMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtFromDate_LostFocus()
    Call FormatDate(txtFromDate)
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtToDate_LostFocus()
    Call FormatDate(txtToDate)
End Sub

Private Sub prcAddChallanNo(cboMain As DataCombo)

    cboMain.Text = ""
    
    Set R = New ADODB.Recordset
    
    
    
    
    
    strRecord = " SELECT dbo.MRR_Challan_Sub.ChallanNo FROM dbo.MRR_Challan_Main INNER JOIN dbo.MRR_Challan_Sub ON dbo.MRR_Challan_Main.MRRID = dbo.MRR_Challan_Sub.MRRID" _
              & " Where (dbo.MRR_Challan_Main.compID = " & bytCompID & ") AND  (dbo.MRR_Challan_Sub.DELETEROW=0) " _
              & "  GROUP BY dbo.MRR_Challan_Sub.ChallanNo ORDER BY dbo.MRR_Challan_Sub.ChallanNo;"
    
    R.Open strRecord, MainConn, adOpenStatic
    
     Set cboMain.RowSource = R
            cboMain.ListField = "ChallanNo"
             cboMain.BoundColumn = "ChallanNo"
    
   
'            If R.RecordCount > 0 Then
'
'
'                Do Until R.EOF
'
'                    cboMain.AddItem UCase(R![ChallanNo])
'
'                R.MoveNext
'
'
'                Loop
'
'
'            End If
'
'
'    R.Close
    Set R = Nothing
    
End Sub

Private Sub prcShowReport(Param As String)


Dim str1 As String
Dim file_num As String
Dim file_length As String
Dim bytes() As Byte
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long
Dim aa As String


    If fncBlank = True Then Exit Sub

        If optSupplierMrrDate.Value = True Then
        
          If ChkOther.Value = 1 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_Supplier_MrrDateOC.rpt")
            REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtQ) & " To " & Format(txtToDate.Text, cnstDtFrmtQ) & "'"
          Else
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_Supplier_MrrDate.rpt")
            REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtQ) & " To " & Format(txtToDate.Text, cnstDtFrmtQ) & "'"
          End If
            
        ElseIf optMRRNO.Value = True Then
            
          If ChkOther.Value = 0 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_MrrNo11.rpt")
          ElseIf ChkOther.Value = 1 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_MrrNo11OC.rpt")
          End If
            
        ElseIf optMRRDate.Value = True Then
        
          If ChkOther.Value = 1 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_MrrDateOC.rpt")
            REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtQ) & " To " & Format(txtToDate.Text, cnstDtFrmtQ) & "'"
          Else
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_MrrDate.rpt")
            REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtQ) & " To " & Format(txtToDate.Text, cnstDtFrmtQ) & "'"
          End If
            
        ElseIf optChallanNo.Value = True Then
        
          If ChkOther.Value = 1 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_ChallanNoOC.rpt")
          Else
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_ChallanNo.rpt")
          End If
            
        ElseIf optChallanDate.Value = True Then
        
          If ChkOther.Value = 1 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_ChallanDateOC.rpt")
            REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtI) & " To " & Format(txtToDate.Text, cnstDtFrmtI) & "'"
          Else
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_ChallanDate.rpt")
            REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtI) & " To " & Format(txtToDate.Text, cnstDtFrmtI) & "'"
          End If
            
        ElseIf OptSupItem.Value = True Then
        
          If chkGroup.Value = 1 Then
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_Supplier_ItemName_Gr.rpt")
          Else
            Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_Supplier_ItemName.rpt")
          End If
          REPORT.FormulaFields(3).Text = "'" & Format(txtFromDate.Text, cnstDtFrmtQ) & " To " & Format(txtToDate.Text, cnstDtFrmtQ) & "'"

        ElseIf Opthread.Value = True Then
           
      '-------------
          If Len(cmbthread.Text) = 0 Then
            MsgBox "SELECT  PO NO.", vbInformation, cnstMsgInfo
            cboeditPO.SetFocus
            Exit Sub
          End If
             
          Set R = New ADODB.Recordset
          Set MainComm = New ADODB.Command
          MainComm.ActiveConnection = MainConn
          MainComm.CommandTimeout = 500
             
          MainComm.CommandText = "delete  from spPhoto;"
          MainComm.Execute
            
          MainComm.CommandText = "delete  from Temp_ThreadInfo;"
          MainComm.Execute
             
          '------------------FIND OUT PATH
          aa = ""
            
          strRecord = "SELECT  Path  FROM [ThreadInfo] WHERE PONO='" & cmbthread.Text & "'   AND DELETEROW=0  ;"
              
          R.Open strRecord, MainConn, adOpenStatic
          If R.RecordCount > 0 Then
            ' --IIf(IsNull(R![Invvalue]), "0.00", R![Invvalue])
            aa = IIf(IsNull(R![Path]), "", R![Path])
          End If
          R.Close
          Set R = Nothing
            
          '----------------------------
          file_num = FreeFile
              
          If aa <> "" Then
              
             Open aa For Binary Access Read As #file_num
                
                file_length = LOF(file_num)
                
                Set StdPicture = LoadPicture(aa)
              
               If file_length > 0 Then
                    
                    num_blocks = file_length / BLOCK_SIZE
                    left_over = file_length Mod BLOCK_SIZE
                      
            'MainComm.CommandText = "delete  from spPhoto;"
            'MainComm.Execute
            '
            'MainComm.CommandText = "delete  from Temp_ThreadInfo;"
            'MainComm.Execute
            
                
                    Set RS = New Recordset
                    
                    str1 = "select* from SpPhoto"
                    RS.Open str1, MainConn, 3, 3
                    
                    
                    
                    RS.AddNew
                    
                    RS.Fields("SpLength").Value = file_length
            
                    ReDim bytes(BLOCK_SIZE)
                    
                    For block_num = 1 To num_blocks
                        Get #file_num, , bytes()
                        RS.Fields("SpPhoto").AppendChunk bytes()
                    Next block_num
            
                    If left_over > 0 Then
                        ReDim bytes(left_over)
                        Get #file_num, , bytes()
                        RS.Fields("SpPhoto").AppendChunk bytes()
                    End If
            
                      RS.Update
                    Close #file_num
                    RS.Close
                    End If
              
            '-------------------
                      
              Set REPORT = Appl.OpenReport(findReportPath & "\rptThread.rpt")
             
                    MainComm.CommandText = " INSERT INTO Temp_ThreadInfo Select * From ThreadInfo Where CompID=" & bytCompID & "  AND PONO='" & cmbthread.Text & "'  AND DELETEROW=0 "
                    MainComm.Execute
                    
                    MainComm.CommandText = "Update Temp_ThreadInfo set Temp_ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM Temp_ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & cmbthread.Text & "'  AND Temp_ThreadInfo.DELETEROW=0 "
                    MainComm.Execute
              
              Else
              
              
               Set REPORT = Appl.OpenReport(findReportPath & "\rptThread.rpt")
             
                    MainComm.CommandText = " INSERT INTO Temp_ThreadInfo Select * From ThreadInfo Where CompID=" & bytCompID & "  AND PONO='" & cmbthread.Text & "'  AND DELETEROW=0 "
                    MainComm.Execute
                    
                    MainComm.CommandText = "Update Temp_ThreadInfo set Temp_ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM Temp_ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & cmbthread.Text & "'  AND Temp_ThreadInfo.DELETEROW=0 "
                    MainComm.Execute
  
       '-------------
           End If
           
            
        End If

    Mysql = fncSQLStatement
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    Set R = New ADODB.Recordset
    
    If aa <> "" Then
    
'            If optMRRNO.Value = False Then
'              strRecord = Mysql
'            End If
'
'           R.Open strRecord, MainConn, adOpenStatic
'
'               Report.DiscardSavedData
'            If optMRRNO.Value = False Then
'                Report.Database.SetDataSource R
'            End If
            
   
    
            frmMainReport.CRVIEWER.Refresh
            frmMainReport.CRVIEWER.ReportSource = REPORT
            frmMainReport.CRVIEWER.ViewReport
            frmMainReport.CRVIEWER.Zoom (82)
            frmMainReport.Show 1
          '  R.Close
    
    Else
    
       
              
            If Opthread.Value = False Then
            
                      If optMRRNO.Value = False Then
                         strRecord = Mysql
                      End If
                     
                    R.Open strRecord, MainConn, adOpenStatic
                     
                        REPORT.DiscardSavedData
                     If optMRRNO.Value = False Then
                         REPORT.Database.SetDataSource R
                     End If
            End If
            
            frmMainReport.CRVIEWER.Refresh
            frmMainReport.CRVIEWER.ReportSource = REPORT
            frmMainReport.CRVIEWER.ViewReport
            frmMainReport.CRVIEWER.Zoom (82)
            frmMainReport.Show 1
    
    End If
    
     Call addSupplierNameD(cmbSupplier)
     Call prcAddChallanNo(cmbChallanNo)

End Sub

Private Function fncBlank() As Boolean
fncBlank = True
    If optSupplierMrrDate.Value = True Then
       
        If Len(cmbSupplier.Text) = 0 Then
            MsgBox "You should provide Supplier name.", vbCritical, AppTitle
            cmbSupplier.SetFocus
            Exit Function
        End If
        If Len(txtFromDate.Text) = 0 Then
            MsgBox "You should provide from date.", vbCritical, AppTitle
            txtFromDate.SetFocus
            Exit Function
        End If
        If Len(txtToDate.Text) = 0 Then
            MsgBox "You should provide to date.", vbCritical, AppTitle
            txtToDate.SetFocus
            Exit Function
        End If
        
    ElseIf optMRRNO.Value = True Then
       
        If Len(cmbSupplier.Text) = 0 Then
            MsgBox "You should provide Supplier name.", vbCritical, AppTitle
            cmbSupplier.SetFocus
            Exit Function
        End If
        If Len(cmbMRRID.Text) = 0 Then
            MsgBox "You should provide MRR No.", vbCritical, AppTitle
            cmbMRRID.SetFocus
            Exit Function
        End If
    ElseIf optMRRDate.Value = True Then
        If Len(txtFromDate.Text) = 0 Then
            MsgBox "You should provide from date.", vbCritical, AppTitle
            txtFromDate.SetFocus
            Exit Function
        End If
        If Len(txtToDate.Text) = 0 Then
            MsgBox "You should provide to date.", vbCritical, AppTitle
            txtToDate.SetFocus
            Exit Function
        End If
    ElseIf optChallanNo.Value = True Then
        If Len(cmbChallanNo.Text) = 0 Then
            MsgBox "You should provide challan No.", vbCritical, AppTitle
            cmbChallanNo.SetFocus
            Exit Function
        End If
    ElseIf optChallanDate.Value = True Then
        '' MRR Date Wise
        If Len(txtFromDate.Text) = 0 Then
            MsgBox "You should provide from date.", vbCritical, AppTitle
            txtFromDate.SetFocus
            Exit Function
        End If
        If Len(txtToDate.Text) = 0 Then
            MsgBox "You should provide to date.", vbCritical, AppTitle
            txtToDate.SetFocus
            Exit Function
        End If
    ElseIf OptSupItem.Value = True Then
        If Len(cmbSupplier.Text) = 0 Then
            MsgBox "You should provide challan No.", vbCritical, AppTitle
            cmbSupplier.SetFocus
            Exit Function
        End If
        If Len(cboItem.Text) = 0 Then
            MsgBox "You should provide Item Name", vbCritical, AppTitle
            cboItem.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function

Private Function fncSQLStatement() As String

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set R = New ADODB.Recordset
Set RS = New ADODB.Recordset

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

MainComm.CommandText = "DELETE FROM TEMPSUM "
MainComm.Execute

If optSupplierMrrDate.Value = True Then
  
  If ChkOther.Value = 1 Then
    ''If chkGroup.Value = 1 Then
    ''  fncSQLStatement = "SELECT * FROM VIEWMRRREPORTOC_Group WHERE suppName = '" & cmbSupplier.Text & "' AND MrrDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "';"
    ''Else
      fncSQLStatement = "Select * from VIEWMRRREPORTOC Where (Deleterow=0 and (Compid = " & bytCompID & ") and (SuppId = '" & findSupplierID(cmbSupplier.Text) & "') and (MrrDate Between '" & txtFromDate & "' and '" & txtToDate & "')); "
    ''End If
  Else
    ''If chkGroup.Value = 1 Then
    ''  fncSQLStatement = "SELECT * FROM VIEWMRRREPORT_Group WHERE suppName = '" & cmbSupplier.Text & "' AND MrrDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "';"
    ''Else
      fncSQLStatement = "Select * from VIEWMRRREPORT Where (Deleterow=0 and (Compid = " & bytCompID & ") and (SuppId = '" & findSupplierID(cmbSupplier.Text) & "') and (MrrDate Between '" & txtFromDate & "' and '" & txtToDate & "')); "
    ''End If
  End If

ElseIf optMRRNO.Value = True Then
  
  If ChkOther.Value = 0 Then
    'fncSQLStatement = "Select * from VIEWMRRREPORT Where (Deleterow=0 and (Compid = " & bytCompID & ") and (MrrID ='" & vmrrid.Text & "')) Order By ItemType, ItemName, MATCODE; "
    strRecord = "SELECT * FROM MRR_Challan_Sub WHERE (dbo.MRR_Challan_Sub.DeleteRow = 0) AND (dbo.MRR_Challan_Sub.MRRID = '" & vmrrid.Text & "')"
    R.Open strRecord, MainConn, adOpenStatic
    
    Do Until R.EOF
      strRecord = "SELECT SUM(dbo.MRR_Challan_Sub.RcvdQty) AS Sum,dbo.MRR_Challan_Sub.PIId, dbo.MRR_Challan_Sub.ItDesc, dbo.MRR_Challan_Sub.PRNO,dbo.MRR_Challan_Sub.ITEMID FROM dbo.MRR_Challan_Sub LEFT OUTER JOIN" _
                & " dbo.MRR_Challan_Main INNER JOIN  dbo.New_MRR ON dbo.MRR_Challan_Main.MRRID = dbo.New_MRR.MRRID" _
                & " AND dbo.MRR_Challan_Main.compID = dbo.New_MRR.ComID ON dbo.MRR_Challan_Sub.MRRID = dbo.MRR_Challan_Main.MRRID" _
                & " WHERE (dbo.MRR_Challan_Sub.DeleteRow = 0) AND (dbo.MRR_Challan_Sub.MRRID <= '" & vmrrid.Text & "') AND (dbo.MRR_Challan_Sub.PIId = '" & R![PIId] & "') AND (dbo.MRR_Challan_Sub.ItDesc='" & R![ItDesc] & "') AND (dbo.MRR_Challan_Sub.PRNO='" & R![PRNo] & "')AND (dbo.MRR_Challan_Sub.ITEMID = '" & R![ITEMID] & "')" _
                & " GROUP BY dbo.MRR_Challan_Sub.PIId,dbo.MRR_Challan_Sub.ItDesc, dbo.MRR_Challan_Sub.PRNO,dbo.MRR_Challan_Sub.ITEMID"
      RS.Open strRecord, MainConn, adOpenStatic
      MainComm.CommandText = "INSERT INTO TEMPSUM  (MRRID,PIID,ITDESC,SUM,ITEMID,PRNO) VALUES ('" & vmrrid.Text & "','" & RS![PIId] & "','" & RS![ItDesc] & "','" & RS![Sum] & "','" & RS![ITEMID] & "','" & RS![PRNo] & "')"
      MainComm.Execute
      RS.Close
      R.MoveNext
    Loop
  
  ElseIf ChkOther.Value = 1 Then
    'fncSQLStatement = "Select * from VIEWMRRREPORTOC Where (Deleterow=0 and (Compid = " & bytCompID & ") and (MrrID ='" & vmrrid.Text & "')) Order By ItemType, ItemName, MATCODE; "
    strRecord = "SELECT * FROM MRR_Challan_SubOC WHERE (dbo.MRR_Challan_SubOC.DeleteRow = 0) AND (dbo.MRR_Challan_SubOC.MRRID = '" & vmrrid.Text & "')"
    R.Open strRecord, MainConn, adOpenStatic
    
    Do Until R.EOF
      strRecord = "SELECT SUM(dbo.MRR_Challan_SubOC.RcvdQty) AS Sum,dbo.MRR_Challan_SubOC.PIId, dbo.MRR_Challan_SubOC.ItDesc,dbo.MRR_Challan_SubOC.PRNO,dbo.MRR_Challan_SubOC.ITEMID FROM dbo.MRR_Challan_SubOC LEFT OUTER JOIN" _
                & " dbo.MRR_Challan_MainOC INNER JOIN  dbo.New_MRR ON dbo.MRR_Challan_MainOC.MRRID = dbo.New_MRR.MRRID" _
                & " AND dbo.MRR_Challan_MainOC.compID = dbo.New_MRR.ComID ON dbo.MRR_Challan_SubOC.MRRID = dbo.MRR_Challan_MainOC.MRRID" _
                & " WHERE (dbo.MRR_Challan_SubOC.DeleteRow = 0) AND (dbo.MRR_Challan_SubOC.MRRID <= '" & vmrrid.Text & "') AND (dbo.MRR_Challan_SubOC.PIId = '" & R![PIId] & "') AND (dbo.MRR_Challan_SubOC.ItDesc='" & R![ItDesc] & "') AND (dbo.MRR_Challan_SubOC.PRNO='" & R![PRNo] & "') AND (dbo.MRR_Challan_SubOC.ITEMID = '" & R![ITEMID] & "')" _
                & " GROUP BY dbo.MRR_Challan_SubOC.PIId,dbo.MRR_Challan_SubOC.ItDesc,dbo.MRR_Challan_SubOC.PRNO,dbo.MRR_Challan_SubOC.ITEMID"
      RS.Open strRecord, MainConn, adOpenStatic
      MainComm.CommandText = "INSERT INTO TEMPSUM  (MRRID,PIID,ITDESC,SUM,ITEMID) VALUES ('" & vmrrid.Text & "','" & RS![PIId] & "','" & RS![ItDesc] & "','" & RS![Sum] & "','" & RS![ITEMID] & "')"
      MainComm.Execute
      RS.Close
      R.MoveNext
    Loop
  End If

ElseIf optMRRDate.Value = True Then
  
  If ChkOther.Value = 1 Then
    ''If chkGroup.Value = 1 Then
    ''  fncSQLStatement = "SELECT * FROM VIEWMRRREPORTOC_Group WHERE MrrDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "' ORDER BY MrrDate,MrrID;"
    ''Else
      fncSQLStatement = "Select * from VIEWMRRREPORTOC Where (Deleterow=0 and (Compid = " & bytCompID & ") and (MRRDate Between '" & txtFromDate & "' and '" & txtToDate & "')) order by MrrDate, MrrID; "
    ''End If
  
  Else
    ''If chkGroup.Value = 1 Then
    ''  fncSQLStatement = "SELECT * FROM VIEWMRRREPORT_Group WHERE MrrDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "' ORDER BY MrrDate,MrrID;"
    ''Else
      fncSQLStatement = "Select * from VIEWMRRREPORT Where (Deleterow=0 and (Compid = " & bytCompID & ") and (MRRDate Between '" & txtFromDate & "' and '" & txtToDate & "')) order by MrrDate, MrrID; "
    ''End If
  End If

ElseIf optChallanNo.Value = True Then
  
  If ChkOther.Value = 1 Then
    fncSQLStatement = "Select * from VIEWMRRREPORTOC Where (Deleterow=0 and (Compid = " & bytCompID & ") and (ChallanNO = '" & cmbChallanNo.Text & "')); "
  Else
    fncSQLStatement = "Select * from VIEWMRRREPORT Where (Deleterow=0 and (Compid = " & bytCompID & ") and (ChallanNO = '" & cmbChallanNo.Text & "')); "
  End If
  
ElseIf OptSupItem.Value = True Then
  
  If chkGroup.Value = 1 Then
    fncSQLStatement = "SELECT * FROM View_Supplier_ItemName_Group WHERE suppName = '" & cmbSupplier.Text & "' AND itemName = '" & cboItem.Text & "' AND MrrDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "';"
  Else
    fncSQLStatement = "select * FROM View_Supplier_ItemName  WHERE DELETEROW=0 AND  suppName='" & cmbSupplier.Text & "' AND itemName ='" & cboItem.Text & "' AND (MRRDate Between '" & txtFromDate & "' and '" & txtToDate & "') "
  End If

ElseIf optChallanDate.Value = True Then
  
  If ChkOther.Value = 1 Then
    ''If chkGroup.Value = 1 Then
    ''  fncSQLStatement = "SELECT * FROM VIEWMRRREPORTOC_Group WHERE ChallanDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "';"
    ''Else
      fncSQLStatement = "Select * from VIEWMRRREPORTOC Where (Deleterow=0 and (Compid = " & bytCompID & ") and (ChallanDate Between '" & txtFromDate & "' and '" & txtToDate & "')); "
    ''End If
  
  Else
    ''If chkGroup.Value = 1 Then
    ''  fncSQLStatement = "SELECT * FROM VIEWMRRREPORT_Group WHERE ChallanDate BETWEEN '" & txtFromDate & "' AND '" & txtToDate & "';"
    ''Else
      fncSQLStatement = "Select * from VIEWMRRREPORT Where (Deleterow=0 and (Compid = " & bytCompID & ") and (ChallanDate Between '" & txtFromDate & "' and '" & txtToDate & "')); "
    ''End If
  End If

End If

MainConn.CommitTrans

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
End If

Screen.MousePointer = vbDefault

End Function
