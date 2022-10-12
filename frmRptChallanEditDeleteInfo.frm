VERSION 5.00
Begin VB.Form frmRptChallanEditDeleteInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4290
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00C0C000&
      Caption         =   "Pre&view"
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
      Left            =   4575
      TabIndex        =   5
      ToolTipText     =   "Preview Report"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
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
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "Close Form"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.TextBox txtChaFrDate 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   2235
      Width           =   1830
   End
   Begin VB.TextBox txtChaToDate 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   2835
      Width           =   1830
   End
   Begin VB.Frame frmMrr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   945
      Left            =   90
      TabIndex        =   0
      Tag             =   "2"
      Top             =   915
      Width           =   5685
      Begin VB.OptionButton optDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   375
         Left            =   2745
         TabIndex        =   2
         Top             =   315
         Width           =   1200
      End
      Begin VB.OptionButton optEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   375
         Left            =   1485
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: CHALLAN Edit/Delete "
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
      Left            =   270
      TabIndex        =   12
      Top             =   210
      Width           =   4155
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
      Index           =   2
      Left            =   15
      TabIndex        =   11
      Top             =   840
      Width           =   5970
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   3420
      Width           =   5835
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   15
      TabIndex        =   9
      Top             =   3370
      Width           =   5955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "From Date :"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1995
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "To Date :"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2595
      Width           =   750
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmRptChallanEditDeleteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Set frmRptMrrEditDeleteInfo = Nothing
    Unload Me
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdReport_Click()
  Call generateReport
  
  Set REPORT = Appl.OpenReport(findReportPath & "\rptChallanEditDeleteInfo.rpt")

  If boEmptyText = True Then Exit Sub
  
  REPORT.DiscardSavedData
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = REPORT
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1

End Sub

Private Sub cmdReport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Form_Load()
Call prcMakeCenter(Me)
'    Call prcConReport(CrystalReport1)
End Sub

Private Sub optDelete_Click()
    frmMrr.Tag = 1
End Sub

Private Sub optEdit_Click()
    frmMrr.Tag = 2
End Sub

Private Sub txtChaFrDate_GotFocus()
    Call txtGotFocus(txtChaFrDate)
End Sub

Private Sub txtChaFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub txtChaFrDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub
Private Sub txtChaFrDate_LostFocus()
    txtChaFrDate.Text = FormatDate(txtChaFrDate)
End Sub

Private Sub txtChaToDate_GotFocus()
    Call txtGotFocus(txtChaToDate)
End Sub

Private Sub txtChaToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub txtChaToDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub
Private Sub txtChaToDate_LostFocus()
    txtChaToDate.Text = FormatDate(txtChaToDate)
End Sub

Sub generateReport()

boEmptyText = False
If Len(txtChaFrDate.Text) = 0 Then
    MsgBox "Type Starting Date..", vbInformation, cnstMsgInfo
    txtChaFrDate.SetFocus
    boEmptyText = True
    Exit Sub
  End If

If Len(txtChaToDate.Text) = 0 Then
    MsgBox "Type Ending Date", vbInformation, cnstMsgInfo
    txtChaToDate.SetFocus
    boEmptyText = True
    Exit Sub
  End If


Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

MainConn.BeginTrans

MainComm.CommandText = "DELETE  FROM Temp_Mrr_Challan_Info;"
MainComm.Execute

If frmMrr.Tag = 2 Then

    MainComm.CommandText = "Insert Into Temp_Mrr_Challan_Info(SuppName,MrrId,MRR_PONo,MrrDate,MRRCurr,MRRValue,ItemName,MRRQty,RcvdQty,ItTotal,Remarks,QtyUnit,UnitPrice,MatCode,PRNO,MainPF,ChallanNo,ChallanDate,ContainerNo,CompID,compName,compAdd,UserName,ComputerName,EntryDate,EntryTime)" _
    & " SELECT new_Supplier.suppName,MRR_Challan_Main.MRRID,New_MRR.MRRNo,MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr,MRR_Challan_Main.MRRValue, new_Item_Name.itemName, SUM(MRR_Challan_Sub.MRRQty) AS MRRQty," _
    & " SUM(MRR_Challan_Sub.RcvdQty) AS RcvdQty, SUM(MRR_Challan_Sub.ItTotal) AS ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice,MRR_Challan_Sub.MatCode,MRR_Challan_Sub.PRNO,MRR_Challan_Sub.MainPF,MRR_Challan_Sub.ChallanNo,MRR_Challan_Sub.ChallanDate,MRR_Challan_Sub.Containerno," _
    & " Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,Sys_Configure_Status.ComName,Sys_User_Name.UserName,MRR_Challan_Sub.EntryDate,MRR_Challan_Sub.EntryTime FROM dbo.Company_Information INNER JOIN new_Item_Name INNER JOIN new_Supplier INNER JOIN MRR_Challan_Main INNER JOIN MRR_Challan_Sub ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID ON" _
    & " new_Supplier.supplierID = MRR_Challan_Main.suppID ON new_Item_Name.itemID = MRR_Challan_Sub.itemID ON " _
    & " Company_Information.CompID = MRR_Challan_Main.compID INNER JOIN Sys_Configure_Status ON MRR_Challan_Sub.PCID = Sys_Configure_Status.IDNo INNER JOIN InventoryPJL.dbo.Sys_User_Name ON MRR_Challan_Sub.UserId = InventoryPJL.dbo.Sys_User_Name.UserID Inner Join New_MRR On New_MRR.MRRID=MRR_Challan_Main.MRRID " _
    & " WHERE (MRR_Challan_Main.compID = " & bytCompID & ") AND (MRR_Challan_Sub.ChallanDate Between '" & Format(txtChaFrDate.Text, cnstDtFrmtQ) & "' And '" & Format(txtChaToDate.Text, cnstDtFrmtQ) & "') AND (MRR_Challan_Sub.DeleteRow = 2)" _
    & " GROUP BY new_Supplier.suppName, MRR_Challan_Main.MRRID, New_MRR.MRRNo,MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr,MRR_Challan_Main.MRRValue, new_Item_Name.itemName, MRR_Challan_Sub.Remarks, MRR_Challan_Sub.QtyUnit," _
    & " MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MainPF,MRR_Challan_Sub.ChallanNo,MRR_Challan_Sub.ChallanDate,MRR_Challan_Sub.Containerno, Company_Information.CompID," _
    & " Company_Information.CompName , Company_Information.CompAdd, Sys_Configure_Status.ComName, InventoryPJL.dbo.Sys_User_Name.UserName,MRR_Challan_Sub.EntryDate,MRR_Challan_Sub.EntryTime;"
ElseIf frmMrr.Tag = 1 Then
    MainComm.CommandText = "Insert Into Temp_Mrr_Challan_Info(SuppName,MrrId,MRR_PONo,MrrDate,MRRCurr,MRRValue,ItemName,MRRQty,RcvdQty,ItTotal,Remarks,QtyUnit,UnitPrice,MatCode,PRNO,MainPF,ChallanNo,ChallanDate,ContainerNo,CompID,compName,compAdd,UserName,ComputerName,EntryDate,EntryTime)" _
    & " SELECT new_Supplier.suppName,MRR_Challan_Main.MRRID,New_MRR.MRRNo,MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr,MRR_Challan_Main.MRRValue, new_Item_Name.itemName, SUM(MRR_Challan_Sub.MRRQty) AS MRRQty," _
    & " SUM(MRR_Challan_Sub.RcvdQty) AS RcvdQty, SUM(MRR_Challan_Sub.ItTotal) AS ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice,MRR_Challan_Sub.MatCode,MRR_Challan_Sub.PRNO,MRR_Challan_Sub.MainPF,MRR_Challan_Sub.ChallanNo,MRR_Challan_Sub.ChallanDate,MRR_Challan_Sub.Containerno," _
    & " Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,Sys_Configure_Status.ComName,Sys_User_Name.UserName,MRR_Challan_Sub.EntryDate,MRR_Challan_Sub.EntryTime FROM dbo.Company_Information INNER JOIN new_Item_Name INNER JOIN new_Supplier INNER JOIN MRR_Challan_Main INNER JOIN MRR_Challan_Sub ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID ON" _
    & " new_Supplier.supplierID = MRR_Challan_Main.suppID ON new_Item_Name.itemID = MRR_Challan_Sub.itemID ON " _
    & " Company_Information.CompID = MRR_Challan_Main.compID INNER JOIN Sys_Configure_Status ON MRR_Challan_Sub.PCID = Sys_Configure_Status.IDNo INNER JOIN InventoryPJL.dbo.Sys_User_Name ON MRR_Challan_Sub.UserId = InventoryPJL.dbo.Sys_User_Name.UserID Inner Join New_MRR On New_MRR.MRRID=MRR_Challan_Main.MRRID  " _
    & " WHERE (MRR_Challan_Main.compID = " & bytCompID & ") AND (MRR_Challan_Sub.ChallanDate Between '" & Format(txtChaFrDate.Text, cnstDtFrmtQ) & "' And '" & Format(txtChaToDate.Text, cnstDtFrmtQ) & "') AND (MRR_Challan_Sub.DeleteRow = 1)" _
    & " GROUP BY new_Supplier.suppName, MRR_Challan_Main.MRRID,New_MRR.MRRNo, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr,MRR_Challan_Main.MRRValue, new_Item_Name.itemName, MRR_Challan_Sub.Remarks, MRR_Challan_Sub.QtyUnit," _
    & " MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MainPF,MRR_Challan_Sub.ChallanNo,MRR_Challan_Sub.ChallanDate,MRR_Challan_Sub.Containerno, Company_Information.CompID," _
    & " Company_Information.CompName , Company_Information.CompAdd, Sys_Configure_Status.ComName, InventoryPJL.dbo.Sys_User_Name.UserName,MRR_Challan_Sub.EntryDate,MRR_Challan_Sub.EntryTime;"
End If

MainComm.Execute
If frmMrr.Tag = 1 Then
    MainComm.CommandText = "Update Temp_Mrr_Challan_Info Set HeaderInfo='CHALLAN DELETE INFORMATION'"
Else
    MainComm.CommandText = "Update Temp_Mrr_Challan_Info Set HeaderInfo='CHALLAN EDIT INFORMATION'"
End If
MainComm.Execute
strRecord = "SELECT MrrId FROM Temp_Mrr_Challan_Info GROUP BY MrrId;"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount = 0 Then
  MsgBox "Record not Found", vbInformation, cnstMsgInfo
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  boEmptyText = True
  Exit Sub
End If
R.Close
MainConn.CommitTrans
ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  boEmptyText = True
  Exit Sub
End If
Screen.MousePointer = vbDefault
Set R = Nothing

End Sub






