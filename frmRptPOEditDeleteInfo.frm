VERSION 5.00
Begin VB.Form frmRptPOEditDeleteInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4260
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPOFrDate 
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
      Left            =   1935
      TabIndex        =   0
      Top             =   2205
      Width           =   1830
   End
   Begin VB.TextBox txtPOToDate 
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
      Left            =   1935
      TabIndex        =   1
      Top             =   2925
      Width           =   1830
   End
   Begin VB.Frame frmPO 
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
      TabIndex        =   9
      Tag             =   "2"
      Top             =   915
      Width           =   5610
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ALL P.O."
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
         Height          =   315
         Left            =   4080
         TabIndex        =   14
         Tag             =   "3"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EDITE"
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
         Left            =   960
         TabIndex        =   11
         Tag             =   "2"
         Top             =   360
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DELETE"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   1200
      End
   End
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
      Left            =   4425
      TabIndex        =   2
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
      Left            =   3090
      TabIndex        =   4
      ToolTipText     =   "Close Form"
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FROM DATE :"
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
      Left            =   1935
      TabIndex        =   13
      Top             =   1950
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "TO DATE :"
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
      Left            =   1935
      TabIndex        =   12
      Top             =   2670
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT ::  Edit/Delete /ALL P.O."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   2
      Left            =   135
      TabIndex        =   7
      Top             =   210
      Width           =   4515
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
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   5850
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   -60
      TabIndex        =   5
      Top             =   3420
      Width           =   5835
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -120
      TabIndex        =   3
      Top             =   3360
      Width           =   5955
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmRptPOEditDeleteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Set frmRptPOEditDeleteInfo = Nothing
    Unload Me
End Sub

Private Sub cmdReport_Click()

Call generateReport
    
    
    If frmPO.Tag = 3 Then

     Set REPORT = Appl.OpenReport(findReportPath & "\rptALLPOInfo.rpt")
     '  Set Report = Appl.OpenReport(findReportPath & "\rptPOEditDeleteInfo.rpt")
    
    
    Else
      Set REPORT = Appl.OpenReport(findReportPath & "\rptPOEditDeleteInfo.rpt")
    
    
    
    End If
  
  
  If boEmptyText = True Then Exit Sub
  
    REPORT.DiscardSavedData
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = REPORT
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1
  
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
'    Call prcConReport(CrystalReport1)
End Sub

Private Sub optAll_Click()
frmPO.Tag = 3
End Sub

Private Sub optDelete_Click()
frmPO.Tag = 1
End Sub

Private Sub optEdit_Click()
    frmPO.Tag = 2
End Sub

Private Sub txtPOFrDate_GotFocus()
    Call txtGotFocus(txtPOFrDate)
End Sub

Private Sub txtPOFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPOFrDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtPOFrDate_LostFocus()
txtPOFrDate.Text = FormatDate(txtPOFrDate)
End Sub

Private Sub txtPOToDate_GotFocus()
    Call txtGotFocus(txtPOToDate)
End Sub

Private Sub txtPOToDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPOToDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtPOToDate_LostFocus()
    txtPOToDate.Text = FormatDate(txtPOToDate)
End Sub

Sub generateReport()
boEmptyText = False

If Len(txtPOFrDate.Text) = 0 Then
    MsgBox "Type Starting Date..", vbInformation, cnstMsgInfo
    txtPOFrDate.SetFocus
    boEmptyText = True
    Exit Sub
  End If

If Len(txtPOToDate.Text) = 0 Then
    MsgBox "Type Ending Date", vbInformation, cnstMsgInfo
    txtPOToDate.SetFocus
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

If frmPO.Tag = 2 Then
MainComm.CommandText = "Insert Into Temp_Mrr_Challan_Info(SuppName,MrrId,MRR_PONo,MrrDate,ItemName,ItemDes,PRNO,MRRQty,QtyUnit,UnitPrice,ItTotal,MRRCurr,CompID,compName,compAdd,UserName,ComputerName,EntryDate,EntryTime) " _
                    & " SELECT new_Supplier.suppName,Pro_Inv_Main.PIID,New_PI.PINo,Pro_Inv_Main.PIDate, new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty,Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " SUM(Pro_Inv_Sub.itTotal) AS ITTotal,Pro_Inv_Sub.Currency,Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,Sys_Configure_Status.ComName,Sys_User_Name.UserName,Pro_Inv_Sub.EntryDate,Pro_Inv_Sub.EntryTime FROM Company_Information INNER JOIN new_Item_Name INNER JOIN new_Supplier INNER JOIN Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID ON " _
                    & " new_Supplier.supplierID = Pro_Inv_Main.suppID ON new_Item_Name.itemID = Pro_Inv_Sub.itemID ON " _
                    & " Company_Information.CompID = Pro_Inv_Main.compID INNER JOIN Sys_Configure_Status ON Pro_Inv_Sub.PCName = Sys_Configure_Status.ComName INNER JOIN InventoryPJL.dbo.Sys_User_Name ON Pro_Inv_Sub.UserId = InventoryPJL.dbo.Sys_User_Name.UserID Inner Join New_PI On Pro_Inv_Main.PIID=New_PI.PIID " _
                    & " WHERE (Pro_Inv_Main.compID = " & bytCompID & ") AND (Pro_Inv_Main.PIDate Between '" & Format(txtPOFrDate.Text, cnstDtFrmtQ) & "' And '" & Format(txtPOToDate.Text, cnstDtFrmtQ) & "') AND (Pro_Inv_Sub.DeleteRow = 2) " _
                    & " GROUP BY new_Supplier.suppName,Pro_Inv_Main.PIID,New_PI.PINo,Pro_Inv_Main.PIDate, new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty,Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " Pro_Inv_Sub.Currency , Company_Information.CompID, Company_Information.CompName, Company_Information.CompAdd, Sys_Configure_Status.ComName, InventoryPJL.dbo.Sys_User_Name.UserName, Pro_Inv_Sub.EntryDate, Pro_Inv_Sub.EntryTime;"

ElseIf frmPO.Tag = 1 Then
MainComm.CommandText = "Insert Into Temp_Mrr_Challan_Info(SuppName,MrrId,MRR_PONo,MrrDate,ItemName,ItemDes,PRNO,MRRQty,QtyUnit,UnitPrice,ItTotal,MRRCurr,CompID,compName,compAdd,UserName,ComputerName,EntryDate,EntryTime) " _
                    & " SELECT new_Supplier.suppName,Pro_Inv_Main.PIID,New_PI.PINo,Pro_Inv_Main.PIDate, new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty,Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " SUM(Pro_Inv_Sub.itTotal) AS ITTotal,Pro_Inv_Sub.Currency,Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,Sys_Configure_Status.ComName,Sys_User_Name.UserName,Pro_Inv_Sub.EntryDate,Pro_Inv_Sub.EntryTime FROM Company_Information INNER JOIN new_Item_Name INNER JOIN new_Supplier INNER JOIN Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID ON " _
                    & " new_Supplier.supplierID = Pro_Inv_Main.suppID ON new_Item_Name.itemID = Pro_Inv_Sub.itemID ON " _
                    & " Company_Information.CompID = Pro_Inv_Main.compID INNER JOIN Sys_Configure_Status ON Pro_Inv_Sub.PCName = Sys_Configure_Status.ComName INNER JOIN InventoryPJL.dbo.Sys_User_Name ON Pro_Inv_Sub.UserId = InventoryPJL.dbo.Sys_User_Name.UserID Inner Join New_PI On Pro_Inv_Main.PIID=New_PI.PIID " _
                    & " WHERE (Pro_Inv_Main.compID = " & bytCompID & ") AND (Pro_Inv_Main.PIDate Between '" & Format(txtPOFrDate.Text, cnstDtFrmtQ) & "' And '" & Format(txtPOToDate.Text, cnstDtFrmtQ) & "') AND (Pro_Inv_Sub.DeleteRow = 1) " _
                    & " GROUP BY new_Supplier.suppName,Pro_Inv_Main.PIID,New_PI.PINo,Pro_Inv_Main.PIDate, new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty,Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " Pro_Inv_Sub.Currency , Company_Information.CompID, Company_Information.CompName, Company_Information.CompAdd, Sys_Configure_Status.ComName, InventoryPJL.dbo.Sys_User_Name.UserName, Pro_Inv_Sub.EntryDate, Pro_Inv_Sub.EntryTime;"



ElseIf frmPO.Tag = 3 Then
MainComm.CommandText = "Insert Into Temp_Mrr_Challan_Info(SuppName,MrrId,MRR_PONo,MrrDate,ItemName,ItemDes,PRNO,MRRQty,QtyUnit,UnitPrice,ItTotal,MRRCurr,CompID,compName,compAdd) " _
                    & " SELECT new_Supplier.suppName,Pro_Inv_Main.PIID,New_PI.PINo,Pro_Inv_Main.PIDate, new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty,Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " SUM(Pro_Inv_Sub.itTotal) AS ITTotal,Pro_Inv_Sub.Currency,Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd FROM Company_Information INNER JOIN new_Item_Name INNER JOIN new_Supplier INNER JOIN Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID ON " _
                    & " new_Supplier.supplierID = Pro_Inv_Main.suppID ON new_Item_Name.itemID = Pro_Inv_Sub.itemID ON " _
                    & " Company_Information.CompID = Pro_Inv_Main.compID  Join New_PI On Pro_Inv_Main.PIID=New_PI.PIID " _
                    & " WHERE (Pro_Inv_Main.compID = " & bytCompID & ") AND (Pro_Inv_Main.PIDate Between '" & Format(txtPOFrDate.Text, cnstDtFrmtQ) & "' And '" & Format(txtPOToDate.Text, cnstDtFrmtQ) & "') AND (Pro_Inv_Sub.DeleteRow = 0) " _
                    & " GROUP BY new_Supplier.suppName,Pro_Inv_Main.PIID,New_PI.PINo,Pro_Inv_Main.PIDate, new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty,Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " Pro_Inv_Sub.Currency , Company_Information.CompID, Company_Information.CompName, Company_Information.CompAdd, Pro_Inv_Sub.EntryDate, Pro_Inv_Sub.EntryTime order by Pro_Inv_Main.PIDate;"










End If

MainComm.Execute

If frmPO.Tag = 1 Then
MainComm.CommandText = "Update Temp_Mrr_Challan_Info Set HeaderInfo='PO DELETE INFORMATION'"

ElseIf frmPO.Tag = 2 Then
MainComm.CommandText = "Update Temp_Mrr_Challan_Info Set HeaderInfo='PO EDIT INFORMATION'"


Else
MainComm.CommandText = "Update Temp_Mrr_Challan_Info Set HeaderInfo=' ALL PO INFORMATION'"
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
