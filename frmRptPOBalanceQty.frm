VERSION 5.00
Begin VB.Form frmRptPOBalanceQty 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Period  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1807
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   5940
      Begin VB.OptionButton OptChallanWise 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CHALLAN wise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   182
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   1605
      End
      Begin VB.OptionButton OptMRR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR wise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   182
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   1131
      End
      Begin VB.OptionButton OptPO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO wise "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   182
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1001
      End
      Begin VB.TextBox txtFrDate 
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
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   2041
      End
      Begin VB.TextBox txtToDate 
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
         Left            =   3600
         TabIndex        =   10
         Top             =   1080
         Width           =   2041
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date From : "
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
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date To : "
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
         Height          =   180
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   900
      End
   End
   Begin VB.ComboBox cmbPONO 
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
      Left            =   1920
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3360
      Width           =   2865
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
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "Preview Report"
      Top             =   4560
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
      Left            =   3120
      TabIndex        =   0
      ToolTipText     =   "Close Form"
      Top             =   4560
      Width           =   1185
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PO No. :"
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
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3030
      Width           =   1365
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   975
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   4245
      Width           =   6315
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   6315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: P.O. Balance Quantity"
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
      Left            =   135
      TabIndex        =   5
      Top             =   210
      Width           =   4125
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
      TabIndex        =   4
      Top             =   840
      Width           =   6330
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "frmRptPOBalanceQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRDate1 As String
Dim STRDate2 As String
Dim STRCritaria As String

Private Sub cmbPONO_GotFocus()
    
If STRDate1 <> txtFrDate.Text Or STRDate2 <> txtToDate.Text Or STRCritaria <> Frame1.Tag Or cmbPONO.ListCount = 0 Then
    cmbPONO.Clear
    Set R = New ADODB.Recordset
    'strRecord = "SELECT PINO FROM New_PI  ;"
    If Len(txtFrDate.Text) > 0 And Len(txtToDate.Text) > 0 Then
    strRecord = "SELECT Distinct PINo FROM DateWisePO Where " & Frame1.Tag & " Between '" & txtFrDate.Text & "' And '" & txtToDate.Text & "';"
    Else
    strRecord = "SELECT Distinct PINo FROM DateWisePO ;"
    End If
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cmbPONO.AddItem R![PINo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
Else
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
Set frmRptPOBalanceQty = Nothing
End Sub

Private Sub cmdReport_Click()
    If Len(cmbPONO.Text) = 0 Then
        MsgBox "Select P.O. No.", vbInformation, "Message"
        cmbPONO.SetFocus
        Exit Sub
    End If
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    On Error GoTo ErrorHandler
'    Screen.MousePointer = vbHourglass
'    MainConn.BeginTrans
'
'    MainComm.CommandText = "Delete From Temp_PO_Balance_Qty;"
'    MainComm.Execute
    
          
' MainComm.CommandText = "Insert Into Temp_PO_Balance_Qty(CompID,ItemType,ItemName,ITDesc,PONo,MRRNo,PRNO,POQty,RevdQty,Balance)" _
'& " SELECT    " & bytCompID & ", B.itemType, B.itemName, dbo.Pro_Inv_Sub.ItDesc as itdesc,dbo.New_PI.PINo " _
'& " ,'1' AS MRR, dbo.Pro_Inv_Sub.PRNo AS prno, dbo.Pro_Inv_Sub.PIQty,0,dbo.Pro_Inv_Sub.PIQty AS BALANCE " _
'& " FROM dbo.Pro_Inv_Sub INNER JOIN  dbo.new_Item_Name B ON dbo.Pro_Inv_Sub.itemID = B.itemID INNER JOIN " _
'& " dbo.New_PI ON dbo.Pro_Inv_Sub.PIID = dbo.New_PI.PIID " _
'& " WHERE (dbo.New_PI.PINo = '" & cmbPONO.Text & "') AND (dbo.Pro_Inv_Sub.DeleteRow = 0) " _
'& " Union SELECT " & bytCompID & ", D.itemType, B.itemName, A.ItDesc, E.PINo, C.MRRNo, A.PRNO, 0,RcvdQty, - A.RcvdQty AS BALANCE " _
'& " FROM  dbo.MRR_Challan_Sub A INNER JOIN  dbo.new_Item_Name B ON A.itemID = B.itemID INNER JOIN " _
'& " dbo.new_Item_Type D ON B.itemType = D.itemType INNER JOIN dbo.New_MRR C ON A.MRRID = C.MRRID INNER JOIN " _
'& " dbo.New_PI E ON E.PIID = A.PIId WHERE  (A.DeleteRow = 0)  AND (E.DeleteRow = 0) AND (E.PINo = '" & cmbPONO.Text & "')"
   
    
    Set R1 = New ADODB.Recordset
   
    strRecord = "SELECT * FROM PO_Balance_Qty WHERE (PONo = '" & cmbPONO.Text & "')"
    
    
    R1.Open strRecord, MainConn, adOpenStatic
   
    
'    MainComm.Execute
'    MainConn.CommitTrans
    
ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  boEmptyText = True
  Exit Sub
End If
    
    Set REPORT = Appl.OpenReport(findReportPath & "\rptPOBalanceQty.rpt")
        
    REPORT.DiscardSavedData
    
    REPORT.Database.SetDataSource R1
    
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = REPORT
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    
    R1.Close
    
    Set R1 = Nothing
    
'CrystalReport1.ReportFileName = findReportPath + "\rptPOBalanceQty.rpt"
'CrystalReport1.PrintReport
'Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    OptPO_Click
'    Call prcConReport(CrystalReport1)
'    Call addPONO(cmbPONO)
End Sub

Private Sub OptChallanWise_Click()
  Frame1.Tag = "CHALLANDATE"
  STRCritaria = Frame1.Tag
End Sub

Private Sub OptMRR_Click()
  Frame1.Tag = "MRRDATE"
  STRCritaria = Frame1.Tag
End Sub

Private Sub OptPO_Click()
  Frame1.Tag = "PIDATE"
  STRCritaria = Frame1.Tag
End Sub

Private Sub txtFrDate_LostFocus()
    txtFrDate.Text = FormatDate(txtFrDate)
    STRDate1 = txtFrDate.Text
End Sub

Private Sub txtToDate_LostFocus()
    txtToDate.Text = FormatDate(txtToDate)
    STRDate2 = txtToDate.Text
End Sub
