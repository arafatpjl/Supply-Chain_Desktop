VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmItemName 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5889
   ClientLeft      =   26
   ClientTop       =   -26
   ClientWidth     =   5330
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5889
   ScaleWidth      =   5330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmsupp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supplier "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4979
      TabIndex        =   16
      Top             =   793
      Visible         =   0   'False
      Width           =   4381
      Begin MSDataListLib.DataCombo cmbSupplier 
         Height          =   299
         Left            =   169
         TabIndex        =   0
         Top             =   312
         Width           =   4030
         _ExtentX        =   7428
         _ExtentY        =   527
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Period  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1807
      Left            =   845
      TabIndex        =   15
      Top             =   3003
      Width           =   3653
      Begin VB.CheckBox chkDetails 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO wise Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1326
         TabIndex        =   10
         Top             =   1378
         Width           =   1573
      End
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1326
         TabIndex        =   6
         Top             =   962
         Width           =   2041
      End
      Begin VB.TextBox txtFrDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1326
         TabIndex        =   5
         Top             =   572
         Width           =   2041
      End
      Begin VB.OptionButton OptPO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO wise "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   182
         Left            =   1287
         TabIndex        =   3
         Top             =   234
         Value           =   -1  'True
         Width           =   1001
      End
      Begin VB.OptionButton OptMRR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR wise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   182
         Left            =   2366
         TabIndex        =   4
         Top             =   234
         Width           =   1131
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "To : "
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
         Height          =   182
         Left            =   208
         TabIndex        =   20
         Top             =   962
         Width           =   1027
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date From : "
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
         Height          =   182
         Left            =   208
         TabIndex        =   19
         Top             =   572
         Width           =   1027
      End
   End
   Begin VB.CheckBox chkGroup 
      BackColor       =   &H00E26101&
      Caption         =   "GROUP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.79
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   364
      TabIndex        =   9
      Top             =   5408
      Width           =   949
   End
   Begin VB.Frame frmitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Item "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1521
      Left            =   480
      TabIndex        =   14
      Top             =   1157
      Width           =   4355
      Begin MSDataListLib.DataCombo cmbItemName 
         Height          =   299
         Left            =   130
         TabIndex        =   2
         Top             =   1092
         Width           =   4056
         _ExtentX        =   7476
         _ExtentY        =   527
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbItemGroup 
         Height          =   299
         Left            =   130
         TabIndex        =   1
         Top             =   520
         Width           =   4056
         _ExtentX        =   7476
         _ExtentY        =   527
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item Group : "
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
         Height          =   182
         Left            =   130
         TabIndex        =   18
         Top             =   234
         Width           =   1235
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item Name : "
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
         Height          =   182
         Left            =   156
         TabIndex        =   17
         Top             =   832
         Width           =   1235
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1495
      TabIndex        =   8
      ToolTipText     =   "Close Form"
      Top             =   5330
      Width           =   1185
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00C0C000&
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2964
      TabIndex        =   7
      ToolTipText     =   "Preview Report"
      Top             =   5343
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report - PO wise Item Details "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   325
      Index           =   2
      Left            =   169
      TabIndex        =   13
      Top             =   156
      Width           =   4134
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   741
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   5187
      Width           =   6318
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   702
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6318
   End
End
Attribute VB_Name = "frmItemName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbItemGroup_LostFocus()
  cmbItemName.Text = ""
  If Len(Trim(cmbItemGroup)) = 0 Then Exit Sub
  Set R = New ADODB.Recordset
  strRecord = "SELECT itemID,itemName FROM new_Item_Name WHERE itemType = '" & cmbItemGroup.Text & "' ORDER BY itemName;"
  R.Open strRecord, MainConn, adOpenStatic
  Set cmbItemName.RowSource = R
  cmbItemName.ListField = "itemName"
  cmbItemName.BoundColumn = "itemID"
  Set R = Nothing
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdReport_Click()

''If Len(Trim(cmbItemGroup)) = 0 Then
''  MsgBox "Enter Group of Item", vbInformation, cnstMsgInfo
''  cmbItemGroup.SetFocus
''  Exit Sub
''End If

If chkDetails.Value = 1 Then
  If Len(Trim(txtFrDate)) = 0 Then
    MsgBox "Enter From Date", vbInformation, cnstMsgInfo
    txtFrDate.SetFocus
    Exit Sub
  End If
  If Len(Trim(txtToDate)) = 0 Then
    MsgBox "Enter To Date", vbInformation, cnstMsgInfo
    txtToDate.SetFocus
    Exit Sub
  End If
End If

Set R1 = New ADODB.Recordset

If OptPO.Value = True And chkDetails.Value = 1 Then
  Screen.MousePointer = vbHourglass
  
  Set Report = Appl.OpenReport(findReportPath & "\rpt_PO_MRR_ItemStatement.rpt")
  
  If Len(Trim(txtFrDate)) = 0 Then
    
    If Len(Trim(cmbItemGroup)) = 0 And Len(Trim(cmbItemName)) = 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "';"
    ElseIf Len(Trim(cmbItemGroup)) > 0 And Len(Trim(cmbItemName)) = 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND itemType = '" & cmbItemGroup.Text & "';"
    ElseIf Len(Trim(cmbItemGroup)) > 0 And Len(Trim(cmbItemName)) > 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND itemType = '" & cmbItemGroup.Text & "' AND itemName = '" & cmbItemName.Text & "';"
    ElseIf Len(Trim(cmbItemGroup)) = 0 And Len(Trim(cmbItemName)) > 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND itemName = '" & cmbItemName.Text & "';"
    End If
  
  Else
    If Len(Trim(cmbItemGroup)) = 0 And Len(Trim(cmbItemName)) = 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "';"
    ElseIf Len(Trim(cmbItemGroup)) > 0 And Len(Trim(cmbItemName)) = 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "' AND itemType = '" & cmbItemGroup.Text & "';"
    ElseIf Len(Trim(cmbItemGroup)) > 0 And Len(Trim(cmbItemName)) > 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "' AND itemType = '" & cmbItemGroup.Text & "' AND itemName = '" & cmbItemName.Text & "';"
    ElseIf Len(Trim(cmbItemGroup)) = 0 And Len(Trim(cmbItemName)) > 0 Then
      strRecord = "SELECT * FROM View_PO_MRR_ItemStatement WHERE PICompID = '" & bytCompID & "' AND PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "' AND itemName = '" & cmbItemName.Text & "';"
    End If
  End If
  
  R1.Open strRecord, MainConn, adOpenStatic
  Report.Database.SetDataSource R1

  frmMainReport.CRVIEWER.Refresh
  frmMainReport.CRVIEWER.ReportSource = Report
  frmMainReport.CRVIEWER.ViewReport
  frmMainReport.CRVIEWER.Zoom (82)
  frmMainReport.Show 1
  
  R1.Close
  Set R1 = Nothing
  Screen.MousePointer = vbDefault
  Exit Sub
End If

If Len(Trim(cmbItemName)) = 0 Then
  MsgBox "Enter Name of Item", vbInformation, cnstMsgInfo
  cmbItemName.SetFocus
  Exit Sub
End If

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

''MainConn.BeginTrans

MainComm.CommandText = "DELETE FROM TEMP_ITEM_NAME;"
MainComm.Execute

If Len(Trim(txtFrDate)) = 0 And Len(Trim(txtToDate)) = 0 Then
  If chkGroup.Value = 0 Then
    If OptPO.Value = True Then
      Set Report = Appl.OpenReport(findReportPath & "\rptitemname.rpt")
      strRecord = "SELECT * FROM VIEW_ITEMNAME WHERE COMPID = '" & bytCompID & "' AND ITEMNAME = '" & cmbItemName.Text & "' AND DELETEROW = 0 ORDER BY PIDATE;"
    Else
      Set Report = Appl.OpenReport(findReportPath & "\rptitemnameMRR.rpt")
      strRecord = "SELECT * FROM VIEW_ITEMNAME_MRR WHERE COMPID = '" & bytCompID & "' AND ITEMNAME = '" & cmbItemName.Text & "' AND DELETEROW = 0 ORDER BY MRRDate;"
    End If

  '----ALL COMPANY FOR SINGLE ITEM
  Else
    
    If OptPO.Value = True Then
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJL.dbo.VIEW_ITEMNAME WHERE COMPID='1' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
 
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryUJL.dbo.VIEW_ITEMNAME WHERE COMPID='7' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryJ2L.dbo.VIEW_ITEMNAME WHERE COMPID='5' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPAL.dbo.VIEW_ITEMNAME WHERE COMPID='3' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHT.dbo.VIEW_ITEMNAME WHERE COMPID='13' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHTHOLD.dbo.VIEW_ITEMNAME WHERE COMPID='14' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLCA.dbo.VIEW_ITEMNAME WHERE COMPID='15' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLFG.dbo.VIEW_ITEMNAME WHERE COMPID='16' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPWP.dbo.VIEW_ITEMNAME WHERE COMPID='21' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPCL.dbo.VIEW_ITEMNAME WHERE COMPID='22' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNZNF.dbo.VIEW_ITEMNAME WHERE COMPID='17' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryHUAL.dbo.VIEW_ITEMNAME WHERE COMPID='20' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
      
      Set Report = Appl.OpenReport(findReportPath & "\rptitemname_Allcompany.rpt")
    
    Else

      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='1' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryUJL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='7' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryJ2L.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='5' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPAL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='3' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHT.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='13' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHTHOLD.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='14' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLCA.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='15' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLFG.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='16' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPWP.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='21' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPCL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='22' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNZNF.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='17' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryHUAL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='20' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      Set Report = Appl.OpenReport(findReportPath & "\rptitemnameMRR_Allcompany.rpt")
    
    End If
  End If

  '----
Else
  
  If Len(Trim(txtFrDate)) = 0 Then
    MsgBox "Enter From Date", vbInformation, cnstMsgInfo
    txtFrDate.SetFocus
    Exit Sub
  End If
  If Len(Trim(txtToDate)) = 0 Then
    MsgBox "Enter To Date", vbInformation, cnstMsgInfo
    txtToDate.SetFocus
    Exit Sub
  End If

  If OptPO.Value = True Then
    Set Report = Appl.OpenReport(findReportPath & "\rptitemname.rpt")
    strRecord = "SELECT * FROM VIEW_ITEMNAME WHERE COMPID='" & bytCompID & "' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0  AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') ORDER BY PIDATE "
  Else
    Set Report = Appl.OpenReport(findReportPath & "\rptitemnameMRR.rpt")
    strRecord = "SELECT * FROM VIEW_ITEMNAME_MRR WHERE COMPID='" & bytCompID & "' AND ITEMNAME='" & cmbItemName.Text & "' AND DELETEROW=0 AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "')  ORDER BY MRRDate "
  End If

  '----ALL COMPANY FOR SINGLE ITEM

  If chkGroup.Value = 1 Then
    If OptPO.Value = True Then
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJL.dbo.VIEW_ITEMNAME WHERE COMPID='1' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
    
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryUJL.dbo.VIEW_ITEMNAME WHERE COMPID='7' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryJ2L.dbo.VIEW_ITEMNAME WHERE COMPID='5' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPAL.dbo.VIEW_ITEMNAME WHERE COMPID='3' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHT.dbo.VIEW_ITEMNAME WHERE COMPID='13' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHTHOLD.dbo.VIEW_ITEMNAME WHERE COMPID='14' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLCA.dbo.VIEW_ITEMNAME WHERE COMPID='15' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLFG.dbo.VIEW_ITEMNAME WHERE COMPID='16' AND ITEMNAME='" & cmbItemName.Text & "'  AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPWP.dbo.VIEW_ITEMNAME WHERE COMPID='21' AND ITEMNAME='" & cmbItemName.Text & "'  AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "')  AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPCL.dbo.VIEW_ITEMNAME WHERE COMPID='22' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "')  AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNZNF.dbo.VIEW_ITEMNAME WHERE COMPID='17' AND ITEMNAME='" & cmbItemName.Text & "' AND (PIDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY PIDATE "
      MainComm.Execute
     
      Set Report = Appl.OpenReport(findReportPath & "\rptitemname_Allcompany.rpt")
  
    Else
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='1' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryUJL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='7' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
    
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryJ2L.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='5' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPAL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='3' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHT.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='13' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNHTHOLD.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='14' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLCA.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='15' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPJLFG.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='16' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPWP.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='21' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryPCL.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='22' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO TEMP_ITEM_NAME SELECT * FROM InventoryNZNF.dbo.VIEW_ITEMNAME_MRR WHERE COMPID='17' AND ITEMNAME='" & cmbItemName.Text & "' AND (MRRDATE BETWEEN '" & Format(txtFrDate.Text, cnstDtFrmtQ) & "' AND '" & Format(txtToDate.Text, cnstDtFrmtQ) & "') AND DELETEROW=0   ORDER BY MRRDATE "
      MainComm.Execute
        
      Set Report = Appl.OpenReport(findReportPath & "\rptitemnameMRR_Allcompany.rpt")
    End If
  End If
  '----
End If

'MainConn.CommitTrans
'
'
ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
'  MainConn.RollbackTrans
  boEmptyText = True
  Exit Sub
End If


If chkGroup.Value = 0 Then
'    R1.Open , MainConn, 3, 3
'    Report.DiscardSavedData
'    Report.Database.SetDataSource R1
  'Set R = New ADODB.Recordset
  R1.Open strRecord, MainConn, adOpenStatic
  Report.Database.SetDataSource R1

  frmMainReport.CRVIEWER.Refresh
  frmMainReport.CRVIEWER.ReportSource = Report
  frmMainReport.CRVIEWER.ViewReport
  frmMainReport.CRVIEWER.Zoom (82)
  frmMainReport.Show 1
  R1.Close
  Set R1 = Nothing

Else
  frmMainReport.CRVIEWER.Refresh
  frmMainReport.CRVIEWER.ReportSource = Report
  frmMainReport.CRVIEWER.ViewReport
  frmMainReport.CRVIEWER.Zoom (82)
  frmMainReport.Show 1
End If

End Sub
Private Sub Form_Load()

Call addSupplierNameD(cmbSupplier)

Set R = New ADODB.Recordset
strRecord = "SELECT DISTINCT itemType FROM new_Item_Name ORDER BY itemType;"
R.Open strRecord, MainConn, adOpenStatic
Set cmbItemGroup.RowSource = R
cmbItemGroup.ListField = "itemType"
cmbItemGroup.BoundColumn = "itemType"
R.Close

strRecord = "SELECT itemName FROM new_Item_Name ORDER BY itemName;"
R.Open strRecord, MainConn, adOpenStatic
Set cmbItemName.RowSource = R
cmbItemName.ListField = "itemName"
cmbItemName.BoundColumn = "itemName"
Set R = Nothing
End Sub
Private Sub frmitem_Click()
''  frmitem.Enabled = True
''  frmsupp.Enabled = False
End Sub
Private Sub frmsupp_Click()
''  frmitem.Enabled = False
''  frmsupp.Enabled = True
End Sub
Private Sub txtFrDate_GotFocus()
  Call txtGotFocus(txtFrDate)
End Sub
Private Sub txtFrDate_LostFocus()
  txtFrDate.Text = FormatDate(txtFrDate)
End Sub
Private Sub txtToDate_GotFocus()
  Call txtGotFocus(txtToDate)
End Sub
Private Sub txtToDate_LostFocus()
  txtToDate.Text = FormatDate(txtToDate)
End Sub
