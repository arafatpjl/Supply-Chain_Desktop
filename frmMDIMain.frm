VERSION 5.00
Begin VB.MDIForm frmMDIMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00000000&
   Caption         =   "Supply Chain Management"
   ClientHeight    =   10647
   ClientLeft      =   65
   ClientTop       =   -7228
   ClientWidth     =   15236
   Icon            =   "frmMDIMain.frx":0000
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   14055
      Left            =   0
      ScaleHeight     =   14001
      ScaleWidth      =   15184
      TabIndex        =   2
      Top             =   546
      Width           =   15236
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Chain"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A0C05&
         Height          =   975
         Left            =   2520
         TabIndex        =   4
         Top             =   3600
         Width           =   6495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Management System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007A0C05&
         Height          =   975
         Left            =   4560
         TabIndex        =   3
         Top             =   4800
         Width           =   7815
      End
      Begin VB.Image Image1 
         Height          =   4380
         Left            =   2280
         Picture         =   "frmMDIMain.frx":1782
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   10200
      End
   End
   Begin VB.PictureBox Banner1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   0
      ScaleHeight     =   520
      ScaleWidth      =   15210
      TabIndex        =   0
      Top             =   0
      Width           =   15236
      Begin VB.Label lblCompName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E26101&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   16.3
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   455
         Left            =   533
         TabIndex        =   1
         Top             =   26
         UseMnemonic     =   0   'False
         Width           =   5980
      End
   End
   Begin VB.Menu mnuHouseKeeping 
      Caption         =   "  &House Keeping"
      Begin VB.Menu mnuNewItemName 
         Caption         =   "New Item Name"
      End
      Begin VB.Menu mnuItemGroup 
         Caption         =   "Item Group"
      End
      Begin VB.Menu mnuNewSupplier 
         Caption         =   "New Supplier"
      End
      Begin VB.Menu mnuSuppInformation 
         Caption         =   "Supplier Information (Bill)"
      End
      Begin VB.Menu mnuNewParty 
         Caption         =   "New Party (Bill)"
      End
      Begin VB.Menu mnuNewPFNo 
         Caption         =   "New File No."
      End
      Begin VB.Menu mnuacchead 
         Caption         =   "Accounts Head"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "   &Data Entry"
      Begin VB.Menu mnuSupplyChain 
         Caption         =   "Supply Chain"
         Begin VB.Menu mnuPurchaseOrderOC 
            Caption         =   "P.O Other Curr."
         End
         Begin VB.Menu mnuPurchaseOrder 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuthread 
            Caption         =   "P.O.(THREAD)"
         End
         Begin VB.Menu MnuNext08 
            Caption         =   "-"
         End
         Begin VB.Menu mencashcr 
            Caption         =   "Cash And Credit"
         End
         Begin VB.Menu menuchk 
            Caption         =   "Cheque And Cash Cheque"
         End
      End
      Begin VB.Menu mnuNext01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStore 
         Caption         =   "Store"
         Begin VB.Menu mnumrrinvchallan 
            Caption         =   "MRR Information"
         End
         Begin VB.Menu mnuMRRInfoOthCur 
            Caption         =   "MRR Information Other Currency"
         End
      End
      Begin VB.Menu menunext09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaccount 
         Caption         =   "Accounts"
         Begin VB.Menu menuchkinfo 
            Caption         =   "Cheque Info"
         End
         Begin VB.Menu menuBillNOCash 
            Caption         =   "Bill Details Cash"
         End
         Begin VB.Menu mnuBillCr 
            Caption         =   "Bill Detalis Credit"
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "   &Reports  "
      Begin VB.Menu mnurptSupplyChain 
         Caption         =   "Supply Chain"
         Begin VB.Menu mnurptpurchaseinvoice 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnurptPOBalanceQty 
            Caption         =   "PO && MRR Balance Qty"
         End
         Begin VB.Menu mnuRptBillwise 
            Caption         =   "Bill Wise Report"
         End
         Begin VB.Menu mnuRptPartystatement 
            Caption         =   "Party Statement"
         End
         Begin VB.Menu mnuMRRwisereport 
            Caption         =   "MRR Wise Report"
         End
         Begin VB.Menu mnuRptTransactionwisereport 
            Caption         =   "Transaction Wise Report"
         End
         Begin VB.Menu mnuRptChallanwisereport 
            Caption         =   "Challan Wise Report"
         End
         Begin VB.Menu mnuitemname 
            Caption         =   "PO wise Item Details"
         End
         Begin VB.Menu mnuMonthlystatement 
            Caption         =   "Monthly Statement"
         End
      End
      Begin VB.Menu mnunext02 
         Caption         =   "-"
      End
      Begin VB.Menu mnurptStore 
         Caption         =   "Store"
         Begin VB.Menu mnurptmrrchallan 
            Caption         =   "MRR Information"
         End
      End
      Begin VB.Menu mnuNext03 
         Caption         =   "-"
      End
      Begin VB.Menu mnurptaudit 
         Caption         =   "Audit"
         Begin VB.Menu mnurptmrrEditDelete 
            Caption         =   "MRR Information"
         End
         Begin VB.Menu mnurptchallanEditDelete 
            Caption         =   "Challan Information"
         End
         Begin VB.Menu MnuNext07 
            Caption         =   "-"
         End
         Begin VB.Menu mnurptPOEditDelete 
            Caption         =   "PO Information"
         End
      End
      Begin VB.Menu mnuNext010 
         Caption         =   "-"
      End
      Begin VB.Menu menurptAccount 
         Caption         =   "Accounts"
         Begin VB.Menu mnurptcheque 
            Caption         =   "Cheque Info"
         End
         Begin VB.Menu mnurptbilcash 
            Caption         =   "Bill Details Cash"
         End
         Begin VB.Menu mnurptbilcr 
            Caption         =   "Bill Details Credit"
         End
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "   Op&tion   "
      Begin VB.Menu mnuModifyComp 
         Caption         =   "Modify Company"
      End
      Begin VB.Menu mnuDeleteComp 
         Caption         =   "Delete Company"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNext04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateUser 
         Caption         =   "Create or Modify User"
      End
      Begin VB.Menu mnuNext05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupPermission 
         Caption         =   "Setup Permission "
         Begin VB.Menu mnuUserCompany 
            Caption         =   "User -> Company"
         End
         Begin VB.Menu mnuUserMenu 
            Caption         =   "User -> Menu"
         End
      End
      Begin VB.Menu mnuNext06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptLockData 
         Caption         =   "Lock Data "
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "   E&xit   "
      Begin VB.Menu mnuQuit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
'   lblCompName.Caption = "Inventory"     ''''&H00BFFBE9&
    lblCompName.Caption = strComNameTxt
    lblCompName.ForeColor = &H0&
End Sub

Private Sub MDIForm_Resize()
    lblCompName.Width = Me.Width
    If WindowState = 0 Then WindowState = 2
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    frmMDIMain.lblCompName = ""
    Set frmMDIMain = Nothing
End Sub

Private Sub mnuAbsentEntry_Click()
    frmEmpAbsentInfo.Show 1
End Sub

Private Sub mencashcr_Click()
  frmStatementCash_and_Credit.Show 1
End Sub

Private Sub menuBillNOCash_Click()
  frmBillNOCash.Show 1
End Sub

Private Sub menuchk_Click()
  frmStatementCheque.Show 1
End Sub

Private Sub menuchkinfo_Click()
 frmChequeInfo.Show 1
End Sub

Private Sub mnuacchead_Click()
  frmAccountsHead.Show 1
End Sub

Private Sub mnuBillCr_Click()
  frmBillNO.Show 1
  ''frmBillEntryAcc.Show 1
End Sub

Private Sub mnuCreateUser_Click()
    frmCreateUser.Show 1
End Sub

Private Sub mnuInsertCalendar_Click()
    frmRptCalendar.Show 1
End Sub

Private Sub mnuItemGroup_Click()
    frmNewItemType.Show 1
End Sub

Private Sub mnuitemname_Click()
  frmItemName.Show 1
End Sub

Private Sub mnuModifyComp_Click()
    frmModifyCompany.Show 1
End Sub

Private Sub mnuNewComp_Click()
    frmCreateCompany.Show 1
End Sub

Private Sub mnuMonthlystatement_Click()
frmrptMonthlyStatement.Show 1
End Sub

Private Sub mnuMRRInfoOthCur_Click()
  frmMrrReceiveChallanOtherCurr.Show 1
End Sub

Private Sub mnumrrinvchallan_Click()
  frmMrrReceiveChallan.Show 1
End Sub

Private Sub mnuMRRwisereport_Click()
  frmrptMRRWiseReport.Show 1
End Sub

Private Sub mnuNewItemName_Click()
    frmNewItemName.Show 1
End Sub

Private Sub mnuNewParty_Click()
  ''frmNewParty2.Show 1
End Sub

Private Sub mnuNewPFNo_Click()
    frmNewPFNo.Show 1
End Sub

Private Sub mnuNewSupplier_Click()
  frmNewSupplier.Show 1
End Sub

Private Sub mnuoptLockData_Click()
    frmLockData.Show 1
End Sub

Private Sub mnuProformaInvoice_Click()
  frmProformaInvoice.Show 1
End Sub

Private Sub mnuPurchaseOrder_Click()
  frmProformaInvoice.Show 1
End Sub

Private Sub mnuPurchaseOrderOC_Click()
  frmProformaInvoiceOC.Show 1
End Sub

Private Sub mnuQuit_Click()
    If MsgBox("Are you sure to Exit", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    End
End Sub

Private Sub mnurptbilcash_Click()
  rptBillDetailsCash.Show 1
End Sub

Private Sub mnurptbilcr_Click()
  rptBillDetailsCredit.Show 1
End Sub

Private Sub mnuRptBillwise_Click()
  frmrptBillWiseReport.Show 1
End Sub

Private Sub mnurptchallanEditDelete_Click()
  frmRptChallanEditDeleteInfo.Show 1
End Sub

Private Sub mnuRptChallanwisereport_Click()
  frmrptChalanWiseReport.Show 1
End Sub

Private Sub mnurptcheque_Click()
  frmrptChequeInfo.Show 1
End Sub

Private Sub mnurptmrrchallan_Click()
  frmrptMrrChallan.Show 1
End Sub

Private Sub mnurptmrrEditDelete_Click()
  frmRptMrrEditDeleteInfo.Show 1
End Sub

Private Sub mnurptproformainvoice_Click()
  frmRptProInvoice.Show 1
End Sub

Private Sub mnurptproinvregister_Click()
    frmRptProInvRegister.Show 1
End Sub

Private Sub mnuRptPartystatement_Click()
  frmrptPartyStatement.Show 1
End Sub

Private Sub mnurptPOBalanceQty_Click()
    frmRptPOBalanceQty.Show 1
End Sub

Private Sub mnurptPOEditDelete_Click()
    frmRptPOEditDeleteInfo.Show 1
End Sub

Private Sub mnurptpurchaseinvoice_Click()
    frmRptProInvoice.Show 1
End Sub

Private Sub mnuRptTransactionwisereport_Click()
    frmrptTransactionWiseReport.Show 1
End Sub

Private Sub mnuSuppInformation_Click()
  frmNewParty.Show 1
End Sub

Private Sub mnuthread_Click()
  frmPOThread.Show 1
End Sub

Private Sub mnuUserCompany_Click()
    frmPermissionCom.Show 1
End Sub

Private Sub mnuUserMenu_Click()
    frmPermissionMenu.Show 1
End Sub

