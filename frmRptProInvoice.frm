VERSION 5.00
Begin VB.Form frmRptProInvoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5395
   ClientLeft      =   26
   ClientTop       =   -78
   ClientWidth     =   5772
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmRptProInvoice.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5395
   ScaleWidth      =   5772
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkOther 
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.CheckBox chkpad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pad Report"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox cmbItemGroup 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   915
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3990
      Width           =   4035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   585
      Left            =   915
      TabIndex        =   10
      Top             =   1665
      Width           =   4035
      Begin VB.OptionButton optItemGroup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1785
      End
      Begin VB.OptionButton optPI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P.O. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   570
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C000&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4365
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00C0C000&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2985
      TabIndex        =   4
      ToolTipText     =   "Preview Report"
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
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
      Height          =   525
      Left            =   1575
      TabIndex        =   6
      ToolTipText     =   "Close Form"
      Top             =   4770
      Width           =   1185
   End
   Begin VB.ComboBox cboSupplier 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   915
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2685
      Width           =   4035
   End
   Begin VB.ComboBox cboPINo 
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
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   915
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   3315
      Width           =   4035
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   -15
      TabIndex        =   16
      Top             =   4605
      Width           =   6195
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   1245
      Left            =   0
      TabIndex        =   15
      Top             =   4560
      Width           =   6315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Group Name : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   195
      Left            =   915
      TabIndex        =   11
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supplier Name : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   195
      Left            =   915
      TabIndex        =   9
      Top             =   2280
      Width           =   1875
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "P.O.  No : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   195
      Left            =   915
      TabIndex        =   8
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Width           =   6315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report :: Purchase Order"
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
      Top             =   225
      Width           =   3135
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "frmRptProInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboPINo_Click()
    cboPINo.Tag = cboPINo.ItemData(cboPINo.ListIndex)
End Sub

Private Sub cboPINo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cboPINo_LostFocus()
  cboPINo.Text = UCase(Trim(cboPINo.Text))
End Sub
Private Sub cboSupplier_Click()
    cboSupplier.Text = UCase(Trim(cboSupplier.Text))
    If Len(cboSupplier) = 0 Then Exit Sub
    cboSupplier.Tag = cboSupplier.ItemData(cboSupplier.ListIndex)
    addSuppPINo cboPINo, cboSupplier.Text
End Sub
Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Check1_Click()

End Sub

Private Sub ChkOther_Click()
If Len(cboSupplier) = 0 Then Exit Sub
    If ChkOther.Value = 0 Then
    cboSupplier.Tag = cboSupplier.ItemData(cboSupplier.ListIndex)
    addSuppPINo cboPINo, cboSupplier.Text
    ElseIf ChkOther.Value = 1 Then
    cboSupplier.Tag = cboSupplier.ItemData(cboSupplier.ListIndex)
    addSuppPINoOC cboPINo, cboSupplier.Text
    End If
End Sub

Private Sub cmbItemGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdPrint_Click()
    Call prcShowReport("Print")
End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdReport_Click()
    Call prcShowReport("Preview")
End Sub

Private Sub cmdReport_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Form_Load()
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
'    Call prcConReport(CrystalReport1)
    addSupplierName cboSupplier
    addItemType cmbItemGroup
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmRptProInvoice = Nothing
End Sub
Sub generateReport()

    boEmptyText = False

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM Temp_Import_Pro_Invoice;"
        MainComm.Execute
    
        If optPI.Value = True Then
            If cboSupplier.Text = "" And cboPINo.Text = "" Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItName)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition,NEW_ITEM_NAME.ITEMNAME FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.ItemID = new_Item_Name.ItemID) " _
                    & " WHERE Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                MainComm.Execute
            ElseIf cboSupplier.Text <> "" And cboPINo.Text = "" Then
                If ChkOther.Value = 0 Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItName)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition,NEW_ITEM_NAME.ITEMNAME FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.ItemID = new_Item_Name.ItemID) " _
                    & " WHERE Pro_Inv_Main.SuppID = " & cboSupplier.Tag & " AND Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                ElseIf ChkOther.Value = 1 Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItName)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_MainOC.PIDate,Pro_Inv_SubOC.PIQty, Pro_Inv_SubOC.QtyUnit,Pro_Inv_SubOC.UnitPrice, " _
                    & " Pro_Inv_SubOC.ItTotal,Pro_Inv_SubOC.ItDesc, Pro_Inv_SubOC.PRNo,Pro_Inv_SubOC.Currency,Pro_Inv_SubOC.Condition,NEW_ITEM_NAME.ITEMNAME FROM Pro_Inv_MainOC " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_MainOC.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_MainOC.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_MainOC.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_SubOC ON (Pro_Inv_MainOC.PIID = Pro_Inv_SubOC.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_SubOC.ItemID = new_Item_Name.ItemID) " _
                    & " WHERE Pro_Inv_MainOC.SuppID = " & cboSupplier.Tag & " AND Pro_Inv_MainOC.compID =" & bytCompID & " AND Pro_Inv_SubOC.DeleteRow = 0;"
                
                End If
                MainComm.Execute
            ElseIf cboSupplier.Text <> "" And cboPINo.Text <> "" Then
                If ChkOther.Value = 0 Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItName)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition,NEW_ITEM_NAME.ITEMNAME FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.ItemID = new_Item_Name.ItemID) " _
                    & " WHERE Pro_Inv_Main.PIID = " & cboPINo.Tag & " AND Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                    
                ElseIf ChkOther.Value = 1 Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItName)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_MainOC.PIDate,Pro_Inv_SubOC.PIQty, Pro_Inv_SubOC.QtyUnit,Pro_Inv_SubOC.UnitPrice, " _
                    & " Pro_Inv_SubOC.ItTotal,Pro_Inv_SubOC.ItDesc, Pro_Inv_SubOC.PRNo,Pro_Inv_SubOC.Currency,Pro_Inv_SubOC.Condition,NEW_ITEM_NAME.ITEMNAME FROM Pro_Inv_MainOC " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_MainOC.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_MainOC.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_MainOC.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_SubOC ON (Pro_Inv_MainOC.PIID = Pro_Inv_SubOC.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_SubOC.ItemID = new_Item_Name.ItemID) " _
                    & " WHERE Pro_Inv_MainOC.PIID = " & cboPINo.Tag & " AND Pro_Inv_MainOC.compID =" & bytCompID & " AND Pro_Inv_SubOC.DeleteRow = 0;"

                End If
                MainComm.Execute
            End If
        Else
            If cboSupplier.Text = "" And cmbItemGroup.Text = "" Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,ItName,Total,Remark,PRNo,PICurr,Condition)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " new_Item_Name.ItemName,Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.itemID = new_Item_Name.itemID) " _
                    & " Inner Join new_Item_Type On (new_Item_Name.ItemType = new_Item_Type.ItemType) " _
                    & " WHERE Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                MainComm.Execute
            ElseIf cboSupplier.Text <> "" And cmbItemGroup.Text = "" Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,ItName,Total,Remark,PRNo,PICurr,Condition)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " new_Item_Name.ItemName,Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.itemID = new_Item_Name.itemID) " _
                    & " Inner Join new_Item_Type On (new_Item_Name.ItemType = new_Item_Type.ItemType) " _
                    & " WHERE Pro_Inv_Main.SuppID = " & cboSupplier.Tag & " AND Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                MainComm.Execute
            ElseIf cboSupplier.Text <> "" And cmbItemGroup.Text <> "" Then
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,ItName,Total,Remark,PRNo,PICurr,Condition)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " new_Item_Name.ItemName,Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.itemID = new_Item_Name.itemID) " _
                    & " Inner Join new_Item_Type On (new_Item_Name.ItemType = new_Item_Type.ItemType) " _
                    & " WHERE Pro_Inv_Main.SuppID = " & cboSupplier.Tag & " AND new_Item_Name.ItemType = '" & cmbItemGroup.Text & "' AND Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                MainComm.Execute
            ElseIf cboSupplier.Text = "" And cmbItemGroup.Text <> "" Then
                
                MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
                    & " PIQty,Unit,UnitPrice,ItName,Total,Remark,PRNo,PICurr,Condition)" _
                    & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
                    & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
                    & " new_Item_Name.ItemName,Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition FROM Pro_Inv_Main " _
                    & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
                    & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) " _
                    & " Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
                    & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
                    & " Inner Join new_Item_Name On (Pro_Inv_Sub.itemID = new_Item_Name.itemID) " _
                    & " Inner Join new_Item_Type On (new_Item_Name.ItemType = new_Item_Type.ItemType) " _
                    & " WHERE new_Item_Name.ItemType = '" & cmbItemGroup.Text & "' AND Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0;"
                
                MainComm.Execute
            End If
        End If
    
        strRecord = "SELECT PINo FROM Temp_Import_Pro_Invoice GROUP BY PINo;"
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
    Set R = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub optItemGroup_Click()
    cboPINo.Enabled = False
    cmbItemGroup.Enabled = True
End Sub

Private Sub optItemGroup_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub optPI_Click()
    cboPINo.Enabled = True
    cmbItemGroup.Enabled = False
End Sub

Public Sub prcShowReport(str As String)
    Call generateReport
    
    If boEmptyText = True Then Exit Sub
    If optPI.Value = True Then
        
        If chkpad.Value = 1 Then
         If ChkOther.Value = 0 Then
         Set Report = Appl.OpenReport(findReportPath & "\rptPOpad.rpt")

         ElseIf ChkOther.Value = 1 Then
         Set Report = Appl.OpenReport(findReportPath & "\rptPOpadOC.rpt")

         End If
        Else
         If ChkOther.Value = 0 Then
         Set Report = Appl.OpenReport(findReportPath & "\rptPO.rpt")

         ElseIf ChkOther.Value = 1 Then
         Set Report = Appl.OpenReport(findReportPath & "\rptPOOC.rpt")

         End If
        End If
        
    Else
        If cboSupplier.Text = "" And cmbItemGroup.Text = "" Then
        Set Report = Appl.OpenReport(findReportPath & "\rptPOItemTypeAll.rpt")

        ElseIf cboSupplier.Text <> "" And cmbItemGroup.Text = "" Then
        Set Report = Appl.OpenReport(findReportPath & "\rptPOSupplierItem.rpt")

        ElseIf cboSupplier.Text = "" And cmbItemGroup.Text <> "" Then
        Set Report = Appl.OpenReport(findReportPath & "\rptPOItemWise.rpt")

        ElseIf cboSupplier.Text <> "" And cmbItemGroup.Text <> "" Then
        Set Report = Appl.OpenReport(findReportPath & "\rptPOItemSupplier.rpt")

        End If
    End If
    
    
    Report.DiscardSavedData
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = Report
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1

End Sub

Private Sub optPI_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
