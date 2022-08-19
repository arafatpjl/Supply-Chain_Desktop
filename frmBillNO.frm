VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBillNO 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8460
   ClientLeft      =   135
   ClientTop       =   15
   ClientWidth     =   10305
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
   Icon            =   "frmBillNO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "New Form"
      CausesValidation=   0   'False
      Height          =   525
      Left            =   390
      TabIndex        =   38
      Top             =   7774
      Width           =   1469
   End
   Begin VB.TextBox txtIDNO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   37
      Top             =   840
      Width           =   3975
   End
   Begin VB.CheckBox ChkShow 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   36
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox vmrid 
      Height          =   375
      Left            =   9600
      TabIndex        =   35
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CmbMRRNo 
      Height          =   315
      Left            =   3600
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   2820
      Width           =   3135
   End
   Begin VB.TextBox txtSL 
      Height          =   315
      Left            =   10080
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   525
      Left            =   5040
      TabIndex        =   13
      Top             =   7800
      Width           =   1185
   End
   Begin VB.ComboBox cmbIDNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmBillNO.frx":000C
      Left            =   1050
      List            =   "frmBillNO.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   795
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.CommandButton cmdAddHead 
      BackColor       =   &H00808080&
      Height          =   285
      Left            =   9840
      Picture         =   "frmBillNO.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Add New Name of Item"
      Top             =   2100
      Width           =   285
   End
   Begin VB.CommandButton cmdAddParty 
      BackColor       =   &H00808080&
      Height          =   285
      Left            =   3210
      Picture         =   "frmBillNO.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Add New Name of Item"
      Top             =   2820
      Width           =   285
   End
   Begin VB.ComboBox cmbPartyName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   180
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2820
      Width           =   3045
   End
   Begin VB.ComboBox cmbAccountsHead 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2100
      Width           =   3675
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   9240
      TabIndex        =   10
      Top             =   3480
      Width           =   885
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   330
      Left            =   5280
      TabIndex        =   17
      Top             =   780
      Width           =   765
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   525
      Left            =   3720
      TabIndex        =   16
      Top             =   7800
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   525
      Left            =   6360
      TabIndex        =   15
      Top             =   7800
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   9000
      TabIndex        =   11
      Top             =   7800
      Width           =   1185
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   525
      Left            =   7680
      TabIndex        =   12
      Top             =   7800
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3588
      Left            =   0
      TabIndex        =   14
      Top             =   3900
      Width           =   10270
      _ExtentX        =   18124
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColorFixed  =   -2147483648
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   5205
   End
   Begin VB.TextBox txtMRRDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   6840
      TabIndex        =   8
      Top             =   2820
      Width           =   2865
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   4170
      TabIndex        =   4
      Top             =   2100
      Width           =   1875
   End
   Begin VB.TextBox txtBillDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   2550
      TabIndex        =   3
      Top             =   2100
      Width           =   1575
   End
   Begin VB.TextBox txtBillNO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   2100
      Width           =   2325
   End
   Begin VB.TextBox txtIDDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   2475
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Details Credit"
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
      Height          =   510
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   5400
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   -360
      TabIndex        =   31
      Top             =   7560
      Width           =   12240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   10
      Left            =   -150
      TabIndex        =   30
      Top             =   7380
      Width           =   10305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   5235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Head Of Accounts"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   5
      Left            =   6090
      TabIndex        =   26
      Top             =   1770
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MRR Date"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   8
      Left            =   6840
      TabIndex        =   25
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MRR NO"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   7
      Left            =   3600
      TabIndex        =   24
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Party"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   6
      Left            =   60
      TabIndex        =   23
      Top             =   2520
      Width           =   3315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill/Invoice Amount"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   4
      Left            =   4200
      TabIndex        =   22
      Top             =   1770
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill/Invoice No"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   21
      Top             =   1770
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill/Invoice Date"
      ForeColor       =   &H00E26101&
      Height          =   285
      Index           =   2
      Left            =   2550
      TabIndex        =   20
      Top             =   1770
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      Caption         =   "ID Date  :"
      ForeColor       =   &H00E26101&
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   1230
      Width           =   825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      Caption         =   "ID NO  :"
      ForeColor       =   &H00E26101&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   825
   End
End
Attribute VB_Name = "frmBillNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytRowNo As Byte
Dim boEdit As Boolean
Dim intGSLNO As Integer
Dim intGMaxSLNO As Long
Dim rstHead As New Recordset
Dim rstParty As New Recordset

Private Sub ChkShow_Click()
If ChkShow.Value = 1 Then
cmbIDNO.Visible = True
txtIDNO.Visible = False
txtIDNO.Text = ""
Call prcAddIDNO
End If
If ChkShow.Value = 0 Then
cmbIDNO.Visible = False
txtIDNO.Visible = True
cmbIDNO.Text = ""
End If
End Sub

Private Sub cmbAccountsHead_Click()
On Error GoTo x
    cmbAccountsHead.Tag = 0
    If cmbAccountsHead.ListIndex >= 0 Then
        cmbAccountsHead.Tag = cmbAccountsHead.ItemData(cmbAccountsHead.ListIndex)
        cmbAccountsHead.ToolTipText = cmbAccountsHead.Text
    End If
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbAccountsHead_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbIDNo_Click()
On Error GoTo x
    Call prcBlank(2)
    txtIDDate.Text = ""
    flexDetails.Rows = 1
    
    cmdAdd.Caption = "&Add"
'    intGMaxSLNO = findmxslnobill(cmbIDNO.Text)
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmbIDNo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then cmdFind.SetFocus
End Sub


Private Sub CmbMRRNo_Click()
'If Len(CmbMRRNo) <> 0 Then
'Call prcAddSuppMrrno(vmrid, CmbMRRNo)
'Else
'MsgBox "Select Mrr No", vbExclamation, "Empty MRR No"
'Exit Sub
'
'End If
'
'Set R = New ADODB.Recordset
'    If Len(cmbPartyName) <> 0 Then
'    strRecord = " SELECT MRR_Challan_Main.compID, Company_Information.compName, Company_Information.compAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " FROM ((Company_Information INNER JOIN MRR_Challan_Main ON Company_Information.compID = MRR_Challan_Main.compID) INNER JOIN new_Supplier ON MRR_Challan_Main.suppID = new_Supplier.supplierID) INNER JOIN (new_Item_Type INNER JOIN ((MRR_Challan_Sub INNER JOIN new_Item_Name ON MRR_Challan_Sub.itemID = new_Item_Name.itemID) INNER JOIN New_PI ON MRR_Challan_Sub.PIID = New_PI.PIID) ON new_Item_Type.itemType = new_Item_Name.itemType) ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID " & _
'        " WHERE MRR_Challan_SUB.DELETEROW = 0 GROUP BY MRR_Challan_Main.compID, Company_Information.compName, Company_Information.compAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " HAVING (((MRR_Challan_Main.compID)=" & bytCompID & ")AND ((MRR_Challan_Main.MRRID)='" & vmrid.Text & "'));"
'
'    R.Open strRecord, MainConn, adOpenStatic
'    Else
'    MsgBox "Select supplier name", vbCritical, "Empty supplier name"
'    Exit Sub
'    End If
'
'    If R.EOF = False And R.BOF = False Then
'
'        txtMRRDate.Text = Format(R![mrrdate], cnstDtFrmtI)
'        txtMrAmount.Text = Format(R![MRRValue], "##,##0.00")
'        txtTotal.Text = txtMrAmount.Text
'
'    Else
'    MsgBox "Record not found", vbExclamation
'    Exit Sub
'    End If
End Sub

Private Sub CmbMRRNo_GotFocus()
'CmbMRRNo.Clear
'Call prcAddSuppMrrID(CmbMRRNo, cmbPartyName.Text)
End Sub

Private Sub cmbMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbPartyName_Click()

On Error GoTo x
    cmbPartyName.Tag = 0
    If cmbPartyName.ListIndex >= 0 Then
        cmbPartyName.Tag = cmbPartyName.ItemData(cmbPartyName.ListIndex)
        cmbPartyName.ToolTipText = cmbPartyName.Text
    End If
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title

End Sub

Private Sub cmbPartyName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo x
    Dim strItem As String
    'If fncBlank(2) = True Then Exit Sub
    
    If Len(txtBillNo) = 0 Then
         MsgBox "Entry bill no", vbExclamation, "Empty bill no"
       Exit Sub
End If


If cmdSave.Caption = "&Update" Then

     If cmbIDNO.Text = "" Then
    
        MsgBox "Select ID NO", vbExclamation, "Empty ID NO"
        
     Exit Sub
     
     Else
     
     End If
 
End If



If Len(cmbAccountsHead) = 0 Then
MsgBox "Select Head of Accounts", vbExclamation, "Empty accounts head"
Exit Sub
End If

If Len(cmbPartyName) = 0 Then
MsgBox "Select party name", vbExclamation, "Empty party name"
Exit Sub
End If

If Len(cmbMRRNo) = 0 Then
MsgBox "Select MRR no", vbExclamation, "Empty MRR no"
Exit Sub
End If
    
    
    With flexDetails
        If .Rows > 1 Then
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0: strItem = .Text
                
                If strItem = txtBillNo.Text And inti <> bytRowNo Then
                    MsgBox "Adding Duplicate Item", vbInformation, cnstMsgInfo
                    txtBillNo.SetFocus
                    Exit Sub
                End If
            Next inti
        End If
    End With
    Dim amt As Currency
    Dim BillDate As Date
    Dim mrrdate As Date
    Dim iddate As Date
    If Len(Trim(txtAmount.Text)) = 0 Then
        amt = "0"
    Else
        amt = CCur(txtAmount.Text)
    End If
    
    If Len(Trim(txtBillDate.Text)) = 0 Then
'        billdate = ""   'Format("0", cnstDtFrmtI)
    Else
        BillDate = txtBillDate.Text
    End If
    
    If Len(Trim(txtMRRDate.Text)) = 0 Then
        'mrrdate = ""  'Format("0", cnstDtFrmtI)
    Else
        mrrdate = txtMRRDate.Text
    End If
    If Len(Trim(txtIDDate.Text)) = 0 Then
'        iddate = ""   'Format("0", cnstDtFrmtI)
    Else
        iddate = txtIDDate.Text
    End If
    
 
    
    With flexDetails
         'If boEdit = False Then
         If cmdAdd.Caption = "&Add" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        Else
            .Row = bytRowNo
        End If
      
        .Col = 0:      .Text = txtBillNo.Text
        .Col = 1:      .Text = txtBillDate.Text
        .Col = 2:      .Text = txtAmount.Text
        .Col = 3:      .Text = cmbAccountsHead.Text
        .Col = 4:      .Text = cmbPartyName.Text
        .Col = 5:      .Text = cmbMRRNo.Text
        .Col = 6:      .Text = txtMRRDate.Text
        .Col = 7:      .Text = cmbAccountsHead.Tag
        .Col = 8:      .Text = cmbPartyName.Tag
        .Col = 9:      .Text = txtRemarks.Text
        
    End With
    
    cmdAdd.Caption = "&Add"
    
    prcBlank (2)
    txtBillNo.SetFocus
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAddHead_Click()
On Error GoTo x
    flagMark = "INSERT"
    strMark = "HEAD"
    strValue = ""

    frmNewAccountsHead.Show 1

    Call prcAddHead(cmbAccountsHead)
    If Len(strValue) > 0 Then cmbAccountsHead.Text = strValue
    cmbAccountsHead.SetFocus
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo x
    If cmdAddNew.Caption = "&Edit" Then
        cmdAddNew.Caption = "Add &New"
        cmdSave.Caption = "&Update"
        cmdDelete.Enabled = True
        cmbIDNO.Visible = True
        txtIDNO.Visible = False
        cmdFind.Enabled = True
        
        cmbIDNO.SetFocus
        
        Call prcAddIDNO

    Else
        cmbIDNO.Tag = ""
        cmdAddNew.Caption = "&Edit"
        cmdSave.Caption = "&Save"
        cmdDelete.Enabled = False
        cmbIDNO.Visible = False
        txtIDNO.Visible = True
        cmdFind.Enabled = False
        
        txtIDNO.SetFocus
    End If
    Call prcBlank(1)
    flexDetails.Rows = 1
    
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAddParty_Click()
    strFormName = "Party Name For BillDetails"
    flagMark = "INSERT"
    strMark = "PARTY"
    strValue = ""
    'frmNewParty2.Caption = "New Supplier of Bill Details"
    'frmNewParty2.cmdClose.Visible = False
    frmNewParty2.Show 1
    
    Call prcAddParty(cmbPartyName)
    If Len(strValue) > 0 Then cmbPartyName.Text = strValue
    strFormName = ""
    cmbPartyName.SetFocus
End Sub

Private Sub cmdDelete_Click()
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
MainConn.BeginTrans
On Error GoTo x
    If Trim(cmbIDNO) = "" Then
        MsgBox "You should provide ID NO", vbCritical, App.Title
        cmbIDNO.SetFocus
        Exit Sub
    End If
    If MsgBox("Are you confirm to delete ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Mysql = "UPDATE BillMain SET DELETEROW=1,USERID='" & strUser & "',PCNAME='" & getComName & "',ENTRYDATE='" & Date & "',ENTRYTIME='" & Time & "' Where IDNO = '" & Trim(cmbIDNO.Text) & "' and CompID = " & bytCompID & " and Status = 'C' and Deleterow=0;"
'        Mysql = "Delete From BillMain Where IDNO = '" & Trim(cmbIDNo.Text) & "' and CompID = " & bytCompID & " and Status = 'C';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
         Mysql = " Update BillDetails set  DELETEROW=1,PCNAME='" & getComName & "',USERID='" & strUser & "',ENTRYDATE='" & Date & "',ENTRYTIME='" & Time & "' from BillDetails Inner Join BillMain On BillMain.BillID = BillDetails.BillID Where BillDetails.BillID = " & Val(cmbIDNO.Tag) & " And BillMain.CompID = " & bytCompID & " And BillMain.Status = 'C' and BillDetails.deleterow=0;"
         MainComm.CommandText = Mysql
         MainComm.Execute
    MainConn.CommitTrans
    
        Call prcBlank(1)
        MsgBox "Deletion complete successfully.", vbInformation, App.Title
        Call prcAddIDNO
    End If
    
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
   If MsgBox("Do you want to exit from this window.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    Set frmBillNO = Nothing

    Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo x
    If cmdAdd.Caption = "&Replace" Then Exit Sub
    
    Dim tmprst As New Recordset

    If Trim(cmbIDNO.Text) = "" Then
        MsgBox "You should provide ID NO. ", vbCritical, App.Title
        cmbIDNO.SetFocus
        Exit Sub
    End If

    Mysql = " SELECT BillDetails.BillID, BillMain.IDNO, BillMain.IDDate, BillDetails.BillNO, BillDetails.BillDate, BillDetails.BillAmount, BillDetails.HeadID, AccountsHead.HeadName, BillDetails.PartyID, Party.PartyName, BillDetails.MRRNO, BillDetails.MRRDate, BillDetails.Remarks "
    Mysql = Mysql + " FROM Party INNER JOIN (BillMain INNER JOIN (AccountsHead INNER JOIN BillDetails ON AccountsHead.HeadID = BillDetails.HeadID) ON BillMain.BillID = BillDetails.BillID) ON Party.PartyID = BillDetails.PartyID "
    Mysql = Mysql + " Where BillMain.Status = 'C' and BillMain.IDNO = '" & Trim(cmbIDNO.Text) & "' and BillMain.CompID = " & bytCompID & " AND BillDetails.Deleterow=0  Order By RowNo;"
    tmprst.Open Mysql, MainConn, adOpenStatic

    If tmprst.EOF = False And tmprst.BOF = False Then
        tmprst.MoveFirst
        cmbIDNO.Tag = tmprst("BillID")
        If tmprst("IDDate") = "12:00:00 AM" Then
            txtIDDate.Text = ""
        Else
            txtIDDate.Text = tmprst("IDDate")
        End If
        
        With flexDetails
            .Rows = 1
            Do Until tmprst.EOF
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .Col = 0:   .Text = tmprst("BillNO") & ""
                .Col = 1
                If tmprst("BillDate") = "12:00:00 AM" Then
                    .Text = ""
                Else
                    .Text = tmprst("BillDate") & ""
                End If
                .Col = 2:   .Text = tmprst("BillAmount") & ""
                .Col = 3:   .Text = tmprst("HeadName") & ""
                .Col = 4:   .Text = tmprst("PartyName") & ""
                .Col = 5:   .Text = tmprst("MRRNO") & ""
                .Col = 6
                If tmprst("MRRDate") = "12:00:00 AM" Then
                    .Text = ""
                Else
                    .Text = tmprst("MRRDate") & ""
                End If
                .Col = 7: .Text = tmprst("HeadID")
                .Col = 8: .Text = tmprst("PartyID")
                .Col = 9: .Text = tmprst("Remarks") & ""
                '.Col = 10: .Text = IIf(Not IsNull(tmprst("SLNO")), tmprst("SLNO"), "")
        tmprst.MoveNext
        Loop
        End With
    End If
    
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdPreview_Click()
On Error GoTo x
    Dim strIDNo As String
'    Call prcConReport(CrystalReport1)
    If txtIDNO.Visible = True Then
        strIDNo = Trim(txtIDNO.Text)
        If Len(Trim(txtIDNO.Text)) = 0 Then
            MsgBox "You should provide IDNO.", vbCritical, App.Title
            txtIDNO.SetFocus
            Exit Sub
        End If
        
    Else
    
        strIDNo = Trim(cmbIDNO.Text)
        If Len(Trim(cmbIDNO.Text)) = 0 Then
            MsgBox "You should provide IDNO.", vbCritical, App.Title
            cmbIDNO.SetFocus
            Exit Sub
        End If
    End If

    Set R1 = New ADODB.Recordset
    Set REPORT = Appl.OpenReport(findReportPath & "Bill Details\rptDetailsIDNoEntry.rpt")
    strRecord = "SELECT * FROM vieBillDetails WHERE IDNO='" & strIDNo & "' And Status='C' And CompID = " & bytCompID & " Order By RowNO;"
    
    R1.Open strRecord, MainConn, adOpenStatic
    
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R1
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = REPORT
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1
    R1.Close

Exit Sub
x:
Screen.MousePointer = 0
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdSave_Click()
On Error GoTo x
    Dim strBillNo As String, strDate As String, dblAmount As Double
    Dim strMRRNO As String, strMRRDate As String, intHead As Integer
    Dim intPartyID As Integer, strRemarks As String
    Dim intACCID As Integer, intACC As String
    Dim intACNo As Integer
    
    If flexDetails.Row < 1 Then
     
        MsgBox "No Data found !", vbInformation, cnstMsgInfo
        cmbPartyName.SetFocus
        Exit Sub
    End If
    
    Set MainComm = New ADODB.Command
    If cmdSave.Caption = "&Save" Then
        If MsgBox("Do you want to Save?", vbInformation + vbYesNo, App.Title) = vbNo Then Exit Sub
    ElseIf cmdSave.Caption = "&Update" Then
        If MsgBox("Do you want to Update?", vbInformation + vbYesNo, App.Title) = vbNo Then Exit Sub
    End If
    If fncBlank(1) = True Then Exit Sub

    If cmdSave.Caption = "&Save" Then
        MainConn.BeginTrans
            Dim tmprst As New Recordset, tmprst2 As New Recordset
            Dim intmaxid As Integer
            
'            Mysql = " Select BillID From BillMain Where IDNO = '" & Trim(txtIDNO.Text) & "' and ComPID = " & bytCompID & " and Status = 'C';"
'             tmprst2.Open Mysql, MainConn, adOpenStatic
'
'            If tmprst2.EOF = False And tmprst2.EOF = False Then
'                intmaxid = tmprst2("BillID")
'            Else
                Mysql = "Select max(BillID) as maxID from BillMain;"
                 tmprst.Open Mysql, MainConn, adOpenStatic
                 
'                If tmprst.EOF = False And tmprst.BOF = False Then
                    If IsNull(tmprst("maxID")) = False Then
                        intmaxid = tmprst("maxid") + 1
                    Else
                        intmaxid = 1
                    End If
'                Else
'                    intmaxid = 1
'                End If
                Set tmprst = Nothing
            
                Dim iddate As Date
                If Len(Trim(txtIDDate.Text)) = 0 Then
                    iddate = ""  'Format("0", cnstDtFrmtI)
                Else
                    iddate = Format(txtIDDate.Text, cnstDtFrmtI)
                End If
                
                intYear = Year(txtIDDate.Text)
                Set RS = New ADODB.Recordset
                    strRecord = "SELECT MAX(ACCID) as ACCID FROM New_ACCID Where ComID = " & bytCompID & " And ACC_Year = " & intYear & " AND TransType='Credit';"
                    MainComm.CommandText = strRecord
                RS.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(RS![ACCID]) Then
                        intACCID = RS.Fields(0) + 1
                    Else
                        intACCID = 1
                    End If
                Set RS = Nothing
            intACC = fncMRR(intACCID)
            txtIDNO.Text = "ACC" & bytCompID & intACC & "-" & intYear
            
            Set R = New ADODB.Recordset
                    strRecord = "SELECT MAX(ACCID) as ACCID FROM New_ACCID;"
                    MainComm.CommandText = strRecord
                R.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(R![ACCID]) Then
                        intACNo = R.Fields(0) + 1
                    Else
                        intACNo = 1
                    End If
                Set R = Nothing
            
                Mysql = "INSERT INTO New_ACCID(ACCid,comID, TransType, ACC_No, ACC_year,ACCNO,EntryDate,EntryTime)" _
                  & " VALUES (" & intACNo & "," & bytCompID & ",'Credit','" & intACC & "','" & intYear & "','" & txtIDNO & "','" & Date & "','" & Time & "');"
'            MainComm.Execute
                MainConn.Execute (Mysql)
            
                Mysql = " Insert Into BillMain (BillID, IDNO, IDDate,CompID, Status,Deleterow,Userid,PCName,EntryDate,EntryTime)"
                Mysql = Mysql + " Values (" & intmaxid & ", '" & Trim(txtIDNO.Text) & "', '" & iddate & "', " & bytCompID & ", 'C', 0,'" & strUser & "', '" & getComName & "', '" & Date & "', '" & Time & "');"
                MainConn.Execute (Mysql)
'            End If
            With flexDetails
                For inti = 1 To .Rows - 1
                    .Row = inti
                    .Col = 0: strBillNo = .Text
                    .Col = 1:
                    If Len(Trim(.Text)) = 0 Then
                        strDate = ""  'Format("0", cnstDtFrmtI)
                    Else
                        strDate = .Text
                    End If
                    .Col = 2
                    If .Text = "" Then
                        dblAmount = 0
                    Else
                        dblAmount = CDbl(.Text)
                    End If
'                   .Col = 3:   Trim (.Text)
'                   .Col = 4:   Trim (.Text)
                    .Col = 5:   strMRRNO = Trim(.Text)
                    .Col = 6:
                    If Len(Trim(.Text)) = 0 Then
                        StrMRRDT = ""   'Format("0", cnstDtFrmtI)
                    Else
                        StrMRRDT = .Text
                    End If
                    .Col = 7:   intHead = CInt(.Text)
                    .Col = 8:   intPartyID = .Text
                    .Col = 9:   strRemarks = Trim(.Text)
                    '.Col = 10:   intSLNO = .Text
                    
                    Mysql = " Insert  Into BillDetails (RowNo, BillID, BillNo, BillDate, BillAmount, HeadID, PartyID, MRRNO, MRRDate, Remarks,pcname,userid,entrydate,entrytime)"
                    Mysql = Mysql + " values (" & inti & ", " & intmaxid & ", '" & strBillNo & "', '" & strDate & "', " & dblAmount & ", " & intHead & ", " & intPartyID & ", '" & strMRRNO & "', '" & StrMRRDT & "', '" & strRemarks & "','" & getComName & "', '" & strUser & "', '" & Date & "', '" & Time & "');"
                    MainConn.Execute (Mysql)
                Next inti
            End With
        MainConn.CommitTrans
    ElseIf cmdSave.Caption = "&Update" Then
        MainConn.BeginTrans
            Mysql = " Update BillMain Set IDNO='" & Trim(cmbIDNO.Text) & "' , IDDate = '" & txtIDDate.Text & "'"
            Mysql = Mysql + " Where BillID = " & Val(cmbIDNO.Tag) & " and CompID = " & bytCompID & " and Status = 'C' and Deleterow=0;"
            MainConn.Execute (Mysql)
            
'            Mysql = " Delete BillDetails from BillDetails Inner Join BillMain On BillMain.BillID = BillDetails.BillID Where BillDetails.BillID = " & Val(cmbIDNO.Tag) & " And BillMain.CompID = " & bytCompID & " And BillMain.Status = 'C' ;"
            Mysql = " Update BillDetails set  Deleterow=2,pcname='" & getComName & "',Userid='" & strUser & "',Entrydate='" & Date & "',Entrytime='" & Time & "' from BillDetails Inner Join BillMain On BillMain.BillID = BillDetails.BillID Where BillDetails.BillID = " & Val(cmbIDNO.Tag) & " And BillMain.CompID = " & bytCompID & " And BillMain.Status = 'C' and BillDetails.Deleterow=0;"
            MainConn.Execute (Mysql)
            
            With flexDetails
                For inti = 1 To .Rows - 1
                    .Row = inti
                    .Col = 0:   strBillNo = .Text
                    .Col = 1
                    If Len(Trim(.Text)) = 0 Then
                        strDate = ""  'Format("0", cnstDtFrmtI)
                    Else
                        strDate = .Text
                    End If
                    .Col = 2:   dblAmount = CDbl(.Text)
'                   .Col = 3:   Trim (.Text)
'                   .Col = 4:   Trim (.Text)
                    .Col = 5:   strMRRNO = Trim(.Text)
                    .Col = 6
                    If Len(Trim(.Text)) = 0 Then
                        StrMRRDT = ""  'Format("0", cnstDtFrmtI)
                    Else
                        StrMRRDT = .Text
                    End If
                    .Col = 7:   intHead = CInt(.Text)
                    .Col = 8:   intPartyID = .Text
                    .Col = 9:   strRemarks = Trim(.Text)
'                    .Col = 10:  intSLNO = .Text
'                    Mysql = " Insert  Into BillDetails (RowNo, BillID, BillNo, BillDate, BillAmount, HeadID, PartyID, MRRNO, MRRDate, Remarks)"
'                    Mysql = Mysql + " values (" & intI & ", " & Val(cmbIDNo.Tag) & ", '" & strBillNo & "', '" & strDate & "', " & dblAmount & ", " & intHead & ", " & intPartyID & ", '" & strMRRNO & "', '" & StrMRRDT & "', '" & strRemarks & "');"
'
'                    Mysql = " Insert Into BillDetails (BillID, IDNO, IDDate,CompID, Status,Userid,PCName,EntryDate,EntryTime)"
'                    Mysql = Mysql + " Values (" & intI & ", '" & Trim(cmbIDNo.Tag) & "', '" & iddate & "', " & bytCompID & ", 'C', " & bytUserID & ", '" & getComName & "', '" & Date & "', '" & Time & "');"
                    
                    Mysql = " Insert  Into BillDetails (RowNo, BillID, BillNo, BillDate, BillAmount, HeadID, PartyID, MRRNO, MRRDate, Remarks,pcname,userid,Deleterow,entrydate,entrytime)"
                    Mysql = Mysql + " values (" & inti & ",'" & Trim(cmbIDNO.Tag) & "', '" & strBillNo & "', '" & strDate & "', " & dblAmount & ", " & intHead & ", " & intPartyID & ", '" & strMRRNO & "', '" & StrMRRDT & "', '" & strRemarks & "','" & getComName & "', '" & strUser & "', 0,'" & Date & "', '" & Time & "');"
                    MainConn.Execute (Mysql)
                   
                  
                Next inti
            End With
        MainConn.CommitTrans
    End If

    Call prcBlank(1)
'    Call prcAddIDNO

    If cmdSave.Caption = "&Save" Then
        MsgBox "Data saved successfully.", vbInformation, App.Title
    ElseIf cmdSave.Caption = "&Update" Then
        MsgBox "Data updated successfully.", vbInformation, App.Title
    End If
    
'    If txtIDNO.Visible = True Then
'        txtIDNO.SetFocus
'    Else
'        cmbIDNo.SetFocus
'    End If
Exit Sub
x:
MainConn.RollbackTrans
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub Command1_Click()
  Unload Me
  frmBillEntryAcc.Show 1
End Sub

Private Sub flexDetails_DblClick()
On Error Resume Next
     flexDetails_KeyDown 13, 0
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
            With flexDetails
            If .RowSel = 0 Then Exit Sub
            .Col = 0:    txtBillNo.Text = .Text
            .Col = 1:    txtBillDate.Text = .Text
            .Col = 2:    txtAmount.Text = .Text
            .Col = 3:    cmbAccountsHead.Text = .Text
            .Col = 4:    cmbPartyName.Text = .Text
            .Col = 5:    cmbMRRNo.Text = .Text
            .Col = 6:    txtMRRDate.Text = .Text
            .Col = 7:    cmbAccountsHead.Tag = .Text
            .Col = 8:    cmbPartyName.Tag = .Text
            '.Col = 8:    cmbPartyName.Text = .Text + " " + .TextMatrix(.Row, 4)
            .Col = 9:    txtRemarks.Text = .Text
            '.Col = 10:    txtSL = .Text
            bytRowNo = .Row
        End With
'          boEdit = True
        cmdAdd.Caption = "&Replace"
        txtBillNo.SetFocus
    End If
    If KeyCode = vbKeyDelete Then
        With flexDetails
         
            If .RowSel = 0 Then
                Beep
            Else
                If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
                If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
            End If
            
        End With
        cmdAdd.Caption = "&Add"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo x
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    strFormName = "Party Name For BillDetails"
    Call prcGridSetting
    Call prcAddHead(cmbAccountsHead)
    Call prcAddParty(cmbPartyName)
'    Call addSuppId(txtIDNO, 2)
    
    cmdFind.Enabled = False
    'cmdDelete.Enabled = False
    
    Set R = New ADODB.Recordset
    strRecord = "Select DelStatus,UpStatus from InventoryPJL.dbo.Sys_User_Name where UserName='" & strUser & "'"
    R.Open strRecord, MainConn, adOpenStatic
    If R!DelStatus = 1 Then
       cmdDelete.Enabled = True
    End If
    If R!UpStatus = 1 Then
        cmdAddNew.Enabled = True
    End If
    
    
    strFormName = ""
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Public Sub prcAddParty(ByRef cmbPartyName As ComboBox)

Set R = New ADODB.Recordset
        Mysql = "Select PartyID, PartyName From Party Order By PartyName"
        R.Open Mysql, MainConn, adOpenStatic
        
        cmbPartyName.Clear
        If R.EOF = False And R.BOF = False Then
            R.MoveFirst
            Do While Not R.EOF
                cmbPartyName.AddItem R("PartyName")
                cmbPartyName.ItemData(cmbPartyName.NewIndex) = R("PartyID")
            R.MoveNext
            Loop
        End If
    Set R = Nothing
    
End Sub

Public Sub prcGridSetting()
    With flexDetails
        .Rows = 1
        .Cols = 10
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignCenterCenter
        '.ColAlignment(10) = flexAlignCenterCenter
        
        .Col = 0:   .Text = "Bill/Inv. No"
        .Col = 1:   .Text = "Bill Date"
        .Col = 2:   .Text = "Bill Amount"
        .Col = 3:   .Text = "A/C Head"
        .Col = 4:   .Text = "Party Name"
        .Col = 5:   .Text = "MRR NO"
        .Col = 6:   .Text = "MRR Date"
        .Col = 7:   .Text = "A/C HeadID"
        .Col = 8:   .Text = "Party ID"
        .Col = 9:   .Text = "Remarks"
        '.Col = 10:   .Text = "Slno"
        
        
        .ColWidth(0) = 1200:     .ColWidth(1) = 1200
        .ColWidth(2) = 1200:     .ColWidth(3) = 2200
        .ColWidth(4) = 2200:     .ColWidth(5) = 1200
        .ColWidth(6) = 1200:     .ColWidth(7) = 0
        .ColWidth(8) = 0:        .ColWidth(9) = 2000
         '.ColWidth(10) = 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' 'If MsgBox("Do you want to exit from this window.", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
   ' Set frmBillNO = Nothing
End Sub

''Public Sub prcCon()
''    With MainConn
''        .Provider = "Microsoft.Jet.OLEDB.4.0"
''        '.ConnectionString = "Data Source=" & App.Path & "\Database.mdb;Persist Security Info=False"
''        .ConnectionString = "Data Source=S:\BillDetails\Database.mdb;Persist Security Info=False"
''        .Open
''    End With
''    MainComm.ActiveConnection = MainConn
''End Sub

Private Sub txtAmount_GotFocus()
On Error Resume Next
    Call txtGotFocus(txtAmount)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtAmount_LostFocus()

If Len(txtAmount.Text) > 0 Then txtAmount.Text = Format(txtAmount, "##,##0.00")
'txtAmount.Text = FormatNumber(txtAmount.Text, 0)
End Sub

Private Sub txtBillDate_GotFocus()
On Error Resume Next
    Call txtGotFocus(txtBillDate)
End Sub

Private Sub txtBillDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBillDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtBillDate_LostFocus()
On Error Resume Next
    Call FormatDate(txtBillDate)
End Sub

Private Sub txtBillNo_GotFocus()
On Error Resume Next
    Call txtGotFocus(txtBillNo)
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDiscount_Change()
On Error GoTo err
If Len(txtDiscount) >= 0 And Val(txtDiscount) >= 0 Then
txtTotal.Text = Format(txtMrAmount - txtDiscount, "##,##0.00")
Else
MsgBox "Entry discount properly", vbCritical, "Discount"
txtDiscount.Text = "0.00"
Exit Sub
End If

Exit Sub
err:
MsgBox "Entry discount properly", vbExclamation, "Discount"
'txtDiscount.Text = "0.00"
txtTotal.Text = Format(txtMrAmount, "##,##0.00")
End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDiscount_LostFocus()
If Val(txtDiscount) < 0 Then
MsgBox "Entry discount properly", vbExclamation, "Discount"
txtDiscount.Text = "0.00"
End If
End Sub
Private Sub txtIDDate_GotFocus()
On Error Resume Next
    Call txtGotFocus(txtIDDate)
End Sub

Private Sub txtIDDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtIDDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtIDDate_LostFocus()
On Error Resume Next
    Call FormatDate(txtIDDate)
End Sub

Private Sub txtIDNO_GotFocus()
On Error Resume Next
    Call txtGotFocus(txtIDNO)
End Sub

Private Sub txtIDNO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMRRDate_GotFocus()
On Error Resume Next
    Call txtGotFocus(txtMRRDate)
End Sub

Private Sub txtMRRDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMRRDate_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtMRRDate_LostFocus()
On Error Resume Next
    Call FormatDate(txtMRRDate)
End Sub

Private Sub txtMRRNO_GotFocus()
'On Error Resume Next
'    Call txtGotFocus(txtMRRNo)
End Sub

Private Sub txtMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Public Sub prcBlank(intval As Integer)
    If intval = 1 Then
        If txtIDNO.Visible = True Then
'            txtIDNO.Text = ""
        Else
            cmbIDNO.Text = ""
        End If
        txtIDDate.Text = ""
        flexDetails.Rows = 1
    End If
    txtBillNo.Text = ""
    txtBillDate.Text = ""
    txtAmount.Text = ""
'    txtMrAmount.Text = ""
'    txtDiscount.Text = "0.00"
'    txtTotal.Text = ""
    'cmbAccountsHead.ListIndex = -1
    'cmbPartyName.ListIndex = -1
'    txtMRRNo.Text = ""
'    CmbMRRNo.ListIndex = -1
    txtMRRDate.Text = ""
    txtRemarks.Text = ""
End Sub

Public Function fncBlank(intval As Integer) As Boolean
fncBlank = True
    If intval = 1 Then
'        If Trim(txtIDNO.Text) = "" And cmdSave.Caption = "&Save" Then
'            MsgBox "You should provide ID NO", vbCritical, App.Title
'            txtIDNO.SetFocus
'            Exit Function
'        End If
        If Trim(txtIDDate.Text) = "" Then
            MsgBox "You should provide ID Date", vbCritical, App.Title
            txtIDDate.SetFocus
            Exit Function
        End If
        
        If flexDetails.Rows = 1 Then MsgBox "No data found to save.", vbCritical, App.Title
    Else
        If Trim(txtBillNo.Text) = "" Then
            MsgBox "You should provide Bill/Invoice NO", vbCritical, App.Title
            txtBillNo.SetFocus
            Exit Function
        End If
        If Trim(txtBillDate.Text) = "" Then
            MsgBox "You should provide Bill/Invoice Date", vbCritical, App.Title
            txtBillDate.SetFocus
            Exit Function
        End If
        If Trim(txtAmount.Text) = "" Then
            MsgBox "You should provide Bill/Invoice Amount", vbCritical, App.Title
            txtAmount.SetFocus
            Exit Function
        End If
        If Trim(cmbAccountsHead.Text) = "" Then
            MsgBox "You should provide accounts head name.", vbCritical, App.Title
            cmbAccountsHead.SetFocus
            Exit Function
        End If
        If Trim(cmbPartyName.Text) = "" Then
            MsgBox "You should provide party name.", vbCritical, App.Title
            cmbPartyName.SetFocus
            Exit Function
        End If
        If Trim(txtMRRNo.Text) = "" Then
            MsgBox "You should provide MRR NO.", vbCritical, App.Title
            txtMRRNo.SetFocus
            Exit Function
        End If
        If Trim(txtMRRDate.Text) = "" Then
            MsgBox "You should provide MRR Date.", vbCritical, App.Title
            txtMRRDate.SetFocus
            Exit Function
        End If
    End If
fncBlank = False
End Function

Public Sub prcAddIDNO()
    Dim tmprst As New Recordset
        Mysql = "Select Distinct IDNO From BillMain Where CompID = " & bytCompID & " and Status = 'C' and Deleterow=0;"
        tmprst.Open Mysql, MainConn, adOpenStatic
        If tmprst.EOF = False And tmprst.BOF = False Then
            cmbIDNO.Clear
            tmprst.MoveFirst
            Do While Not tmprst.EOF
                cmbIDNO.AddItem tmprst("IDNO")
                tmprst.MoveNext
            Loop
        Else
            cmbIDNO.Clear
        End If
    Set tmprst = Nothing
End Sub
