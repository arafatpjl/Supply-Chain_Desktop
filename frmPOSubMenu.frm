VERSION 5.00
Begin VB.Form frmPOSubMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H00C0C000&
      Caption         =   "&Show"
      Enabled         =   0   'False
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
      Left            =   4860
      TabIndex        =   21
      ToolTipText     =   "Show Record"
      Top             =   2340
      Width           =   615
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
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
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "New/Edit Record"
      Top             =   4320
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
      Caption         =   "&Delete"
      Enabled         =   0   'False
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
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Delete Record"
      Top             =   4320
      Width           =   1185
   End
   Begin VB.ComboBox cmbSubject 
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
      ItemData        =   "frmPOSubMenu.frx":0000
      Left            =   1860
      List            =   "frmPOSubMenu.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2340
      Width           =   2910
   End
   Begin VB.TextBox txtAmount 
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
      Height          =   315
      Left            =   1860
      TabIndex        =   3
      Top             =   2820
      Width           =   2835
   End
   Begin VB.ComboBox cboSupplier 
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
      Left            =   1845
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1395
      Width           =   4350
   End
   Begin VB.ComboBox cboPINo 
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
      Left            =   1845
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1860
      Width           =   4350
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      CausesValidation=   0   'False
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "flexAlignCenterCenter"
      Top             =   4320
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   4920
      TabIndex        =   5
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
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
      Left            =   1620
      TabIndex        =   20
      Top             =   2820
      Width           =   180
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
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
      Left            =   1620
      TabIndex        =   19
      Top             =   2340
      Width           =   180
   End
   Begin VB.Label lblItem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subject Name"
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
      Left            =   330
      TabIndex        =   18
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
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
      Left            =   1620
      TabIndex        =   17
      Top             =   1380
      Width           =   180
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount"
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
      Left            =   330
      TabIndex        =   16
      Top             =   2820
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
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
      Left            =   1620
      TabIndex        =   15
      Top             =   1860
      Width           =   180
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P.O.  No."
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
      Left            =   330
      TabIndex        =   14
      Top             =   1860
      Width           =   930
   End
   Begin VB.Label Label38 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supplier Name"
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
      Left            =   330
      TabIndex        =   13
      Top             =   1380
      Width           =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO :: DEDUCTION...."
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
      Left            =   165
      TabIndex        =   11
      Top             =   240
      Width           =   2835
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
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   6555
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   4125
      Width           =   6675
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -120
      TabIndex        =   8
      Top             =   4080
      Width           =   6675
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6600
   End
End
Attribute VB_Name = "frmPOSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IncrSL, SLNo As Integer

Private Sub cboPINo_Click()
'   cmbMRRID.Clear
    If Len(cboPINo) = 0 Then Exit Sub
 '   addSuppPIMRRNo cboPINo, cmbMRRID, cboSupplier.Text
End Sub

Private Sub cboPINo_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cboPINo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cboSupplier_Click()
  'cmbMRRID.Clear
    If Len(cboSupplier) = 0 Then Exit Sub
    addSuppPINo cboPINo, cboSupplier.Text
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmbMrrID_Click()
If Len(cmbMRRID) = 0 Then Exit Sub
'addSubject cboSupplier.Text, cboPINo.Text, cmbMRRID.Text
Set R = New ADODB.Recordset
End Sub

Private Sub cmbMrrID_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbMRRID_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmbSubject_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbSubject_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdDelete_Click()

   If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans

    MainComm.CommandText = "UPDATE PO_Sub_Menu SET DeleteRow = 1,UserId=" & bytUserID & ",PCName='" & getComName & "',EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' Where SupplierID =" & findSupplierID(cboSupplier.Text) & " AND POID=" & findPIID(cboPINo.Text) & " And DeleteRow = 0;"
    MainComm.Execute
    
    MainConn.CommitTrans
ErrorHandler:
    If err.Number <> 0 Then
        If err.Number = -2147467259 Then
            MsgBox "This record cannot be deleted", vbCritical, cnstMsgErDB
        Else
            MsgBox err.Description, vbCritical, cnstMsgErDB
        End If
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        Exit Sub
    End If

    Screen.MousePointer = vbDefault
    
 cmdEdit.Caption = "&Edit"
 cmdDelete.Enabled = False
 cmdSave.Enabled = True
 txtAmount.Text = ""
End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdEdit_Click()
If cmdEdit.Caption = "&Edit" Then
    cmdSave.Caption = "&Update"
    cmdDelete.Enabled = True
    cmdShow.Enabled = True
    cmdEdit.Caption = "&New"
    cboSupplier.SetFocus
 Else
 cmdEdit.Caption = "&Edit"
 cmdSave.Caption = "&Save"
 cmdDelete.Enabled = False
 cmdShow.Enabled = False
 txtAmount.Text = ""
 
 End If
End Sub

Private Sub cmdEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdSave_Click()
If Len(cboSupplier.Text) = 0 Then
    MsgBox "Select Supplier Name...", vbOKOnly, cnstMsgInfo
    cboSupplier.SetFocus
    Exit Sub
End If

If Len(cboPINo.Text) = 0 Then
    MsgBox "Select PO...", vbOKOnly, cnstMsgInfo
    cboPINo.SetFocus
    Exit Sub
End If

'If Len(cmbMRRID.Text) = 0 Then
'    MsgBox "Select MRRNO...", vbOKOnly, cnstMsgInfo
'    cmbMRRID.SetFocus
'    Exit Sub
'End If

If Len(cmbSubject.Text) = 0 Then
    MsgBox "Select Subject Name...", vbOKOnly, cnstMsgInfo
    cmbSubject.SetFocus
    Exit Sub
End If

If Len(txtAmount.Text) = 0 Then
    MsgBox "Enter Amount...", vbOKOnly, cnstMsgInfo
    txtAmount.SetFocus
    Exit Sub
End If

If cmdSave.Caption = "&Save" Then
    If DuplicateRecord = True Then
        Exit Sub
    End If
End If


 If cmdSave.Caption = "&Save" Then
        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    Else
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

End If

On Error GoTo ErrHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

If cmdSave.Caption = "&Save" Then
    MainComm.CommandText = "Insert Into PO_Sub_Menu(SupplierID,POID,Subject,ChargeAmt,DeleteRow,UserID,PCName,EntryDate,EntryTime) " _
                            & "Values(" & findSupplierID(cboSupplier.Text) & "," & findPIID(cboPINo.Text) & ",'" & cmbSubject.Text & "'," & CCur(txtAmount.Text) & ",0," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "') ;"
    MainComm.Execute
Else

    MainComm.CommandText = "Update PO_Sub_Menu Set DeleteRow=2,UserId=" & bytUserID & ",PCName='" & getComName & "',EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' Where SupplierID =" & findSupplierID(cboSupplier.Text) & " AND POID=" & findPIID(cboPINo.Text) & " And DeleteRow = 0 "
    MainComm.Execute
    MainComm.CommandText = "Insert Into PO_Sub_Menu(SupplierID,POID,Subject,ChargeAmt,DeleteRow,UserID,PCName,EntryDate,EntryTime) " _
                            & "Values(" & findSupplierID(cboSupplier.Text) & "," & findPIID(cboPINo.Text) & ",'" & cmbSubject.Text & "'," & CCur(txtAmount.Text) & ",0," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "') ;"
    MainComm.Execute
'MsgBox "Successfully Save", vbOKOnly, cnstMsgInfo
End If
MainConn.CommitTrans

ErrHandler:
If err.Number <> 0 Then
    MsgBox err.Description, vbCritical, cnstMsgErDB
    Screen.MousePointer = vbDefault
    MainConn.RollbackTrans
    Exit Sub
End If
 Screen.MousePointer = vbDefault
 txtAmount.Text = ""
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdShow_Click()
On Error GoTo ErrorHandler
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset

    strRecord = " SELECT Subject,ChargeAmt From PO_Sub_Menu " _
        & " Where SupplierID =" & findSupplierID(cboSupplier.Text) & " AND POID=" & findPIID(cboPINo.Text) & " And DeleteRow = 0 ;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
    cmbSubject.Text = R!Subject
    txtAmount.Text = R!ChargeAmt
   
Else
    MsgBox "Record Not Found", vbInformation, cnstMsgInfo
    cmdDelete.Enabled = False
    cmdSave.Enabled = True
    'cmdShow.Enabled = False
    txtAmount.Text = ""
        Exit Sub
    End If
    R.Close
    Set R = Nothing
ErrorHandler:
    If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub

Private Sub Form_Load()
addSupplierName cboSupplier
cmbSubject.AddItem "AIR FREIGHT CHARGE"
cmbSubject.AddItem "DISCOUNT"
cmbSubject.AddItem "SERVICE CHARGE"
cmbSubject.AddItem "OTHERS"

End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Call CurrFuncForMrrRecv(KeyAscii)
End Sub


Function DuplicateRecord() As Boolean
Set R = New ADODB.Recordset
'strRecord = " SELECT * From PO_Sub_Menu " _
'        & " Where SupplierID =" & findSupplierID(cboSupplier.Text) & " AND POID=" & findPIID(cboPINo.Text) & " And MRRID=" & findMRRID(cmbMRRID.Text) & " And DeleteRow = 0 ;"
    
   strRecord = " SELECT * From PO_Sub_Menu " _
        & " Where SupplierID =" & findSupplierID(cboSupplier.Text) & " AND POID=" & findPIID(cboPINo.Text) & "  And DeleteRow = 0 ;"
 
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        MsgBox "Duplicate Record..", vbCritical, cnstMsgInfo
        DuplicateRecord = True
        'Exit Function
     End If
End Function
