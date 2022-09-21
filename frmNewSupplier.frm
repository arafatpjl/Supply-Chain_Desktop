VERSION 5.00
Begin VB.Form frmNewSupplier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4979
   ClientLeft      =   4264
   ClientTop       =   2886
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4979
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4575
      TabIndex        =   14
      ToolTipText     =   "Edit/Delete Record"
      Top             =   4335
      Width           =   1185
   End
   Begin VB.ComboBox cboCountry 
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
      Height          =   315
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3570
      Width           =   3630
   End
   Begin VB.CommandButton cmdAddCountry 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Add New Name of Country"
      Top             =   3570
      Width           =   285
   End
   Begin VB.TextBox txtEditSupplier 
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
      Height          =   315
      Left            =   990
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2940
      Visible         =   0   'False
      Width           =   3630
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
      Height          =   315
      Left            =   990
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1050
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
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
      Left            =   2055
      TabIndex        =   4
      ToolTipText     =   "Edit/Delete Record"
      Top             =   4335
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
      Caption         =   "&Save"
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
      Left            =   3315
      TabIndex        =   3
      ToolTipText     =   "Save/Update Record"
      Top             =   4335
      Width           =   1185
   End
   Begin VB.TextBox txtSupplier 
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
      Height          =   315
      Left            =   990
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1620
      Width           =   3630
   End
   Begin VB.TextBox txtAddress 
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
      Height          =   1005
      Left            =   990
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2250
      Width           =   3630
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   5850
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   0
      TabIndex        =   15
      Top             =   4125
      Width           =   5955
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Top             =   4140
      Width           =   5835
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   -120
      TabIndex        =   12
      Top             =   4440
      Width           =   4305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Information"
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
      Left            =   135
      TabIndex        =   11
      Top             =   210
      Width           =   2700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Country  :"
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
      Left            =   990
      TabIndex        =   10
      Top             =   3330
      Width           =   810
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address / Details  :"
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
      Height          =   255
      Left            =   990
      TabIndex        =   7
      Top             =   2010
      Width           =   1605
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supplier Name  :"
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
      Height          =   255
      Left            =   990
      TabIndex        =   6
      Top             =   1320
      Width           =   1605
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmNewSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSupplier_Click()
  If Len(cboSupplier) = 0 Then Exit Sub
  Set R = New ADODB.Recordset
  strRecord = "SELECT suppName,suppAddress,suppCountry FROM new_Supplier WHERE suppName='" & cboSupplier.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    txtEditSupplier.Text = R![suppName]
    txtAddress.Text = R![suppAddress] & ""
    IIf IsNull(R![suppCountry]), cboCountry.ListIndex = -1, cboCountry.Text = R![suppCountry]
  End If
  R.Close
  Set R = Nothing
End Sub
Private Sub cboSupplier_LostFocus()
  cboSupplier.Text = Trim(UCase(cboSupplier.Text))
  If Len(cboSupplier) > 50 Then
    MsgBox "Maximum size is '50'", vbInformation, cnstMsgInfo
    cboSupplier.SetFocus
  End If
End Sub
Private Sub cmdAddCountry_Click()
  frmNewCountry.Show 1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()

If cmdEdit.Caption = "&Delete" Then

  If Len(cboSupplier) = 0 Then
    MsgBox "Select name of supplier", vbInformation, cnstMsgInfo
    cboSupplier.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  On Error GoTo ErrorHandler
  
  Set MainComm = New ADODB.Command
  MainComm.ActiveConnection = MainConn
  
  ''MainComm.CommandText = "DELETE FROM InventoryPJL.DBO.new_Supplier WHERE suppName = '" & cboSupplier.Text & "';"
  MainComm.CommandText = "DELETE FROM new_Supplier WHERE suppName = '" & cboSupplier.Text & "';"
  MainComm.Execute
  
  addSupplierName cboSupplier
  txtAddress.Text = ""
  cboCountry.ListIndex = -1
  cboSupplier.SetFocus
Else
  txtSupplier.Visible = False
  cboSupplier.Top = txtSupplier.Top
  cboSupplier.Visible = True
  cmdSave.Caption = "&Update"
  cmdEdit.Caption = "&Delete"
  Exit Sub
End If

ErrorHandler:
If err.Number <> 0 Then
  If err.Number = -2147467259 Then
    MsgBox "This record cannot be deleted", vbCritical, cnstMsgErDB
  Else
    MsgBox err.Description, vbCritical, cnstMsgErDB
  End If
End If

End Sub
Private Sub cmdSave_Click()

Dim lngLastID As Integer

Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

If cmdSave.Caption = "&Save" Then

  If Len(txtSupplier) = 0 Then
    MsgBox "Type name of supplier", vbInformation, cnstMsgInfo
    txtSupplier.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  lngLastID = 1
  strRecord = "SELECT Max(supplierID) AS LastID FROM InventoryPJL.DBO.new_Supplier HAVING Max(supplierID) Is Not Null;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then lngLastID = lngLastID + R![LastID]
  R.Close

  MainConn.BeginTrans
    
    MainComm.CommandText = "INSERT INTO InventoryPJL.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
  
    MainComm.CommandText = "INSERT INTO InventoryJ2L.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
  
    MainComm.CommandText = "INSERT INTO InventoryPAL.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
  
    MainComm.CommandText = "INSERT INTO InventoryUJL.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryNHT.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryNHTHOLD.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryPJLCA.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryPJLFG.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
''    MainComm.CommandText = "INSERT INTO InventoryNZNF.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
''        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
''    MainComm.Execute
''
''    MainComm.CommandText = "INSERT INTO InventoryNZNG.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
''        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
''    MainComm.Execute
''
''    MainComm.CommandText = "INSERT INTO InventoryNZND.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
''        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
''    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryPCL.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryPWP.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryHUAL.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryCRP.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryB5MN.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO InventoryB3OP.DBO.new_Supplier(supplierID,suppName,suppAddress,suppCountry)" _
        & " VALUES(" & lngLastID & ",'" & txtSupplier.Text & "','" & txtAddress.Text & "','" & cboCountry.Text & "');"
    MainComm.Execute
    
  MainConn.CommitTrans
  
  txtSupplier.Text = ""
  txtAddress.Text = ""
  cboCountry.ListIndex = -1
  txtSupplier.SetFocus

Else

  If Len(cboSupplier) = 0 Then
    MsgBox "Select name of supplier", vbInformation, cnstMsgInfo
    cboSupplier.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

  If cboSupplier.Text <> txtEditSupplier.Text Then
    strRecord = "SELECT suppName FROM new_Supplier WHERE suppName='" & cboSupplier.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      MsgBox "Duplicate name of supplier", vbInformation, cnstMsgInfo
      cboSupplier.SetFocus
      Exit Sub
    End If
    R.Close
  End If
  
  MainConn.BeginTrans
  
    MainComm.CommandText = "UPDATE InventoryPJL.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
  
    MainComm.CommandText = "UPDATE InventoryJ2L.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
  
    MainComm.CommandText = "UPDATE InventoryPAL.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
  
    MainComm.CommandText = "UPDATE InventoryUJL.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute

    MainComm.CommandText = "UPDATE InventoryNHT.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryNHTHOLD.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryPJLCA.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryPJLFG.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
''    MainComm.CommandText = "UPDATE InventoryNZNF.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
''        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
''    MainComm.Execute
''
''    MainComm.CommandText = "UPDATE InventoryNZNG.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
''        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
''    MainComm.Execute
''
''    MainComm.CommandText = "UPDATE InventoryNZND.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
''        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
''    MainComm.Execute

    MainComm.CommandText = "UPDATE InventoryPCL.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryPWP.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryHUAL.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryCRP.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryB5MN.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE InventoryB3OP.DBO.new_Supplier SET suppName='" & cboSupplier.Text & "'," _
        & " suppAddress='" & txtAddress.Text & "',suppCountry='" & cboCountry.Text & "' WHERE suppName='" & txtEditSupplier.Text & "';"
    MainComm.Execute
    
  MainConn.CommitTrans

  txtAddress.Text = ""
  cboCountry.ListIndex = -1
  cboSupplier.SetFocus

End If
Set R = Nothing
addSupplierName cboSupplier
End Sub
Private Sub Form_Load()
  addSupplierName cboSupplier
  addCountry cboCountry
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmNewSupplier = Nothing
End Sub

Private Sub txtAddress_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtAddress_LostFocus()
  'txtAddress.Text = Trim(UCase(txtAddress.Text))
  txtAddress.Text = Trim(txtAddress.Text)
End Sub
Private Sub txtSupplier_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtSupplier_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtSupplier_LostFocus()
  If Len(txtSupplier) = 0 Then Exit Sub
  txtSupplier.Text = Trim(UCase(txtSupplier.Text))
  Set R = New ADODB.Recordset
  strRecord = "SELECT suppName FROM new_Supplier WHERE suppName='" & txtSupplier.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    MsgBox "Duplicate name of supplier", vbInformation, cnstMsgInfo
    txtSupplier.SetFocus
  End If
  R.Close
  Set R = Nothing
End Sub
