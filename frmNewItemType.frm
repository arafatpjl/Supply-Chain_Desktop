VERSION 5.00
Begin VB.Form frmNewItemType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4251
   ClientLeft      =   2834
   ClientTop       =   3094
   ClientWidth     =   5811
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNewItemType.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2516.948
   ScaleMode       =   0  'User
   ScaleWidth      =   5450.58
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
      Left            =   4515
      TabIndex        =   8
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3585
      Width           =   1185
   End
   Begin VB.ComboBox cboItemType 
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
      Left            =   1170
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtEditItemType 
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
      IMEMode         =   3  'DISABLE
      Left            =   1170
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1485
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   1890
      TabIndex        =   3
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3585
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
      Left            =   3195
      TabIndex        =   2
      ToolTipText     =   "Save/Update Record"
      Top             =   3585
      Width           =   1185
   End
   Begin VB.TextBox txtItemType 
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
      IMEMode         =   3  'DISABLE
      Left            =   1170
      MaxLength       =   100
      TabIndex        =   0
      Top             =   2205
      Width           =   3015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Group"
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
      TabIndex        =   6
      Top             =   225
      Width           =   1470
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   0
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
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   5835
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   5835
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   3405
      Width           =   5835
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name of Item Type"
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
      Index           =   0
      Left            =   1170
      TabIndex        =   4
      Top             =   1845
      Width           =   1695
   End
End
Attribute VB_Name = "frmNewItemType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboItemType_Click()
  If cboItemType.Text = "" Then Exit Sub
  cboItemType.Text = Trim(UCase(cboItemType.Text))
  txtEditItemType.Text = cboItemType.Text
End Sub
Private Sub cboItemType_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub cboItemType_LostFocus()
  cboItemType.Text = Trim(UCase(cboItemType.Text))
  If Len(cboItemType) > 30 Then
    MsgBox "Maximum size is '30'", vbInformation, cnstMsgInfo
    cboItemType.SetFocus
  End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()

If cmdEdit.Caption = "&Delete" Then

  If Len(cboItemType) = 0 Then
    MsgBox "Select Type of Item ", vbInformation, cnstMsgInfo
    cboItemType.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

  On Error GoTo ErrorHandler
  
  Set MainComm = New ADODB.Command
  MainComm.ActiveConnection = MainConn
  
  MainComm.CommandText = "DELETE  FROM new_Item_Type WHERE itemType='" & cboItemType.Text & "';"
  MainComm.Execute
  
  addItemType cboItemType
  Screen.MousePointer = vbDefault

Else
  txtItemType.Visible = False
  cboItemType.Top = txtItemType.Top
  cboItemType.Visible = True
  cmdSave.Caption = "&Update"
  cmdEdit.Caption = "&Delete"
  Exit Sub
End If

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
End If

End Sub
Private Sub cmdSave_Click()

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

If cmdSave.Caption = "&Save" Then

  If Len(txtItemType) = 0 Then
    MsgBox "Type Name of Item Type", vbInformation, cnstMsgInfo
    txtItemType.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  MainConn.BeginTrans
  
  MainComm.CommandText = "INSERT INTO InventoryPJL.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryJ2L.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPAL.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryUJL.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNHT.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNHTHOLD.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPJLCA.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPJLFG.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNZNF.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNZNG.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNZND.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPCL.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPWP.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryHUAL.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryCRP.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryB3OP.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryB5MN.DBO.new_Item_Type(itemType,itemGroup,itemShortName)VALUES('" & txtItemType.Text & "','" & txtItemType.Text & "','" & Mid(txtItemType.Text, 1, 3) & "');"
  MainComm.Execute
  
  MainConn.CommitTrans
  
  txtItemType.Text = ""
  txtItemType.SetFocus

Else

  If Len(cboItemType) = 0 Then
    MsgBox "Select Type of Item", vbInformation, cnstMsgInfo
    cboItemType.SetFocus
    Exit Sub
  End If

  If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

  If cboItemType.Text <> txtEditItemType.Text Then
    Set R = New ADODB.Recordset
    ''***************PJL*********************
    strRecord = "SELECT itemType FROM new_Item_Type WHERE itemType='" & cboItemType.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      MsgBox "Duplicate Type of Item", vbInformation, cnstMsgInfo
      cboItemType.SetFocus
      Exit Sub
    End If
    R.Close
  End If
  
  MainConn.BeginTrans
  
    MainComm.CommandText = "UPDATE InventoryPJL.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryJ2L.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryPAL.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryUJL.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryNHT.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryNHTHOLD.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryPJLCA.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryPJLFG.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryNZNF.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryNZNG.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryPCL.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryPWP.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryHUAL.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryCRP.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryB3OP.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute
    MainComm.CommandText = "UPDATE InventoryB5MN.DBO.new_Item_Type SET itemType='" & cboItemType.Text & "',itemGroup='" & cboItemType.Text & "', itemShortName = '" & Mid(cboItemType.Text, 1, 3) & "' WHERE itemType='" & txtEditItemType.Text & "';"
    MainComm.Execute

  MainConn.CommitTrans

End If

addItemType cboItemType

End Sub
Private Sub Form_Load()
  addItemType cboItemType
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmNewItemType = Nothing
End Sub
Private Sub txtItemType_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtItemType_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtItemType_LostFocus()
If txtItemType.Text = "" Then Exit Sub
txtItemType.Text = Trim(UCase(txtItemType.Text))

Set R = New ADODB.Recordset
strRecord = "SELECT itemType FROM new_Item_Type WHERE itemType='" & txtItemType.Text & "';"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  MsgBox "Duplicate Type of Item", vbInformation, cnstMsgInfo
  txtItemType.SetFocus
End If
R.Close
End Sub
