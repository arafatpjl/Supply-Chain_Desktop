VERSION 5.00
Begin VB.Form frmNewItemName 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4251
   ClientLeft      =   2834
   ClientTop       =   3094
   ClientWidth     =   5811
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNewItemName.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2516.951
   ScaleMode       =   0  'User
   ScaleWidth      =   5450.579
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
      Left            =   4560
      TabIndex        =   12
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3630
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
      Left            =   3211
      TabIndex        =   8
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3630
      Width           =   1185
   End
   Begin VB.CommandButton cmdNewItemType 
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
      Left            =   3945
      Picture         =   "frmNewItemName.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add New Item Type"
      Top             =   2730
      Width           =   285
   End
   Begin VB.ComboBox cboUnder 
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
      Left            =   1215
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "cboUnder"
      Top             =   2730
      Width           =   2679
   End
   Begin VB.ComboBox cboItem 
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
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1275
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtEditItem 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1215
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1290
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
      Left            =   1859
      TabIndex        =   3
      ToolTipText     =   "Edit/Delete Record"
      Top             =   3630
      Width           =   1185
   End
   Begin VB.TextBox txtItem 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1215
      MaxLength       =   100
      TabIndex        =   0
      Top             =   2010
      Width           =   3015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Item"
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
      Left            =   75
      TabIndex        =   9
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5835
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -15
      TabIndex        =   13
      Top             =   3405
      Width           =   5835
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   -15
      TabIndex        =   11
      Top             =   3420
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
      Index           =   1
      Left            =   -120
      TabIndex        =   10
      Top             =   855
      Width           =   6075
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Under Type of Item  :"
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
      Left            =   1215
      TabIndex        =   6
      Top             =   2490
      Width           =   1845
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name of Item  :"
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
      Left            =   1215
      TabIndex        =   4
      Top             =   1650
      Width           =   1395
   End
End
Attribute VB_Name = "frmNewItemName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboItem_Click()
  cboItem.Text = Trim(UCase(cboItem.Text))
  If Len(cboItem) = 0 Then Exit Sub
  txtEditItem.Text = cboItem.Text
  cboUnder.Text = findItemType(cboItem.Text)
End Sub
Private Sub cboItem_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub cboItem_LostFocus()
  cboItem.Text = Trim(UCase(cboItem.Text))
  If Len(cboItem) > 30 Then
    MsgBox "Maximum size is '30'", vbInformation, cnstMsgInfo
    cboItem.SetFocus
  End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()

If cmdEdit.Caption <> "&Delete" Then
  txtItem.Visible = False
  cboItem.Top = txtItem.Top
  cboItem.Visible = True
  cmdSave.Caption = "&Update"
  cmdEdit.Caption = "&Delete"
  Exit Sub
End If

If Len(cboItem) = 0 Then
  MsgBox "Select Name of Item ", vbInformation, cnstMsgInfo
  cboItem.SetFocus
  Exit Sub
End If

If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
  
MainConn.BeginTrans

MainComm.CommandText = "DELETE FROM new_Item_Name WHERE itemID = " & findItemID(cboItem.Text) & ";"
MainComm.Execute
  
MainConn.CommitTrans
  
ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
  Exit Sub
End If

On Error GoTo ErrorX

addItemNameAll cboItem

ErrorX:
If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB
Screen.MousePointer = vbDefault

End Sub
Private Sub cmdNewItemType_Click()
  strMark = "ITTYPE"
  frmNewItemType.Show 1
End Sub
Private Sub cmdSave_Click()

Dim lngLastID As Long

If cmdSave.Caption = "&Save" Then
  If Len(txtItem) = 0 Then
    MsgBox "Type Name of Item", vbInformation, cnstMsgInfo
    txtItem.SetFocus
    Exit Sub
  End If
Else
  If Len(cboItem) = 0 Then
    MsgBox "Select Name of Item", vbInformation, cnstMsgInfo
    cboItem.SetFocus
    Exit Sub
  End If
End If

If Len(cboUnder) = 0 Then
  MsgBox "Select Under of Item", vbInformation, cnstMsgInfo
  cboUnder.SetFocus
  Exit Sub
End If

On Error GoTo ErrorX

Set R = New ADODB.Recordset

If cmdSave.Caption = "&Save" Then
  If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  
  lngLastID = 1
  strRecord = "SELECT Max(itemID) AS LastID FROM InventoryPJL.DBO.new_Item_Name HAVING Max(itemID) Is Not Null;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then lngLastID = lngLastID + R![LastID]
  R.Close

Else
  If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub

  If cboItem.Text <> txtEditItem.Text Then
    strRecord = "SELECT itemName FROM new_Item_Name WHERE itemName = '" & cboItem.Text & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      MsgBox "Duplicate Name of Item", vbInformation, cnstMsgInfo
      cboItem.SetFocus
      Exit Sub
    End If
    R.Close
  End If
End If

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans
  
If cmdSave.Caption = "&Save" Then
  MainComm.CommandText = "INSERT INTO InventoryPJL.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryJ2L.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPAL.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryUJL.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNHT.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNHTHOLD.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPJLCA.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPJLFG.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNZNF.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNZNG.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryNZND.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryHUAL.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPWP.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryPCL.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryCRP.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryB3OP.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute
  MainComm.CommandText = "INSERT INTO InventoryB5MN.DBO.new_Item_Name(itemID,itemName,itemType)VALUES(" & lngLastID & ",'" & txtItem.Text & "','" & cboUnder.Text & "');"
  MainComm.Execute

Else
  
  MainComm.CommandText = "UPDATE InventoryPJL.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryJ2L.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryPAL.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryUJL.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryNHT.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryNHTHOLD.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryPJLCA.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryPJLFG.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryNZNF.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryNZNG.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryNZND.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryHUAL.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryPWP.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryPCL.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryCRP.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryB3OP.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  MainComm.CommandText = "UPDATE InventoryB5MN.DBO.new_Item_Name SET itemName='" & cboItem.Text & "',itemType='" & cboUnder.Text & "' WHERE itemID=" & findItemID(txtEditItem.Text) & ";"
  MainComm.Execute
  
End If

MainConn.CommitTrans

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
  Exit Sub
End If

On Error GoTo ErrorX

addItemNameAll cboItem

txtItem.Text = ""
txtItem.SetFocus

Set R = Nothing

ErrorX:
If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB
Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Activate()
  If strMark = "ITTYPE" Then addItemType cboUnder
End Sub
Private Sub Form_Load()
  addItemNameAll cboItem
  addItemType cboUnder
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If strMark = "ITTYPE" Then strMark = ""
  Set frmNewItemName = Nothing
End Sub
Private Sub txtItem_GotFocus()
  txtGotFocus txtItem
End Sub
Private Sub txtItem_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtItem_LostFocus()
  txtItem.Text = Trim(UCase(txtItem.Text))
  If Len(txtItem) = 0 Then Exit Sub

  On Error GoTo ErrorX
  
  Set R = New ADODB.Recordset
  strRecord = "SELECT itemName FROM new_Item_Name WHERE itemName='" & txtItem.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    MsgBox "Duplicate Name of Item", vbInformation, cnstMsgInfo
    txtItem.SetFocus
  End If
  R.Close
  Set R = Nothing

ErrorX:
  If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB

End Sub



