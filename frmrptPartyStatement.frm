VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrptPartyStatement 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Party Statement"
   ClientHeight    =   5278
   ClientLeft      =   39
   ClientTop       =   429
   ClientWidth     =   6786
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmrptPartyStatement.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5278
   ScaleWidth      =   6786
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSummary 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Party Bill Amount Summary"
      Height          =   247
      Left            =   650
      TabIndex        =   20
      Top             =   3822
      Width           =   3328
   End
   Begin VB.ComboBox cmbQueryType 
      Height          =   286
      ItemData        =   "frmrptPartyStatement.frx":000C
      Left            =   450
      List            =   "frmrptPartyStatement.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1950
      Width           =   3562
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      Picture         =   "frmrptPartyStatement.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1260
      Width           =   345
   End
   Begin VB.TextBox txtNameOfParty 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   1260
      Width           =   4455
   End
   Begin VB.TextBox txtPartyID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   1260
      Width           =   885
   End
   Begin VB.CheckBox chkAllCompany 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All Company"
      Height          =   247
      Left            =   650
      TabIndex        =   5
      Top             =   4173
      Visible         =   0   'False
      Width           =   2805
   End
   Begin MSComCtl2.DTPicker DTPickerToDate 
      Height          =   377
      Left            =   2275
      TabIndex        =   4
      Top             =   3315
      Width           =   1755
      _ExtentX        =   3091
      _ExtentY        =   671
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.5094
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124583937
      CurrentDate     =   39284
   End
   Begin MSComCtl2.DTPicker DTPickerFromDate 
      Height          =   375
      Left            =   450
      TabIndex        =   3
      Top             =   3315
      Width           =   1785
      _ExtentX        =   3139
      _ExtentY        =   671
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.5094
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124583937
      CurrentDate     =   39284
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Close"
      Height          =   416
      Left            =   3861
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4732
      Width           =   1274
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Preview"
      Height          =   416
      Left            =   5343
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4732
      Width           =   1245
   End
   Begin VB.ComboBox cmbStatementType 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   286
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2580
      Width           =   3562
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3330
      Top             =   4620
      _ExtentX        =   671
      _ExtentY        =   671
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   4560
      Width           =   6795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report ::  Party Statement"
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
      TabIndex        =   18
      Top             =   120
      Width           =   3345
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6795
   End
   Begin VB.Label Label4 
      Caption         =   "Query Type"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   450
      TabIndex        =   16
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Party ID"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   450
      TabIndex        =   12
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   39
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   4524
      Width           =   6799
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   2
      Left            =   450
      TabIndex        =   10
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H00000080&
      Height          =   221
      Index           =   1
      Left            =   2288
      TabIndex        =   9
      Top             =   3055
      Width           =   1157
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Type "
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   8
      Top             =   2340
      Width           =   1725
   End
End
Attribute VB_Name = "frmrptPartyStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()

If Len(Trim(txtPartyID)) = 0 And chkSummary.Value = 0 Then
  MsgBox "You should provide Party ID", vbCritical, App.Title
  txtPartyID.SetFocus
  Exit Sub
End If
If Len(cmbStatementType) = 0 Then
  MsgBox "You should select Statement Type", vbCritical, App.Title
  cmbStatementType.SetFocus
  Exit Sub
End If

If chkSummary.Value = 1 Then
  If cmbQueryType.Text <> "Bill Date" Then
    MsgBox "You must select 'Bill Date' in Query Type for Party Summary", vbCritical, App.Title
    cmbQueryType.SetFocus
    Exit Sub
  End If
  If cmbStatementType.Text <> "All Type" Then
    MsgBox "You must select 'All Type' in Statement Type for Party Summary", vbCritical, App.Title
    cmbStatementType.SetFocus
    Exit Sub
  End If
End If

Screen.MousePointer = vbHourglass

If chkAllCompany.Value = 0 Then GoTo showReport

On Error GoTo ErrorHandler

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

MainComm.CommandText = "DELETE FROM Temp_Statement_of_Cash_and_Credit"
MainComm.Execute

If cmbQueryType.Text = "Bill Date" Then
  If cmbStatementType.Text = "All Type" Then
    
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJL.dbo.viewStatement_Cash_Credit Where CompanyID=1 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryUJL.dbo.viewStatement_Cash_Credit Where CompanyID=7 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryJ2L.dbo.viewStatement_Cash_Credit Where CompanyID=5 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPAL.dbo.viewStatement_Cash_Credit Where CompanyID=3 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHT.dbo.viewStatement_Cash_Credit Where CompanyID=13 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHTHOLD.dbo.viewStatement_Cash_Credit Where CompanyID=14 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLCA.dbo.viewStatement_Cash_Credit Where CompanyID=15 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLFG.dbo.viewStatement_Cash_Credit Where CompanyID=16 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryHUAL.dbo.viewStatement_Cash_Credit Where CompanyID=20 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPWP.dbo.viewStatement_Cash_Credit Where CompanyID=21 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPCL.dbo.viewStatement_Cash_Credit Where CompanyID=22 AND BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute

  ElseIf cmbStatementType.Text = "Cheque" Then
     
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJL.dbo.viewStatement_Cash_Credit Where CompanyID=1 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryUJL.dbo.viewStatement_Cash_Credit Where CompanyID=7 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryJ2L.dbo.viewStatement_Cash_Credit Where CompanyID=5 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
        
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPAL.dbo.viewStatement_Cash_Credit Where CompanyID=3 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHT.dbo.viewStatement_Cash_Credit Where CompanyID=13 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHTHOLD.dbo.viewStatement_Cash_Credit Where CompanyID=14 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLCA.dbo.viewStatement_Cash_Credit Where CompanyID=15 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLFG.dbo.viewStatement_Cash_Credit Where CompanyID=16 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryHUAL.dbo.viewStatement_Cash_Credit Where CompanyID=20 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPWP.dbo.viewStatement_Cash_Credit Where CompanyID=21 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPCL.dbo.viewStatement_Cash_Credit Where CompanyID=22 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
        
  Else

      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJL.dbo.viewStatement_Cash_Credit Where CompanyID=1 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryUJL.dbo.viewStatement_Cash_Credit Where CompanyID=7 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryJ2L.dbo.viewStatement_Cash_Credit Where CompanyID=5 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPAL.dbo.viewStatement_Cash_Credit Where CompanyID=3 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHT.dbo.viewStatement_Cash_Credit Where CompanyID=13 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHTHOLD.dbo.viewStatement_Cash_Credit Where CompanyID=14 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLCA.dbo.viewStatement_Cash_Credit Where CompanyID=15 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLFG.dbo.viewStatement_Cash_Credit Where CompanyID=16 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryHUAL.dbo.viewStatement_Cash_Credit Where CompanyID=20 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPWP.dbo.viewStatement_Cash_Credit Where CompanyID=21 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPCL.dbo.viewStatement_Cash_Credit Where CompanyID=22 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
  End If
  
Else
  If cmbStatementType.Text = "All Type" Then
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJL.dbo.viewStatement_Cash_Credit Where CompanyID=1 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryUJL.dbo.viewStatement_Cash_Credit Where CompanyID=7 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryJ2L.dbo.viewStatement_Cash_Credit Where CompanyID=5 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPAL.dbo.viewStatement_Cash_Credit Where CompanyID=3 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHT.dbo.viewStatement_Cash_Credit Where CompanyID=13 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHTHOLD.dbo.viewStatement_Cash_Credit Where CompanyID=14 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLCA.dbo.viewStatement_Cash_Credit Where CompanyID=15 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLFG.dbo.viewStatement_Cash_Credit Where CompanyID=16 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryHUAL.dbo.viewStatement_Cash_Credit Where CompanyID=20 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPWP.dbo.viewStatement_Cash_Credit Where CompanyID=21 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPCL.dbo.viewStatement_Cash_Credit Where CompanyID=22 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No"
      MainComm.Execute
          
  ElseIf cmbStatementType.Text = "Cheque" Then

      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJL.dbo.viewStatement_Cash_Credit Where CompanyID=1 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryUJL.dbo.viewStatement_Cash_Credit Where CompanyID=7 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryJ2L.dbo.viewStatement_Cash_Credit Where CompanyID=5 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPAL.dbo.viewStatement_Cash_Credit Where CompanyID=3 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHT.dbo.viewStatement_Cash_Credit Where CompanyID=13 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHTHOLD.dbo.viewStatement_Cash_Credit Where CompanyID=14 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLCA.dbo.viewStatement_Cash_Credit Where CompanyID=15 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLFG.dbo.viewStatement_Cash_Credit Where CompanyID=16 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryHUAL.dbo.viewStatement_Cash_Credit Where CompanyID=20 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPWP.dbo.viewStatement_Cash_Credit Where CompanyID=21 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPCL.dbo.viewStatement_Cash_Credit Where CompanyID=22 AND SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
        
  Else
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJL.dbo.viewStatement_Cash_Credit Where CompanyID=1 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryUJL.dbo.viewStatement_Cash_Credit Where CompanyID=7 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryJ2L.dbo.viewStatement_Cash_Credit Where CompanyID=5 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPAL.dbo.viewStatement_Cash_Credit Where CompanyID=3 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHT.dbo.viewStatement_Cash_Credit Where CompanyID=13 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryNHTHOLD.dbo.viewStatement_Cash_Credit Where CompanyID=14 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLCA.dbo.viewStatement_Cash_Credit Where CompanyID=15 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPJLFG.dbo.viewStatement_Cash_Credit Where CompanyID=16 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryHUAL.dbo.viewStatement_Cash_Credit Where CompanyID=20 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPWP.dbo.viewStatement_Cash_Credit Where CompanyID=21 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
      MainComm.CommandText = " INSERT INTO Temp_Statement_of_Cash_and_Credit SELECT * FROM InventoryPCL.dbo.viewStatement_Cash_Credit Where CompanyID=22 AND SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      MainComm.Execute
  End If
End If

MainConn.CommitTrans

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  err.Number = 0
  Exit Sub
End If

showReport:

On Error GoTo ErrorX

If cmbQueryType.Text = "Bill Date" Then
  
  If chkAllCompany.Value = 0 Then
    
    If chkSummary.Value = 1 Then
      Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatementSummary.rpt")
      
      If Len(Trim(txtPartyID)) > 0 Then
        strRecord = "SELECT CompanyID,CompName,CompAdd,SupplierID,SupplierName,SUM(Amount) AS Amount FROM viewStatement_Cash_Credit WHERE CompanyID = " & bytCompID & " AND SupplierID = '" & txtPartyID.Text & "'" _
                  & " AND BillDate Between '" & DTPickerFromDate.Value & "' AND '" & DTPickerToDate.Value & "' GROUP BY CompanyID,CompName,CompAdd,SupplierID,SupplierName ORDER BY SupplierName;"
      Else
        strRecord = "SELECT CompanyID,CompName,CompAdd,SupplierID,SupplierName,SUM(Amount) AS Amount FROM viewStatement_Cash_Credit WHERE CompanyID = " & bytCompID & "" _
                  & " AND BillDate Between '" & DTPickerFromDate.Value & "' AND '" & DTPickerToDate.Value & "' GROUP BY CompanyID,CompName,CompAdd,SupplierID,SupplierName ORDER BY SupplierName;"
      End If
    
      Report.FormulaFields.GetItemByName("FromDate").Text = "'" & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & "'"
      Report.FormulaFields.GetItemByName("ToDate").Text = "'" & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & "'"

    Else
      Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatement.rpt")
      
      If cmbStatementType.Text = "All Type" Then
        strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      ElseIf cmbStatementType.Text = "Cheque" Then
        strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      Else
        strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      End If
      Report.FormulaFields(5).Text = "' " & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & " '"
      Report.FormulaFields(6).Text = "' " & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & " '"
    End If
  
  Else
  
    If chkSummary.Value = 1 Then
      Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatementSummary.rpt")

      Report.FormulaFields.GetItemByName("AllComp").Text = "'Y'"
      Report.FormulaFields.GetItemByName("FromDate").Text = "'" & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & "'"
      Report.FormulaFields.GetItemByName("ToDate").Text = "'" & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & "'"

    Else
      Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatement_AllComp.rpt")
      
      If cmbStatementType.Text = "All Type" Then Report.FormulaFields(8).Text = "'All Transaction'"
      
      Report.FormulaFields(7).Text = "'All Company'"
      Report.FormulaFields(5).Text = "' " & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & " '"
      Report.FormulaFields(6).Text = "' " & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & " '"
    End If
    
    If cmbStatementType.Text = "All Type" Then
      If chkSummary.Value = 1 Then
        strRecord = "SELECT '0' AS CompanyID,'' AS CompName,'' AS CompAdd,SupplierID,SupplierName,SUM(Amount) AS Amount FROM Temp_Statement_of_Cash_and_Credit GROUP BY SupplierID,SupplierName ORDER BY SupplierName;"
      Else
        If Len(Trim(txtPartyID)) > 0 Then
          strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit WHERE SupplierID = '" & txtPartyID.Text & "';"
        Else
          strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit;"
        End If
      End If

      'strRecord = "Select * From viewStatement_Cash_Credit Where SupplierID = '" & txtPartyID.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
      '-----------------
    ElseIf cmbStatementType.Text = "Cheque" Then
              
      ''If chkSummary.Value = 1 Then
        ''strRecord = "SELECT CompanyID,CompName,CompAdd,SupplierID,SupplierName,SUM(Amount) AS Amount FROM Temp_Statement_of_Cash_and_Credit GROUP BY CompanyID,CompName,CompAdd,SupplierID,SupplierName ORDER BY SupplierName;"
      ''Else
        If Len(Trim(txtPartyID)) > 0 Then
          strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit WHERE SupplierID = '" & txtPartyID.Text & "';"
        Else
          strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit;"
        End If
      ''End If
      'strRecord = "Select * From viewStatement_Cash_Credit Where SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    Else
    
      ''If chkSummary.Value = 1 Then
      ''  strRecord = "SELECT CompanyID,CompName,CompAdd,SupplierID,SupplierName,SUM(Amount) AS Amount FROM Temp_Statement_of_Cash_and_Credit GROUP BY CompanyID,CompName,CompAdd,SupplierID,SupplierName ORDER BY SupplierName;"
      ''Else
        If Len(Trim(txtPartyID)) > 0 Then
          strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit WHERE SupplierID = '" & txtPartyID.Text & "';"
        Else
          strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit;"
        End If
      ''End If
      'strRecord = "Select * From viewStatement_Cash_Credit Where SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And BillDate >= '" & DTPickerFromDate.Value & "' And BillDate <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    End If
  End If

Else
  If chkAllCompany.Value = 0 Then
    Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatement.rpt")
    
    If cmbStatementType.Text = "All Type" Then
      strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    ElseIf cmbStatementType.Text = "Cheque" Then
      strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    Else
      strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    End If

    Report.FormulaFields(5).Text = "' " & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & " '"
    Report.FormulaFields(6).Text = "' " & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & " '"
  
  Else
    Set Report = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatement_AllComp.rpt")
    
    If cmbStatementType.Text = "All Type" Then
      
      If Len(Trim(txtPartyID)) > 0 Then
        strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit WHERE SupplierID = '" & txtPartyID.Text & "';"
      Else
        strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit;"
      End If
      ''''''''''''''
      'strRecord = "Select * From viewStatement_Cash_Credit Where SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
  
      Report.FormulaFields(8).Text = "'All Transaction'"
    
    ElseIf cmbStatementType.Text = "Cheque" Then
        
      If Len(Trim(txtPartyID)) > 0 Then
        strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit WHERE SupplierID = '" & txtPartyID.Text & "';"
      Else
        strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit;"
      End If
      ''''''''''''''
      'strRecord = "Select * From viewStatement_Cash_Credit Where SupplierID = '" & txtPartyID.Text & "' And (((Tran_Type) = '" & "Cheque" & "') Or ((Tran_Type) = '" & "Cash Cheque" & "')) And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    
    Else
    
      If Len(Trim(txtPartyID)) > 0 Then
        strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit WHERE SupplierID = '" & txtPartyID.Text & "';"
      Else
        strRecord = "SELECT * FROM Temp_Statement_of_Cash_and_Credit;"
      End If
      ''''''''''''''
      'strRecord = "Select * From viewStatement_Cash_Credit Where SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbStatementType.Text & "' And Tran_Date >= '" & DTPickerFromDate.Value & "' And Tran_Date <= '" & DTPickerToDate.Value & "' Order By Tr_No;"
    End If

    Report.FormulaFields(7).Text = "'All Company'"
    Report.FormulaFields(5).Text = "' " & Format(DTPickerFromDate.Value, "DD-MMM-YYYY") & " '"
    Report.FormulaFields(6).Text = "' " & Format(DTPickerToDate.Value, "DD-MMM-YYYY") & " '"
  End If
End If

Set R = New ADODB.Recordset
R.Open strRecord, MainConn, adOpenStatic

Report.DiscardSavedData
Report.Database.SetDataSource R
frmMainReport.CRViewer.Refresh
frmMainReport.CRViewer.ReportSource = Report
frmMainReport.CRViewer.ViewReport
frmMainReport.CRViewer.Zoom (82)
frmMainReport.Show 1
R.Close

ErrorX:
If err.Number <> 0 Then MsgBox err.Description, vbCritical, App.Title
Screen.MousePointer = vbDefault
Set R = Nothing
End Sub
Private Sub Command1_Click()
  frmSerchForm.txtFormName.Text = "Party ID and Name for PartyStatement"
  frmSerchForm.Show 1
End Sub
Private Sub Form_Load()

Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2

cmbStatementType.AddItem "All Type"
cmbStatementType.AddItem "Cash"
cmbStatementType.AddItem "Cheque"
cmbStatementType.AddItem "Credit"
cmbStatementType.AddItem "Cash Cheque"
    
Call prcLoadQueryType
    
DTPickerFromDate.Value = Date
DTPickerToDate.Value = Date

If bytCompID = 1 Then chkAllCompany.Visible = True

End Sub

Private Sub prcLoadQueryType()
  cmbQueryType.Clear
  cmbQueryType.AddItem "Transaction Date"
  cmbQueryType.AddItem "Bill Date"
End Sub
Private Sub cmbStatementType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPartyID_GotFocus()
    txtPartyID.SelStart = 0
    txtPartyID.SelLength = Len(txtPartyID.Text)
End Sub

Private Sub txtPartyID_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtPartyID_LostFocus()
Dim sql1 As String
    If Val(txtPartyID.Text) > 0 And Val(txtPartyID.Text) < 10 Then
       txtPartyID.Text = "000" & Val(txtPartyID.Text)
    ElseIf Val(txtPartyID.Text) > 9 And Val(txtPartyID.Text) < 100 Then
       txtPartyID.Text = "00" & Val(txtPartyID.Text)
    ElseIf Val(txtPartyID.Text) > 99 And Val(txtPartyID.Text) < 1000 Then
       txtPartyID.Text = "0" & Val(txtPartyID.Text)
    ElseIf Val(txtPartyID.Text) > 999 And Val(txtPartyID.Text) < 10000 Then
       txtPartyID.Text = Val(txtPartyID.Text)
    End If
    If txtPartyID.Text = "" Then
    Else
        On Error GoTo error1
        Set R = New Recordset
            sql1 = "Select SupplierID, SupplierName, SupplierAddress, GoodsName From Supplier_Information Where (((SupplierID)='" & txtPartyID.Text & "'))"
            R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
error1:
        If err.Number = 3021 Then
            Set R = Nothing
            MsgBox "This Is Not Valit Number", 16
            txtPartyID.SetFocus
        Else
            txtNameOfParty.Text = R("SupplierName").Value
'            txtParticulers.Text = R("GoodsName").Value & ""
            Set R = Nothing
        End If
    End If
End Sub
