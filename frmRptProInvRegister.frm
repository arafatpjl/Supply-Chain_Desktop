VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptProInvRegister 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmRptProInvRegister.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2115
      Left            =   945
      TabIndex        =   8
      Top             =   1080
      Width           =   3915
      Begin VB.TextBox txtPINo 
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
         IMEMode         =   3  'DISABLE
         Left            =   390
         MaxLength       =   30
         TabIndex        =   2
         Top             =   900
         Width           =   3285
      End
      Begin VB.OptionButton optPI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Particular PI"
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
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   630
         Width           =   1935
      End
      Begin VB.OptionButton optAllPI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show All PI"
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
         Height          =   225
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton optSupplier 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Particular Supplier"
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
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   1320
         Width           =   2355
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
         Height          =   315
         Left            =   390
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1590
         Width           =   3285
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C000&
      Caption         =   "&Print"
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
      Left            =   3135
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   3615
      Width           =   1185
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
      Left            =   4500
      TabIndex        =   5
      ToolTipText     =   "Preview Report"
      Top             =   3615
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
      Left            =   1800
      TabIndex        =   7
      ToolTipText     =   "Close Form"
      Top             =   3600
      Width           =   1185
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   13
      Top             =   3435
      Width           =   5835
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -60
      TabIndex        =   12
      Top             =   3370
      Width           =   5955
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proforma Invoice Register"
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
      TabIndex        =   10
      Top             =   210
      Width           =   3330
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
      TabIndex        =   9
      Top             =   840
      Width           =   5850
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmRptProInvRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub cmdPrint_Click()
  Call generateReport
  If boEmptyText = True Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  CrystalReport1.Destination = crptToPrinter
  CrystalReport1.PrintReport
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdReport_Click()
  Call generateReport
  If boEmptyText = True Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  CrystalReport1.Destination = crptToWindow
  CrystalReport1.PrintReport
  Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
  Top = (Screen.Height - Height) / 2
  Left = (Screen.Width - Width) / 2
  addSupplierName cboSupplier
  optAllPI.Value = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmRptProInvRegister = Nothing
End Sub

Private Sub optAllPI_Click()
If optAllPI.Value = True Then
  txtPINo.Text = ""
  txtPINo.Enabled = False
  cboSupplier.ListIndex = -1
  cboSupplier.Enabled = False
ElseIf optPI.Value = True Then
  txtPINo.Enabled = True
  cboSupplier.ListIndex = -1
  cboSupplier.Enabled = False
Else
  txtPINo.Text = ""
  txtPINo.Enabled = False
  cboSupplier.Enabled = True
End If
End Sub
Private Sub optSupplier_Click()
  optAllPI_Click
End Sub
Private Sub optPI_Click()
  optAllPI_Click
End Sub
Private Sub txtPINo_GotFocus()
  SendKeys "{Home}+{End}"
End Sub
Private Sub txtPINo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtPINo_LostFocus()
  txtPINo.Text = Trim(UCase(txtPINo.Text))
End Sub

Sub generateReport()

boEmptyText = False

If optPI.Value = True Then
  If Len(txtPINo) = 0 Then
    MsgBox "Type PI No", vbInformation, cnstMsgInfo
    txtPINo.SetFocus
    boEmptyText = True
    Exit Sub
  End If
End If
If optSupplier.Value = True Then
  If Len(cboSupplier) = 0 Then
    MsgBox "Select Name of Supplier", vbInformation, cnstMsgInfo
    cboSupplier.SetFocus
    boEmptyText = True
    Exit Sub
  End If
End If

varNullDate = "#" + "Dec-30-1899" + "#"

Set R = New ADODB.Recordset
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

'On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

MainConn.BeginTrans

MainComm.CommandText = "DELETE * FROM Temp_Pro_Inv_Register;"
MainComm.Execute

If optAllPI.Value = True Then
  MainComm.CommandText = "INSERT INTO Temp_Pro_Inv_Register(suppName,PINo,PIDate,PICurr,PIValue,MainPF,UsesPF,itemName,PIQty,QtyUnit,compName,compAdd)" _
        & " SELECT new_Supplier.suppName,Pro_Inv_Main.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Main.PICurr,Pro_Inv_Main.PIValue,Pro_Inv_Main.MainPF," _
        & " Pro_Inv_Main.UsesPF,new_Item_Name.itemName,Sum(Pro_Inv_Sub.PIQty) AS PIQty,Pro_Inv_Sub.QtyUnit,Company_Information.compName,Company_Information.compAdd" _
        & " FROM Company_Information INNER JOIN (new_Item_Name INNER JOIN (new_Supplier INNER JOIN (Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PINo = Pro_Inv_Sub.PINo)" _
        & " ON new_Supplier.supplierID = Pro_Inv_Main.suppID) ON new_Item_Name.itemID = Pro_Inv_Sub.itemID) ON Company_Information.compID = Pro_Inv_Main.compID" _
        & " Where (((Pro_Inv_Main.compID) =" & findCompID & ")) GROUP BY new_Supplier.suppName,Pro_Inv_Main.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Main.PICurr," _
        & " Pro_Inv_Main.PIValue,Pro_Inv_Main.MainPF,Pro_Inv_Main.UsesPF,new_Item_Name.itemName,Pro_Inv_Sub.QtyUnit,Company_Information.compName,Company_Information.compAdd;"
  
ElseIf optPI.Value = True Then
  MainComm.CommandText = "INSERT INTO Temp_Pro_Inv_Register(suppName,PINo,PIDate,PICurr,PIValue,MainPF,UsesPF,itemName,PIQty,QtyUnit,compName,compAdd)" _
        & " SELECT new_Supplier.suppName,Pro_Inv_Main.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Main.PICurr,Pro_Inv_Main.PIValue,Pro_Inv_Main.MainPF," _
        & " Pro_Inv_Main.UsesPF,new_Item_Name.itemName,Sum(Pro_Inv_Sub.PIQty) AS PIQty,Pro_Inv_Sub.QtyUnit,Company_Information.compName,Company_Information.compAdd" _
        & " FROM Company_Information INNER JOIN (new_Item_Name INNER JOIN (new_Supplier INNER JOIN (Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PINo = Pro_Inv_Sub.PINo)" _
        & " ON new_Supplier.supplierID = Pro_Inv_Main.suppID) ON new_Item_Name.itemID = Pro_Inv_Sub.itemID) ON Company_Information.compID = Pro_Inv_Main.compID" _
        & " Where (((Pro_Inv_Main.compID) =" & findCompID & ")) GROUP BY new_Supplier.suppName,Pro_Inv_Main.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Main.PICurr," _
        & " Pro_Inv_Main.PIValue,Pro_Inv_Main.MainPF,Pro_Inv_Main.UsesPF,new_Item_Name.itemName,Pro_Inv_Sub.QtyUnit,Company_Information.compName," _
        & " Company_Information.compAdd HAVING (((Pro_Inv_Main.PINo)='" & txtPINo.Text & "'));"
Else
  MainComm.CommandText = "INSERT INTO Temp_Pro_Inv_Register(suppName,PINo,PIDate,PICurr,PIValue,MainPF,UsesPF,itemName,PIQty,QtyUnit,compName,compAdd)" _
        & " SELECT new_Supplier.suppName,Pro_Inv_Main.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Main.PICurr,Pro_Inv_Main.PIValue,Pro_Inv_Main.MainPF," _
        & " Pro_Inv_Main.UsesPF,new_Item_Name.itemName,Sum(Pro_Inv_Sub.PIQty) AS PIQty,Pro_Inv_Sub.QtyUnit,Company_Information.compName,Company_Information.compAdd" _
        & " FROM Company_Information INNER JOIN (new_Item_Name INNER JOIN (new_Supplier INNER JOIN (Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PINo = Pro_Inv_Sub.PINo)" _
        & " ON new_Supplier.supplierID = Pro_Inv_Main.suppID) ON new_Item_Name.itemID = Pro_Inv_Sub.itemID) ON Company_Information.compID = Pro_Inv_Main.compID" _
        & " Where (((Pro_Inv_Main.compID) =" & findCompID & ")) GROUP BY new_Supplier.suppName,Pro_Inv_Main.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Main.PICurr," _
        & " Pro_Inv_Main.PIValue,Pro_Inv_Main.MainPF,Pro_Inv_Main.UsesPF,new_Item_Name.itemName,Pro_Inv_Sub.QtyUnit,Company_Information.compName," _
        & " Company_Information.compAdd HAVING (((new_Supplier.suppName)='" & cboSupplier.Text & "'));"
End If

MainComm.Execute

strRecord = "SELECT PINo FROM Temp_Pro_Inv_Register GROUP BY PINo;"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount = 0 Then
  MsgBox "Record not Found", vbInformation, cnstMsgInfo
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  boEmptyText = True
  Exit Sub
End If
R.Close

MainComm.CommandText = "UPDATE Import_LC_Main INNER JOIN (Import_LC_Sub INNER JOIN Temp_Pro_Inv_Register ON Import_LC_Sub.PINo = Temp_Pro_Inv_Register.PINo)" _
      & " ON Import_LC_Main.LCID = Import_LC_Sub.LCID SET Temp_Pro_Inv_Register.BBLCNo = [Import_LC_Main]![LCNo],Temp_Pro_Inv_Register.MainMLCNo = [Import_LC_Main]![MLCNo];"
MainComm.Execute
MainComm.CommandText = "UPDATE Temp_Pro_Inv_Register SET SubHeader='PROFORMA INVOICE REGISTER';"
MainComm.Execute

MainConn.CommitTrans

ErrorHandler:
If Err.Number <> 0 Then
  MsgBox Err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  boEmptyText = True
  Exit Sub
End If

Screen.MousePointer = vbDefault
Set R = Nothing

CrystalReport1.ReportFileName = App.Path + "\Report\pro invoice register.rpt"

End Sub


