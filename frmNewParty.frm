VERSION 5.00
Begin VB.Form frmNewParty 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Information"
   ClientHeight    =   5460
   ClientLeft      =   39
   ClientTop       =   481
   ClientWidth     =   7956
   Icon            =   "frmNewParty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7956
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDeptUnit 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1170
      Width           =   3330
   End
   Begin VB.TextBox txtGoodsName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   660
      TabIndex        =   13
      Top             =   2610
      Width           =   6375
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Find"
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
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1170
      Width           =   540
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4890
      Width           =   960
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4890
      Width           =   930
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00CED5B0&
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
      Height          =   435
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4890
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00CED5B0&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4890
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00CED5B0&
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
      Height          =   435
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4890
      Width           =   915
   End
   Begin VB.TextBox txtSupplierAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   660
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3330
      Width           =   6405
   End
   Begin VB.TextBox txtSupplierName 
      Appearance      =   0  'Flat
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
      Left            =   660
      TabIndex        =   3
      Top             =   1890
      Width           =   6375
   End
   Begin VB.TextBox txtSupplierID 
      Appearance      =   0  'Flat
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
      Left            =   660
      TabIndex        =   0
      Top             =   1170
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Goods Name  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   660
      TabIndex        =   14
      Top             =   2400
      Width           =   1635
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   4515
      Left            =   120
      Top             =   150
      Width           =   7695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address  ( Optional )  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   690
      TabIndex        =   4
      Top             =   3090
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   660
      TabIndex        =   2
      Top             =   1650
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   0
      Left            =   690
      TabIndex        =   1
      Top             =   930
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   5
      Left            =   0
      TabIndex        =   15
      Top             =   4800
      Width           =   8115
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   150
      TabIndex        =   12
      Top             =   180
      Width           =   7635
   End
End
Attribute VB_Name = "frmNewParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbDeptUnit_Click()
    cmbDeptUnit.Tag = cmbDeptUnit.ItemData(cmbDeptUnit.ListIndex)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo x
    If MsgBox("Are you sure to delete Supplier name [" & txtSupplierName.Text & " ] ?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    MainConn.BeginTrans
        
        Mysql = " Delete FROM InventoryPJL.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryJ2L.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryPAL.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryUJL.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryNHT.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryNHTHOLD.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryNZNF.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryNZNG.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryNZND.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryPJLCA.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryPJLFG.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryNZNF.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryPWP.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryPCL.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryHUAL.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryB5MN.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Delete FROM InventoryB3OP.DBO.Supplier_Information Where SupplierID = '" & txtSupplierID.Text & "';"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
    MainConn.CommitTrans
    blank
    Call Form_Load
    Exit Sub
x:
    MainConn.RollbackTrans
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdEdit_Click()
Dim Mysql As String, strPCName As String, strUserName As String
Dim intDeptUnit As Integer
Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
'    strPCName = GetComputerName
'    strUserName = frmMaster.StatusBar1.Panels(1).Text
    intDeptUnit = IIf(cmbDeptUnit.Tag = "", 0, cmbDeptUnit.Tag)
    If txtSupplierID.Text = "" Or txtSupplierName.Text = "" Then
    Else
        If MsgBox("Are you confirm to Edit ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            
            Mysql = "UPDATE InventoryPJL.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryJ2L.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryPAL.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryUJL.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryNHT.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryNHTHOLD.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryNZNF.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
                        
            Mysql = "UPDATE InventoryNZNG.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryPJLCA.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
            Mysql = "UPDATE InventoryPJLFG.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
             Mysql = "UPDATE InventoryNZNF.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
            MainComm.CommandText = Mysql
            Set RS = MainComm.Execute
            
             Mysql = "UPDATE InventoryPWP.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
             MainComm.CommandText = Mysql
             Set RS = MainComm.Execute
            
             Mysql = "UPDATE InventoryPCL.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
             MainComm.CommandText = Mysql
             Set RS = MainComm.Execute
             
             Mysql = "UPDATE InventoryHUAL.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
             MainComm.CommandText = Mysql
             Set RS = MainComm.Execute
            
             Mysql = "UPDATE InventoryB5MN.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
             MainComm.CommandText = Mysql
             Set RS = MainComm.Execute
             
             Mysql = "UPDATE InventoryB3OP.DBO.Supplier_Information Set SupplierName = '" & txtSupplierName.Text & "', SupplierAddress = '" & txtSupplierAddress.Text & "', GoodsName = '" & txtGoodsName.Text & "',UserName = '" & strUser & "',PCName = '" & getComName & "', DeptUnitID = " & intDeptUnit & " WHERE SupplierID='" & txtSupplierID.Text & "';"
             MainComm.CommandText = Mysql
             Set RS = MainComm.Execute
            
            blank
            Call Form_Load
            txtSupplierID.SetFocus
        End If
    End If
End Sub

Private Sub cmdFind_Click()
    Set RS = New ADODB.Recordset
    strRecord = "SELECT SupplierID, SupplierName, SupplierAddress, GoodsName FROM Supplier_Information WHERE SupplierID ='" & txtSupplierID.Text & "';"
   RS.Open strRecord, MainConn, adOpenStatic
    'Set RS = MainComm.Execute
    If RS.EOF = True And RS.BOF = True Then
        MsgBox "Can't Find this ID.", vbInformation, cnstMsgInfo
        blank
        txtSupplierID.SetFocus
        Exit Sub
    End If
    txtSupplierName.Text = RS("SupplierName").Value
    txtSupplierAddress.Text = RS("SupplierAddress").Value & ""
    txtGoodsName.Text = RS("GoodsName").Value & ""
End Sub

Private Sub cmdSearch_Click()
    frmSerchForm.txtFormName.Text = "Supplier ID and Name for Edit"
    frmSerchForm.Show 1
End Sub

Private Sub cmdUpdate_Click()
Dim Mysql As String, strPCName As String, strUser As String
Dim dtSupplierCreateDate As Date
Dim intDeptUnit As Integer
    dtSupplierCreateDate = Date
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set RS = New ADODB.Recordset
'    strPCName = GetComputerName
'    strUserName = frmMaster.StatusBar1.Panels(1).Text
    intDeptUnit = IIf(cmbDeptUnit.Tag = "", 0, cmbDeptUnit.Tag)
    If txtSupplierID.Text = "" Or txtSupplierName.Text = "" Then
    Else
        strRecord = "SELECT SupplierID FROM Supplier_Information WHERE SupplierID ='" & txtSupplierID.Text & "';"
        RS.Open strRecord, MainConn, adOpenStatic, adLockOptimistic
        If RS.EOF = False And RS.BOF = False Then
            MsgBox "Duplicate ID of Supplier.", vbInformation, cnstMsgInfo
            txtSupplierID.SetFocus
            Exit Sub
        End If
        
        
        Mysql = " Insert Into InventoryPJL.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Insert Into InventoryJ2L.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute

        Mysql = " Insert Into InventoryPAL.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute

        Mysql = " Insert Into InventoryUJL.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute

        Mysql = " Insert Into InventoryNHT.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute

        Mysql = " Insert Into InventoryNHTHOLD.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Insert Into InventoryPJLCA.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Insert Into InventoryPJLFG.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Insert Into InventoryNZNF.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
        Mysql = " Insert Into InventoryNZNG.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
        Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
        MainComm.CommandText = Mysql
        MainComm.Execute
        
         Mysql = " Insert Into InventoryPWP.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
         Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
         MainComm.CommandText = Mysql
         MainComm.Execute
         
         Mysql = " Insert Into InventoryPCL.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
         Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
         MainComm.CommandText = Mysql
         MainComm.Execute
         
         Mysql = " Insert Into InventoryHUAL.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
         Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
         MainComm.CommandText = Mysql
         MainComm.Execute
        
         Mysql = " Insert Into InventoryB5MN.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
         Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
         MainComm.CommandText = Mysql
         MainComm.Execute
        
         Mysql = " Insert Into InventoryB3OP.DBO.Supplier_Information (SupplierID,SupplierCreateDate,SupplierName,SupplierAddress,GoodsName,UserName,PCName,DeptUnitID)"
         Mysql = Mysql + " Values ('" & txtSupplierID.Text & "','" & dtSupplierCreateDate & "','" & txtSupplierName.Text & "', '" & txtSupplierAddress.Text & "', '" & txtGoodsName.Text & "','" & strUser & "','" & getComName & "'," & intDeptUnit & ");"
         MainComm.CommandText = Mysql
         MainComm.Execute
         
        If Me.Caption = "New Supplier" Then
            blank
            Call Form_Load
            txtSupplierID.SetFocus
        ElseIf Me.Caption = "New Supplier (Cash)" Then
            frmStatementCash_and_Credit.txtPartyID.Text = Me.txtSupplierID.Text
            frmStatementCash_and_Credit.txtNameOfParty.Text = Me.txtSupplierName.Text
            Unload Me
        ElseIf Me.Caption = "New Supplier (Cheque)" Then
            frmStatementCheque.txtPartyID.Text = Me.txtSupplierID.Text
            frmStatementCheque.txtNameOfParty.Text = Me.txtSupplierName.Text
            frmStatementCheque.txtParticulers.Text = Me.txtGoodsName.Text
            Unload Me
        ElseIf Me.Caption = "New Supplier of Bill Details" Then
            strValue = txtSupplierName.Text
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
Dim strSupplierID As Integer
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    Me.Caption = "New Supplier"
    Call prcLoadDeptUnit(cmbDeptUnit)
    Set RS = New Recordset
        sql1 = "Select max(SupplierID) as maxid From Supplier_Information"
        RS.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
        strSupplierID = RS.Fields(0).Value + 1
    Set RS = Nothing
    If Val(strSupplierID) > 0 And Val(strSupplierID) < 10 Then
       txtSupplierID.Text = "000" & Val(strSupplierID)
    ElseIf Val(strSupplierID) > 9 And Val(strSupplierID) < 100 Then
       txtSupplierID.Text = "00" & Val(strSupplierID)
    ElseIf Val(strSupplierID) > 99 And Val(strSupplierID) < 1000 Then
       txtSupplierID.Text = "0" & Val(strSupplierID)
    ElseIf Val(strSupplierID) > 999 And Val(strSupplierID) < 10000 Then
       txtSupplierID.Text = Val(strSupplierID)
    End If
End Sub

Private Sub blank()
    txtSupplierID.Text = ""
    txtSupplierName.Text = ""
    txtGoodsName.Text = ""
    txtSupplierAddress.Text = ""
End Sub

Private Sub txtSupplierAddress_GotFocus()
    txtSupplierAddress.SelStart = 0
    txtSupplierAddress.SelLength = Len(txtSupplierAddress.Text)
End Sub

Private Sub txtSupplierID_GotFocus()
    txtSupplierID.SelStart = 0
    txtSupplierID.SelLength = Len(txtSupplierID.Text)
End Sub

Private Sub txtSupplierID_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSupplierID_LostFocus()
    If Val(txtSupplierID.Text) > 0 And Val(txtSupplierID.Text) < 10 Then
       txtSupplierID.Text = "000" & Val(txtSupplierID.Text)
    ElseIf Val(txtSupplierID.Text) > 9 And Val(txtSupplierID.Text) < 100 Then
       txtSupplierID.Text = "00" & Val(txtSupplierID.Text)
    ElseIf Val(txtSupplierID.Text) > 99 And Val(txtSupplierID.Text) < 1000 Then
       txtSupplierID.Text = "0" & Val(txtSupplierID.Text)
    ElseIf Val(txtSupplierID.Text) > 999 And Val(txtSupplierID.Text) < 10000 Then
       txtSupplierID.Text = Val(txtSupplierID.Text)
    End If
End Sub

Private Sub txtSupplierName_GotFocus()
    txtSupplierName.SelStart = 0
    txtSupplierName.SelLength = Len(txtSupplierName.Text)
End Sub
