VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSerchForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select ID"
   ClientHeight    =   6942
   ClientLeft      =   39
   ClientTop       =   377
   ClientWidth     =   8060
   ControlBox      =   0   'False
   Icon            =   "frmSerchForm.frx":0000
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6942
   ScaleWidth      =   8060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   270
      Left            =   6690
      TabIndex        =   8
      Top             =   915
      Width           =   240
   End
   Begin VB.ComboBox cmbDeptUnit 
      Height          =   315
      Left            =   3585
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   945
      Width           =   3075
   End
   Begin VB.TextBox txtFormName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   7515
   End
   Begin VB.TextBox txtPartyName 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00004080&
      Height          =   330
      Left            =   330
      TabIndex        =   1
      Top             =   945
      Width           =   3210
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
      Height          =   345
      Left            =   6930
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   930
      Width           =   915
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0C000&
      Caption         =   "Click &Here"
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
      Left            =   6780
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
      Default         =   -1  'True
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1185
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4515
      Left            =   330
      TabIndex        =   2
      Top             =   1410
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   7955
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.1509
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   -480
      TabIndex        =   9
      Top             =   6000
      Width           =   12240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   45
      Index           =   3
      Left            =   -30
      TabIndex        =   6
      Top             =   6120
      Width           =   8355
   End
End
Attribute VB_Name = "frmSerchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ab As ListItem

Private Sub cmdClose_Click()
    Set ado = Nothing
    Me.Refresh
    Unload Me
End Sub

Private Sub cmdBack_Click()
    Call ListView1_DblClick
End Sub

Private Sub cmdSearch_Click()

ListView1.ListItems.Clear
    Set RS = New Recordset
        sql1 = "Select SupplierID,SupplierName From Supplier_Information Order By SupplierName"
        RS.Open sql1, MainConn, adOpenStatic, adCmdText
        While Not RS.EOF
            Set ab = ListView1.ListItems. _
                 Add(, , CStr(RS.Fields(0)))
            If Not IsNull(ado) Then
               ab.SubItems(1) = CStr(RS.Fields(1))
            End If
            RS.MoveNext
        Wend
    Set RS = Nothing
    
'    ListView1.ListItems.Clear
'    Set RS = New Recordset
'        sql1 = "Select SupplierID,SuppName From new_Supplier Order By SupplierID"
''        RS.Open sql1, con, adOpenDynamic, adLockPessimistic, adCmdText
'        RS.Open sql1, MainConn, adOpenStatic, adCmdText
'        While Not RS.EOF
'            Set ab = ListView1.ListItems. _
'                 Add(, , CStr(RS.Fields(0)))
'            If Not IsNull(ado) Then
'               ab.SubItems(1) = CStr(RS.Fields(1))
'            End If
'            RS.MoveNext
'        Wend
'    Set RS = Nothing
End Sub

Private Sub Form_Load()
    Call prcLoadDeptUnit(cmbDeptUnit)
    ListView1.ColumnHeaders. _
       Add , , "Party ID", ListView1.Width / 4
    ListView1.ColumnHeaders. _
       Add , , "Party Name", ListView1.Width
    ListView1.View = lvwReport
End Sub

Private Sub ListView1_DblClick()
    If txtFormName.Text = "Supplier ID and Name for Statement of Cash" Then
        frmStatementCash_and_Credit.txtPartyID.Text = ListView1.SelectedItem
        Unload Me
        frmStatementCash_and_Credit.txtPartyID.SetFocus
    ElseIf txtFormName.Text = "Supplier ID and Name for Statement of Cheque" Then
        frmStatementCheque.txtPartyID.Text = ListView1.SelectedItem
        Unload Me
        frmStatementCheque.txtPartyID.SetFocus
    ElseIf txtFormName.Text = "Supplier ID and Name for Edit" Then
        frmNewParty.txtSupplierID.Text = ListView1.SelectedItem
        Unload Me
        frmNewParty.txtSupplierID.SetFocus
    ElseIf txtFormName.Text = "Party ID and Name for PartyStatement" Then
        frmrptPartyStatement.txtPartyID.Text = ListView1.SelectedItem
        Unload Me
        frmrptPartyStatement.txtPartyID.SetFocus
    End If
End Sub

Private Sub txtPartyName_Change()
    ListView1.ListItems.Clear
    Set RS = New Recordset
        sql1 = "Select SupplierID, SupplierName From Supplier_Information Where (((SupplierName) Like ('" + txtPartyName.Text + "%')))Order by SupplierID"
        RS.Open sql1, MainConn, adOpenStatic, adCmdText
        While Not RS.EOF
            Set ab = ListView1.ListItems. _
                 Add(, , CStr(RS.Fields(0)))
            If Not IsNull(ado) Then
               ab.SubItems(1) = CStr(RS.Fields(1))
            End If
            RS.MoveNext
        Wend
    Set RS = Nothing
End Sub
