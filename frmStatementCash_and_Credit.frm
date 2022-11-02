VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatementCash_and_Credit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8790
   ClientLeft      =   15
   ClientTop       =   315
   ClientWidth     =   10290
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmStatementCash_and_Credit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSCID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   3480
      TabIndex        =   47
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox vmrid 
      Height          =   285
      Left            =   5880
      TabIndex        =   46
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CmbMrrno 
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
      Left            =   5160
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Top             =   3390
      Width           =   2895
   End
   Begin VB.TextBox txtChalanNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8625
      TabIndex        =   5
      Top             =   2700
      Width           =   1530
   End
   Begin VB.TextBox txtChalanDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6990
      TabIndex        =   4
      Top             =   2700
      Width           =   1575
   End
   Begin VB.ComboBox cmbStatementNo 
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
      Left            =   3480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1800
      Width           =   2865
   End
   Begin MSComCtl2.DTPicker DTPToDate 
      Height          =   345
      Left            =   8370
      TabIndex        =   19
      Top             =   1200
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   100597761
      CurrentDate     =   39189
   End
   Begin MSComCtl2.DTPicker DTPFromDate 
      Height          =   345
      Left            =   6615
      TabIndex        =   18
      Top             =   1200
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   100597761
      CurrentDate     =   39189
   End
   Begin VB.CommandButton cmdNewParty 
      Height          =   345
      Left            =   6555
      Picture         =   "frmStatementCash_and_Credit.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "New Party"
      Top             =   2700
      Width           =   375
   End
   Begin VB.TextBox txtNameOfParty 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2720
      Width           =   4980
   End
   Begin VB.CommandButton cmdFindDate 
      BackColor       =   &H00C0C000&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6615
      TabIndex        =   20
      Top             =   1680
      Width           =   585
   End
   Begin VB.ComboBox cmbPreDate 
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
      Left            =   7410
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1680
      Width           =   2595
   End
   Begin VB.CommandButton cmdFindParty 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1080
      Picture         =   "frmStatementCash_and_Credit.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Search Party Name"
      Top             =   2700
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Update"
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
      Left            =   6240
      TabIndex        =   15
      Top             =   8160
      Width           =   1185
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0C000&
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   17
      Top             =   1815
      Width           =   885
   End
   Begin VB.TextBox txtCDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1800
      Width           =   2085
   End
   Begin VB.TextBox txtMRRDate 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8175
      TabIndex        =   9
      Top             =   3390
      Width           =   1965
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0C000&
      Caption         =   "&Preview"
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
      Left            =   8880
      TabIndex        =   14
      Top             =   8160
      Width           =   1185
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   3000
      TabIndex        =   11
      Top             =   4080
      Width           =   4125
   End
   Begin VB.ComboBox cmbCashAndCredit 
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
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   2115
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
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
      Left            =   9225
      TabIndex        =   12
      Top             =   4080
      Width           =   915
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   4080
      Width           =   2640
   End
   Begin VB.TextBox txtBill 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   3390
      Width           =   2625
   End
   Begin VB.TextBox txtBillDate 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2955
      TabIndex        =   7
      Top             =   3390
      Width           =   2040
   End
   Begin VB.TextBox txtPartyID 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   2700
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C000&
      Caption         =   "&Save"
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
      TabIndex        =   13
      Top             =   8160
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
      Left            =   7560
      TabIndex        =   16
      Top             =   8160
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Close"
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
      Left            =   3480
      TabIndex        =   25
      Top             =   8160
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3405
      Left            =   150
      TabIndex        =   24
      ToolTipText     =   "Select && Press Delete To Remove a Row"
      Top             =   4530
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6006
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Of Cash and Credit"
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
      TabIndex        =   45
      Top             =   120
      Width           =   5280
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   43
      Top             =   8040
      Width           =   12240
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "CHALAN NO."
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
      Height          =   165
      Left            =   8640
      TabIndex        =   42
      Top             =   2460
      Width           =   1110
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "CHALAN DATE"
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
      Height          =   180
      Left            =   7005
      TabIndex        =   41
      Top             =   2460
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      Height          =   1305
      Left            =   6435
      Top             =   840
      Width           =   3705
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date "
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
      Left            =   8370
      TabIndex        =   39
      Top             =   930
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date "
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
      Left            =   6615
      TabIndex        =   38
      Top             =   930
      Width           =   1065
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement No. "
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
      Left            =   3480
      TabIndex        =   37
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
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
      Left            =   180
      TabIndex        =   36
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MRR NO /REFERENCE "
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
      Left            =   5145
      TabIndex        =   35
      Top             =   3150
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL DATE  "
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
      Left            =   2985
      TabIndex        =   34
      Top             =   3150
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   33
      Top             =   2280
      Width           =   10275
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   32
      Top             =   7560
      Width           =   10275
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
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
      Left            =   210
      TabIndex        =   31
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL NO "
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
      Left            =   210
      TabIndex        =   30
      Top             =   3150
      Width           =   855
   End
   Begin VB.Label lblRemark 
      BackStyle       =   0  'Transparent
      Caption         =   "REMARKS"
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
      Left            =   3000
      TabIndex        =   29
      Top             =   3840
      Width           =   1245
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "MRR DATE"
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
      Left            =   8175
      TabIndex        =   28
      Top             =   3150
      Width           =   1245
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE PARTY / SUPPLIER"
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
      Left            =   210
      TabIndex        =   27
      Top             =   2430
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type "
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
      Left            =   180
      TabIndex        =   26
      Top             =   930
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Begin VB.Menu mnuTransactionPreview 
         Caption         =   "Transaction Preview"
      End
      Begin VB.Menu mnuPartyStatement 
         Caption         =   "Party Statement"
      End
   End
End
Attribute VB_Name = "frmStatementCash_and_Credit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytRowNo As Integer

Private Sub cmbCashAndCredit_Click()
'    txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
'    txtCDate.SetFocus
'Dim Stmnt As Integer
'Stmnt = MaxStatementNo(cmbCashAndCredit.Text)
''Set RS = New ADODB.Recordset
''                    strRecord = "SELECT MAX(SCID) as SCID FROM New_SCID Where ComID = " & bytCompID & " And SC_Year = " & intYear & ";"
''                    MainComm.CommandText = strRecord
''                RS.Open strRecord, MainConn, adOpenStatic
''                    If Not IsNull(RS![SCID]) Then
''                        intSCID = RS.Fields(0) + 1
''
''                    Else
''                        intSCID = 1
''                    End If
''                Set RS = Nothing
'intYear = Year(Date)
'            intSC = fncMRR(Stmnt)
'
'            txtSCID.Text = "SC" & bytCompID & intSC & "-" & intYear
End Sub

Private Sub cmbCashAndCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub CmbMRRNo_Click()
'If Len(CmbMrrno) <> 0 Then
'Call prcAddSuppMrrno(vmrid, CmbMrrno)
'Else
'MsgBox "Select Mrr No", vbExclamation
'Exit Sub
'
'End If
'
'Set R = New ADODB.Recordset
'    If Len(txtNameOfParty) <> 0 Then
'    strRecord = " SELECT MRR_Challan_Main.CompID, Company_Information.CompName, Company_Information.CompAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " FROM ((Company_Information INNER JOIN MRR_Challan_Main ON Company_Information.CompID = MRR_Challan_Main.CompID) INNER JOIN new_Supplier ON MRR_Challan_Main.suppID = new_Supplier.supplierID) INNER JOIN (new_Item_Type INNER JOIN ((MRR_Challan_Sub INNER JOIN new_Item_Name ON MRR_Challan_Sub.itemID = new_Item_Name.itemID) INNER JOIN New_PI ON MRR_Challan_Sub.PIID = New_PI.PIID) ON new_Item_Type.itemType = new_Item_Name.itemType) ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID " & _
'        " WHERE MRR_Challan_SUB.DELETEROW = 0 GROUP BY MRR_Challan_Main.CompID, Company_Information.CompName, Company_Information.CompAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " HAVING (((MRR_Challan_Main.CompID)=" & bytCompID & ") AND ((MRR_Challan_Main.MRRID)='" & vmrid.Text & "'));"
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
'        txtAmount.Text = Format(R![MRRValue], "##,##0.00")
'        txtTotal.Text = txtAmount.Text
'
'    Else
'    MsgBox "Record not found", vbExclamation
'    Exit Sub
'    End If
End Sub

Private Sub CmbMRRNo_GotFocus()
'CmbMrrno.Clear
'Call prcAddSuppMrrID(CmbMrrno, txtNameOfParty.Text)
End Sub

Private Sub cmbMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbMRRNO_LostFocus()
'If Len(CmbMrrno) <> 0 Then
'Call prcAddSuppMrrno(vmrid, CmbMrrno)
'Else
'MsgBox "Select Mrr No", vbExclamation
'Exit Sub
'
'End If
'
'Set R = New ADODB.Recordset
'    If Len(txtNameOfParty) <> 0 Then
'    strRecord = " SELECT MRR_Challan_Main.CompID, Company_Information.CompName, Company_Information.CompAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.MainConntainerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " FROM ((Company_Information INNER JOIN MRR_Challan_Main ON Company_Information.CompID = MRR_Challan_Main.CompID) INNER JOIN new_Supplier ON MRR_Challan_Main.suppID = new_Supplier.supplierID) INNER JOIN (new_Item_Type INNER JOIN ((MRR_Challan_Sub INNER JOIN new_Item_Name ON MRR_Challan_Sub.itemID = new_Item_Name.itemID) INNER JOIN New_PI ON MRR_Challan_Sub.PIID = New_PI.PIID) ON new_Item_Type.itemType = new_Item_Name.itemType) ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID " & _
'        " WHERE MRR_Challan_SUB.DELETEROW = 0 GROUP BY MRR_Challan_Main.CompID, Company_Information.CompName, Company_Information.CompAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.MainConntainerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " HAVING (((MRR_Challan_Main.CompID)=" & bytCompID & ") AND ((MRR_Challan_Main.MRRID)='" & vmrid.Text & "'));"
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
'        txtAmount.Text = Format(R![MRRValue], "##,##0.00")
'
'    Else
'    MsgBox "Record not found", vbExclamation
'    Exit Sub
'    End If
End Sub

Private Sub cmbPreDate_Click()
    txtCDate.Text = Format(CDate(cmbPreDate.Text), cnstDtFrmtI)
    If cmbPreDate.Text = "" Then
    Else
        Set R = New Recordset
            sql1 = "Select CompanyID, Tran_Type, SCNO From Statement_of_Cash_and_Credit Group By CompanyID,Tran_Type,Tran_Date,SCNo,Deleterow HAVING (((CompanyID)= " & bytCompID & ")And((Tran_Type)='" & cmbCashAndCredit.Text & "')And((Tran_Date)='" & txtCDate.Text & "')) And Deleterow=0 Order By Tran_Date"
            R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            cmbStatementNo.Clear
            Do While Not R.EOF
                cmbStatementNo.AddItem R("SCNO")
                R.MoveNext
            Loop
        Set R = Nothing
    End If
End Sub


Private Sub cmbStatementNo_Click()
  txtSCID.Text = cmbStatementNo.Text
End Sub

Private Sub cmbStatementNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo x
    If cmdAdd.Caption = "&Add" Then bytRowNo = 0
    If fncBlank(3) = True Then Exit Sub

'    Dim strItem As String
'    With flexDetails
'        If .Rows > 1 Then
'            For intI = 1 To .Rows - 1
'                .Row = intI
'                .Col = 1:   strItem = .Text
'                If strItem = txtNameOfParty.Text And intI <> bytRowNo Then
'                    MsgBox "Adding Duplicate Name of Party", vbInformation, cnstMsgInfo
'                    txtNameOfParty.SetFocus
'                    Exit Sub
'                End If
'            Next intI
'        End If
'    End With
    
    With flexDetails
        If cmdAdd.Caption = "&Add" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        Else
            .Row = bytRowNo
        End If
        .Col = 0:      .Text = flexDetails.Rows - 1                     'Count
        .Col = 1:      .Text = cmbCashAndCredit.Text                    'Transaction Type
        .Col = 2:      .Text = txtNameOfParty.Text                      'Name of the Party
        If cmdAdd.Caption = "&Add" Then
            .Col = 3:      .Text = FormatNumber(txtBill.Text, 0)        'Bill No
        Else
        End If
        .Col = 4:      .Text = txtBillDate.Text                         'Bill Date
        .Col = 5:      .Text = cmbMRRNo.Text                             'MRR No
        .Col = 6:      .Text = txtMRRDate.Text                          'MRR Date
        .Col = 7:      .Text = FormatNumber(txtAmount.Text, 2)          'Amount
        .Col = 8:      .Text = txtRemarks.Text                          'Remarks
'        .Col = 9:      .Text = txtSCID.Text                      'Statement No
        .Col = 9:     .Text = txtPartyID.Text                          'Party ID
        .Col = 10:     .Text = txtChalanDate.Text                       'Chalan Date
        .Col = 11:     .Text = txtChalanNo.Text                         'Chalan No
        
    End With
    cmdAdd.Caption = "&Add"
    Call prcBlank(1)
    txtPartyID.SetFocus
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdEdit_Click()
 Dim PC_Name1 As String, Unser_Name1 As String
 Dim intSLNO As Integer
 Dim strTranType As String
 Dim strPartyName As String, strMRRNO As String
 Dim strBillNo As String
 Dim intAmount As Double
 Dim dtBillDate As Date, dtMRRDate As Date, dtChalanDate As Date
 Dim strParticulers As String, strRemarks As String, strStatementNo As String
 Dim strSupplierID As String, strChalanNo As String
 Dim intIDNO As Double
 Dim strRecord As String
 Set MainComm = New ADODB.Command
 MainComm.ActiveConnection = MainConn

On Error GoTo x
    MainConn.BeginTrans
        'Mysql = " Delete FROM Statement_of_Cash_and_Credit Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And StatementNo = " & txtStatementNo.Text & ";"
        strRecord = "Update Statement_of_Cash_and_Credit set Deleterow=2 ,User_Name='" & strUser & "',PC_Name='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & cmbStatementNo.Text & "' And DeleteRow=0;"
        MainComm.CommandText = strRecord
        MainComm.Execute (strRecord)
        
'        Set RS = New ADODB.Recordset
'                    strRecord = "SELECT SCNO FROM Statement_of_Cash_and_Credit Where CompanyID = " & bytCompID & " and Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = " & txtStatementNo.Text & " DeleteRow=2; ;"
'                    MainComm.CommandText = strRecord
'                RS.Open strRecord, MainConn, adOpenStatic
'                    If Not IsNull(RS![SCID]) Then
'                        intSCID = RS.Fields(0).Value
'                    End If

        With flexDetails
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0: intSLNO = .Text
                .Col = 1: strTranType = .Text
                .Col = 2: strPartyName = .Text
                .Col = 3: strBillNo = .Text
                .Col = 4: dtBillDate = .Text
                .Col = 5: strMRRNO = .Text
                .Col = 6: dtMRRDate = IIf(Len(.Text) = 0, "01-Jan-2006", .Text) ''"01-Jan-2006"
                .Col = 7: intAmount = .Text
                .Col = 8: strRemarks = .Text
'                .Col = 9: strStatementNo = .Text
                .Col = 9: strSupplierID = .Text
                .Col = 10: dtChalanDate = IIf(Len(.Text) = 0, "01-Jan-2006", .Text) ''"01-Jan-2006"
                .Col = 11: strChalanNo = .Text
                
                intIDNO = fncNewID
                
                Mysql = " Insert Into Statement_of_Cash_and_Credit (Tr_No,CompanyID,Tran_Type,Tran_Date,Sl_No,NameOfTheParty,BillNo,BillDate,MRRNo,MRRDate,Amount,Remark,User_Name,PC_Name,Tran_Time,SCNo,SupplierID,Chalan_Date,Chalan_No,Entrydate,Entrytime)"
                Mysql = Mysql + " Values (" & intIDNO & "," & bytCompID & ",'" & strTranType & "','" & txtCDate.Text & "'," & intSLNO & ",'" & strPartyName & "','" & strBillNo & "','" & dtBillDate & "','" & strMRRNO & "','" & dtMRRDate & "'," & intAmount & ",'" & strRemarks & "','" & strUser & "','" & getComName & "','" & Time & "','" & cmbStatementNo & "','" & strSupplierID & "','" & dtChalanDate & "','" & strChalanNo & "','" & Date & "','" & Time & "');"
                MainComm.CommandText = Mysql
                MainComm.Execute (Mysql)
            Next inti
        End With
    MainConn.CommitTrans
'    txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
    Call prcBlank(1)
    flexDetails.Rows = 1
    
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
    MainConn.RollbackTrans
End Sub

Private Sub cmdDelete_Click()
Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
On Error GoTo x
    If MsgBox("Are you confirm to delete ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        MainConn.BeginTrans
            'Mysql = " Delete FROM Statement_of_Cash_and_Credit Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And StatementNo = " & txtStatementNo.Text & ";"
            Mysql = " Update Statement_of_Cash_and_Credit Set Deleterow=1,User_name='" & strUser & "',Pc_name='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & cmbStatementNo.Text & "' And Deleterow=0;"
            MainComm.CommandText = Mysql
            MainComm.Execute (Mysql)
            
            Call prcBlank(1)
        MainConn.CommitTrans
        MsgBox "Deletion Commplete successfully. ", vbInformation, App.Title
'        txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
        flexDetails.Rows = 1
    End If
    
    Exit Sub
x:
    MainConn.RollbackTrans
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()

Set MainComm = New ADODB.Command

    If cmbCashAndCredit.Text = "" Or txtCDate.Text = "" Or cmbStatementNo.Text = "" Then
        MsgBox "Please type Transaction Date, Statement No. and Transaction Type.", vbCritical, App.Title
    Else
        If cmdFind.Caption = "&Find" Then
        txtSCID.Text = cmbStatementNo.Text
            On Error GoTo finderror
            Set RS = New Recordset
                sql1 = "Select Tran_Type,Sl_No,NameOfTheParty,BillNo,BillDate,MRRNo,MRRDate,Amount,Remark,SCNO,SupplierID,Chalan_Date,Chalan_No From Statement_of_Cash_and_Credit Where (((CompanyID) = " & bytCompID & ")And ((Tran_Date)='" & txtCDate.Text & "') And ((Tran_Type)='" & cmbCashAndCredit.Text & "')And((SCNO)='" & cmbStatementNo.Text & "')) and Deleterow=0 Order By Tr_No"
                RS.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
                    With flexDetails
                        Dim i As Integer
                        i = 1
                        .Rows = 1
                        RS.MoveFirst
                        Do While Not RS.EOF
                            .Rows = .Rows + 1
                            .Row = i
                            .Col = 0:   .Text = RS("Sl_No").Value
                            .Col = 1:   .Text = RS("Tran_Type").Value
                            .Col = 2:   .Text = RS("NameOfTheParty").Value
                            .Col = 3:   .Text = RS("BillNo").Value
                            .Col = 4:   .Text = IIf(Format(RS("BillDate").Value, "DD-MMM-YYYY") = "01-Jan-1990", "", Format(RS("BillDate").Value, "DD-MMM-YYYY"))
                            .Col = 5:   .Text = RS("MRRNo").Value
                            .Col = 6:   .Text = IIf(Format(RS("MRRDate").Value, "DD-MMM-YYYY") = "01-Jan-1990", "", Format(RS("MRRDate").Value, "DD-MMM-YYYY"))
                            .Col = 7:   .Text = FormatNumber(RS("Amount").Value, 2)
                            .Col = 8:   .Text = RS("Remark").Value & ""
'                            .Col = 9:   .Text = RS("StatementNo").Value & ""
                            .Col = 9:  .Text = RS("SupplierID").Value & ""
                            .Col = 10:  .Text = RS("Chalan_Date").Value & ""
                            .Col = 11:  .Text = RS("Chalan_No").Value & ""
                            i = i + 1
                            RS.MoveNext
                        Loop
                    End With
            Set RS = Nothing
            cmdFind.Caption = "&Clear"
'            txtStatementNo.Locked = True
            cmdUpdate.Enabled = False
            cmdEdit.Enabled = True
finderror:
            Exit Sub
        ElseIf cmdFind.Caption = "&Clear" Then
            flexDetails.Rows = 1
            cmdFind.Caption = "&Find"
'            txtStatementNo.Locked = False
            txtSCID.Text = ""
            cmdUpdate.Enabled = True
            cmdEdit.Enabled = False
        End If
    End If
End Sub

Private Sub cmdFindDate_Click()
    If cmbCashAndCredit.Text = "" Then
        MsgBox "Please type Transaction Type.", vbCritical, App.Title 'Transaction Date and
    Else
        Set RS = New Recordset
            sql1 = "Select CompanyID, Tran_Type, Tran_Date From Statement_of_Cash_and_Credit WHERE Deleterow=0 Group By CompanyID,Tran_Type,Tran_Date HAVING (((CompanyID)= " & bytCompID & ")And((Tran_Type)='" & cmbCashAndCredit.Text & "')And((Tran_Date)>='" & DTPFromDate.Value & "')And((Tran_Date)<='" & DTPToDate.Value & "')) Order By Tran_Date"
            RS.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            cmbPreDate.Clear
            cmbStatementNo.Clear
            Do While Not RS.EOF
                cmbPreDate.AddItem RS("Tran_Date")
                RS.MoveNext
            Loop
        Set RS = Nothing
    End If
End Sub

Private Sub cmdFindParty_Click()
    frmSerchForm.txtFormName.Text = "Supplier ID and Name for Statement of Cash"
    frmSerchForm.Show 1
End Sub

Private Sub cmdNewParty_Click()
  frmNewParty.Caption = "New Supplier (Cash)"
    frmNewParty.cmdClose.Visible = False
    frmNewParty.Show 1
'frmNewSupplier.Show 1
End Sub

Private Sub cmdPreview_Click()
    If txtCDate.Text = "" Then
        MsgBox "You should provide Transaction Date.", vbCritical, App.Title
        txtCDate.SetFocus
        Exit Sub
    End If
    If Len(cmbCashAndCredit.Text) = 0 Then
        MsgBox "You should provide Transaction Type.", vbCritical, App.Title
        cmbCashAndCredit.SetFocus
        Exit Sub
    End If
    If cmbCashAndCredit.Text = "Cash" Then
    Set REPORT = Appl.OpenReport(findReportPath & "Supply Chain\rptStatementCashBill.rpt")
    strRecord = "Select * From viewStatement_Cash_Credit where CompanyID = " & bytCompID & " And Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & txtSCID.Text & "' Order By Tr_No;"
       
    ElseIf cmbCashAndCredit.Text = "Credit" Then
    Set REPORT = Appl.OpenReport(findReportPath & "Supply Chain\rptStatementCreditBill.rpt")
    strRecord = "Select * From viewStatement_Cash_Credit where CompanyID = " & bytCompID & " And Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & txtSCID.Text & "' Order By Tr_No;"

    End If
    
    
    Set R = New ADODB.Recordset
    R.Open strRecord, MainConn, adOpenStatic
    
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = REPORT
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1
    R.Close
    
    
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo x
    Dim PC_Name1 As String, Unser_Name1 As String
    Dim intSLNO As Integer
    Dim strTranType As String
    Dim strPartyName As String, strMRRNO As String
    Dim strBillNo As String
    Dim intAmount As Double
    Dim dtBillDate As Date, dtMRRDate As Date, dtChalanDate As Date
    Dim strParticulers As String, strRemarks As String, strStatementNo As String
    Dim strSupplierID As String, strChalanNo As String
    Dim intIDNO As Double
    Dim intSCID As Integer
    Dim intSC As String
    Dim intSCNo As Integer
    
    If MsgBox("Do you want to Save?", vbInformation + vbYesNo, App.Title) = vbNo Then Exit Sub
    Set MainComm = New ADODB.Command
    intYear = Year(txtCDate.Text)
    txtSCID.Text = ""
    Set RS = New ADODB.Recordset
                    strRecord = "SELECT MAX(SC_No) as SC_No FROM New_SCID Where ComID = " & bytCompID & " And SC_Year = " & intYear & " AND TransType='" & cmbCashAndCredit & "';"
                    MainComm.CommandText = strRecord
                RS.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(RS![SC_No]) Then
                        intSCID = RS.Fields(0) + 1

                    Else
                        intSCID = 1
                    End If
                Set RS = Nothing
            intSC = fncMRR(intSCID)
            txtSCID.Text = "SC" & bytCompID & intSC & "-" & intYear
            
            
            Set R = New ADODB.Recordset
                    strRecord = "SELECT MAX(SCID) as SCID FROM New_SCID;"
                    MainComm.CommandText = strRecord
                R.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(R![SCID]) Then
                        intSCNo = R.Fields(0) + 1

                    Else
                        intSCNo = 1
                    End If
                Set R = Nothing
    
        MainConn.BeginTrans
        With flexDetails
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0: intSLNO = .Text
                .Col = 1: strTranType = .Text
                .Col = 2: strPartyName = .Text
                .Col = 3: strBillNo = .Text
                .Col = 4: dtBillDate = .Text
                .Col = 5: strMRRNO = .Text
                .Col = 6: dtMRRDate = IIf(Len(.Text) = 0, "01-Jan-2006", .Text)
                .Col = 7: intAmount = .Text
                .Col = 8: strRemarks = .Text
'                .Col = 9: strStatementNo = .Text
                .Col = 9: strSupplierID = .Text
                .Col = 10: dtChalanDate = IIf(Len(.Text) = 0, "01-Jan-2006", .Text)
                .Col = 11: strChalanNo = .Text
                                               
                intIDNO = fncNewID
                                
                Mysql = " Insert Into Statement_of_Cash_and_Credit (Tr_No,CompanyID,Tran_Type,Tran_Date,Sl_No,NameOfTheParty,BillNo,BillDate,MRRNo,MRRDate,Amount,Remark,User_Name,PC_Name,Tran_Time,SCNO,SupplierID,Chalan_Date,Chalan_No,Entrydate,Entrytime)"
                Mysql = Mysql + " Values (" & intIDNO & "," & bytCompID & ",'" & strTranType & "','" & txtCDate.Text & "'," & intSLNO & ",'" & strPartyName & "','" & strBillNo & "','" & dtBillDate & "','" & strMRRNO & "','" & dtMRRDate & "'," & intAmount & ",'" & strRemarks & "','" & strUser & "','" & getComName & "','" & Time & "','" & txtSCID.Text & "','" & strSupplierID & "','" & dtChalanDate & "','" & strChalanNo & "','" & Date & "','" & Time & "');"
'                MainConn.MainCommmandText = Mysql
                MainConn.Execute (Mysql)
            
            Next inti
            
        End With
              
            Mysql = "INSERT INTO New_SCID(SCid,comID, TransType, SC_No, SC_year,SCNO,EntryDate,EntryTime)" _
                  & " VALUES (" & intSCNo & "," & bytCompID & ",'" & cmbCashAndCredit & "','" & intSC & "','" & intYear & "','" & txtSCID & "','" & Date & "','" & Time & "');"
'            MainComm.Execute
            MainConn.Execute (Mysql)
            
        MainConn.CommitTrans
    Call prcBlank(1)
'    txtStatementNo.Text = 0
    MsgBox "Data updated successfully.", vbInformation, App.Title
'    txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
    flexDetails.Rows = 1
    Exit Sub
x:
    MainConn.RollbackTrans
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub flexDetails_DblClick()
On Error Resume Next
    flexDetails_KeyDown 13, 0
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        With flexDetails
            If .RowSel = 0 Then Exit Sub
            
            .Col = 1:    cmbCashAndCredit.Text = .Text
            .Col = 2:    txtNameOfParty.Text = .Text
            .Col = 3:    txtBill.Text = .Text
            .Col = 4:    txtBillDate.Text = .Text
            .Col = 5:    cmbMRRNo.Text = .Text
            .Col = 6:    txtMRRDate.Text = .Text
            .Col = 7:    txtAmount.Text = .Text
            .Col = 8:    txtRemarks.Text = .Text
'            .Col = 9:    txtSCID.Text = .Text
            .Col = 9:   txtPartyID.Text = .Text
            .Col = 10:   txtChalanDate.Text = .Text
            .Col = 11:   txtChalanNo.Text = .Text
            
            bytRowNo = .Row
        End With
        cmdAdd.Caption = "&Replace"
        txtNameOfParty.SetFocus
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
'    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
' Call prcMakeCenter(Me)
    mnuFile.Visible = False
'    Call prcConReport(CrystalReport1)
    DTPFromDate.Value = Date
    DTPToDate.Value = Date
    cmbCashAndCredit.AddItem "Cash"
    cmbCashAndCredit.AddItem "Credit"
    'cmbCashAndCredit.AddItem "Cheque"
    Call prcGridSetting
    'cmdEdit.Enabled = False
    
    Set R = New ADODB.Recordset
    strRecord = "Select DelStatus,UpStatus from InventoryPJL.dbo.Sys_User_Name where UserName='" & strUser & "'"
    R.Open strRecord, MainConn, adOpenStatic
    If R!DelStatus = 1 Then
       cmdDelete.Enabled = True
    End If
    If R!UpStatus = 1 Then
        cmdEdit.Enabled = True
    End If
    
    
Exit Sub
x:
MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub prcBlank(Param As Integer)
'    If param = 1 Then
'        txtNameOfParty.Text = ""
'        txtBill.Text = ""
'        txtBillDate.Text = 0
'        flexDetails.Rows = 1
'    End If
   ' cmbCashAndCredit.ListIndex = -1
  '  txtPartyID.Text = ""
    txtSCID.Text = ""
   ' txtCDate.Text = ""
    txtChalanDate.Text = ""
    txtChalanNo.Text = ""
  '  txtNameOfParty.Text = ""
    'cmbStatementNo.ListIndex = -1
    txtBill.Text = ""
    txtBillDate.Text = ""
'    txtDiscount.Text = "0.00"
'    txtTotal.Text = ""
    cmbMRRNo.Text = ""
    txtMRRDate.Text = ""
    txtAmount.Text = ""
    txtRemarks.Text = ""
End Sub


Private Function fncBlank(Param As Integer) As Boolean
    fncBlank = True
        If Param = 3 Then
            If cmbCashAndCredit.Text = "" Then
                MsgBox "You should Select Cash or Credit", vbCritical, App.Title
                cmbCashAndCredit.SetFocus
                Exit Function
            End If
            If Len(Trim(txtNameOfParty.Text)) = 0 Then
                MsgBox "You should provide Name Of Party", vbCritical, App.Title
                txtNameOfParty.SetFocus
                Exit Function
            End If
            If Len(Trim(txtBill.Text)) = 0 Then
                MsgBox "You should provide Bill No", vbCritical, App.Title
                txtBill.SetFocus
                Exit Function
            End If
            If Len(Trim(txtBillDate.Text)) = 0 Then
                MsgBox "You should provide Bill Date", vbCritical, App.Title
                txtBillDate.SetFocus
                Exit Function
            End If
'            If Len(Trim(txtMRRNo.Text)) = 0 Then
'                MsgBox "You should provide MRR No.", vbCritical, App.Title
'                txtMRRNo.SetFocus
'                Exit Function
'            End If
            If Len(Trim(txtAmount.Text)) = 0 Then
                MsgBox "You should provide Amount", vbCritical, App.Title
                txtAmount.SetFocus
                Exit Function
            End If
            On Error GoTo error3
            Set R = New Recordset
                sql1 = "SELECT BillNo, SupplierID FROM Statement_of_Cash_and_Credit WHERE (((SupplierID)='" & txtPartyID.Text & "')And((BillNo)='" & txtBill.Text & "'))"
                R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
error3:
                If err.Number = 3021 Then
                    Set R = Nothing
                Else
                    txtBill.Text = R("BillNo").Value
                    txtBill.ForeColor = &HFF&
                    Set R = Nothing
                    'MsgBox "Duplicate Bill Number", 16
                    If MsgBox("Duplicate Bill Number, If you want to change then presh Yes.", vbInformation + vbYesNo, App.Title) = vbYes Then
                    
                        'Exit Function
                    Else
                        txtBill.SetFocus
                    Exit Function
                    End If
                End If
        End If
    fncBlank = False
End Function

Public Sub prcGridSetting()
    With flexDetails
        .Rows = 1
        .Cols = 12
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
'        .ColAlignment(12) = flexAlignCenterCenter
        
        .Col = 0:   .Text = "SL NO."
        .Col = 1:   .Text = "TYPE"
        .Col = 2:   .Text = "NAME OF THE PARTY/SUPP.."
        .Col = 3:   .Text = "Bill/Ref.."
        .Col = 4:   .Text = "BILL DATE"
        .Col = 5:   .Text = "MRR"
        .Col = 6:   .Text = "MRR DATE"
        .Col = 7:   .Text = "AMOUNT"
        .Col = 8:   .Text = "REMARKS"
'        .Col = 9:   .Text = "St. No"
        .Col = 9:  .Text = "P.ID"
        .Col = 10:  .Text = "Chalan Date"
        .Col = 11:  .Text = "Chalan No"
        
        .ColWidth(0) = 600:      .ColWidth(1) = 1000
        .ColWidth(2) = 2200:     .ColWidth(3) = 800
        .ColWidth(4) = 1100:     .ColWidth(5) = 1500
        .ColWidth(6) = 1100:     .ColWidth(7) = 1400
        .ColWidth(8) = 1500:     .ColWidth(9) = 800
        .ColWidth(10) = 1000:     .ColWidth(11) = 1000
'        .ColWidth(12) = 800
    End With
End Sub

Private Function fncFindData() As Boolean
'    Set R = New ADODB.Recordset
'        Mysql = "Select Count(*) AS RC From PR_Main Where PRID = " & cmbPRNo.Tag & " AND CompID = " & CompID & ";"
'        MainMainCommm.MainCommmandText = Mysql
'    Set R = MainMainCommm.Execute
'
'    If R("RC") > 0 Then
'        fncFindData = True
'    Else
'        fncFindData = False
'    End If
'
'    Set R = Nothing
End Function

''Private Function fncFindID() As Integer
''    Set R = New ADODB.Recordset
''
''    Mysql = "Select PRID from PR_Main Where PR_NO = " & cmbPRNo.Text & " and CompID = " & CompID & ";"
''    MainMainCommm.MainCommmandText = Mysql
''    Set R = MainMainCommm.Execute
''
''    If R.EOF = False And R.BOF = False Then
''        fncFindID = R("IDNO")
''    End If
''
''    Set R = Nothing
''End Function

Private Function fncNewID() As Double
    Set R = New ADODB.Recordset
        Mysql = "Select max(Tr_No) as maxID from Statement_of_Cash_and_Credit;"
        R.Open Mysql, MainConn, adOpenStatic
'        MainComm.MainCommmandText = Mysql
'    Set R = MainComm.Execute
    
    If R.EOF = False And R.BOF = False Then
        If IsNull(R("maxID")) = False Then
            fncNewID = R("maxid") + 1
        Else
            fncNewID = 1
        End If
    Else
        fncNewID = 1
    End If
    Set R = Nothing
End Function


Private Sub txtAmount_GotFocus()
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount.Text)
End Sub

Private Sub txtBill_Change()
    flexDetails.ForeColor = &H0&
    txtBill.ForeColor = &H0&
End Sub

Private Sub txtBill_GotFocus()
    txtBill.SelStart = 0
    txtBill.SelLength = Len(txtBill.Text)
End Sub

Private Sub txtBill_LostFocus()
    If txtBill.Text = "" Or txtPartyID.Text = "" Then
    Else
        With flexDetails
            For inti = 1 To .Rows - 1
                .Row = inti
                If txtBill.Text = .TextMatrix(inti, 3) And txtPartyID.Text = .TextMatrix(inti, 10) Then
                    MsgBox "Duplicate Bill Number", 16
                    .ForeColor = &HFF&
                    txtBill.ForeColor = &HFF&
                    Exit For
                Else
                End If
            Next inti
        End With
    End If
End Sub

Private Sub txtBillDate_GotFocus()
    txtBillDate.SelStart = 0
    txtBillDate.SelLength = Len(txtBillDate.Text)
End Sub

Private Sub txtBillDate_LostFocus()
    Call FormatDate(txtBillDate)
End Sub

Private Sub txtCDate_GotFocus()
    txtCDate.SelStart = 0
    txtCDate.SelLength = Len(txtCDate.Text)
End Sub

Private Sub txtCDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCDate_LostFocus()
    Call FormatDate(txtCDate)
    txtBillDate.Text = txtCDate.Text
End Sub

Private Sub txtChalanDate_GotFocus()
    txtChalanDate.SelStart = 0
    txtChalanDate.SelLength = Len(txtChalanDate.Text)
End Sub

Private Sub txtChalanDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChalanDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtChalanDate_LostFocus()
    Call FormatDate(txtChalanDate)
End Sub

Private Sub txtChalanNo_GotFocus()
    txtBill.SelStart = 0
    txtBill.SelLength = Len(txtBill.Text)
End Sub

Private Sub txtChalanNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDiscount_Change()
On Error GoTo err
If Len(txtDiscount) >= 0 And Val(txtDiscount) >= 0 Then
txtTotal.Text = Format(txtAmount - txtDiscount, "##,##0.00")
Else
MsgBox "Entry discount properly", vbCritical, "Discount"
txtDiscount.Text = "0.00"
Exit Sub
End If


Exit Sub
err:
MsgBox "Entry discount properly", vbExclamation, "Discount"
'txtDiscount.Text = "0.00"
txtTotal.Text = Format(txtAmount, "##,##0.00")
End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDiscount_LostFocus()
If Val(txtDiscount) < 0 Then
MsgBox "Entry discount properly", vbExclamation
txtDiscount.Text = "0.00"
End If
End Sub

Private Sub txtMRRDate_GotFocus()
    txtMRRDate.SelStart = 0
    txtMRRDate.SelLength = Len(txtMRRDate.Text)
End Sub

Private Sub txtMRRDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMRRDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtMRRDate_LostFocus()
    Call FormatDate(txtMRRDate)
End Sub

Private Sub txtBillDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBillDate_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtBill_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtBill_KeyPress(KeyAscii As Integer)
    Call prcNumCheck(KeyAscii)
End Sub

Private Sub txtMRRNO_GotFocus()
    txtMRRNo.SelStart = 0
    txtMRRNo.SelLength = Len(txtMRRNo.Text)
End Sub

Private Sub txtNameOfParty_GotFocus()
    txtNameOfParty.SelStart = 0
    txtNameOfParty.SelLength = Len(txtNameOfParty.Text)
End Sub

Private Sub txtNameOfParty_LostFocus()
'CmbMrrno.Clear
'Call prcAddSuppMrrID(CmbMrrno, txtNameOfParty.Text)
End Sub

Private Sub txtPartyID_GotFocus()
    txtPartyID.SelStart = 0
    txtPartyID.SelLength = Len(txtPartyID.Text)
End Sub

Private Sub txtPartyID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
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
        Set RS = New Recordset
        
        sql1 = "Select SupplierID, SupplierName, SupplierAddress From Supplier_Information Where (((SupplierID)='" & txtPartyID.Text & "'))"
            RS.Open sql1, MainConn, adOpenStatic, adCmdText, adCmdText
            
'            sql1 = "Select SupplierID, SuppName, SuppAddress From new_Supplier Where (((SupplierID)='" & txtPartyID.Text & "'))"
''            RS.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
'             RS.Open sql1, MainConn, adOpenStatic, adCmdText, adCmdText
error1:
        If err.Number = 3021 Then
            Set RS = Nothing
            MsgBox "This Is Not Valit Number", 16
            txtPartyID.SetFocus
        Else
            txtNameOfParty.Text = RS("SupplierName").Value
            Set RS = Nothing
        End If
    End If
End Sub

Private Sub txtRemarks_GotFocus()
    txtRemarks.SelStart = 0
    txtRemarks.SelLength = Len(txtRemarks.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNameOfParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
