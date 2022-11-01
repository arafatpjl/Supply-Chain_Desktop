VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatementCheque 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
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
   Icon            =   "frmStatementCheque.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStatementCheque.frx":000C
   ScaleHeight     =   9030
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3240
      TabIndex        =   50
      Top             =   1290
      Width           =   2895
   End
   Begin VB.TextBox vmrid 
      Height          =   285
      Left            =   6240
      TabIndex        =   49
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmbMRRNO 
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
      Left            =   3720
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Top             =   3435
      Width           =   2895
   End
   Begin VB.TextBox txtChalanNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   8790
      TabIndex        =   5
      Top             =   2745
      Width           =   1530
   End
   Begin VB.TextBox txtChalanDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   7155
      TabIndex        =   4
      Top             =   2745
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
      Left            =   3240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdNewParty 
      Height          =   345
      Left            =   6690
      Picture         =   "frmStatementCheque.frx":04C2
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "New Party"
      Top             =   2745
      Width           =   375
   End
   Begin VB.TextBox txtNameOfParty 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2745
      Width           =   5130
   End
   Begin VB.CommandButton cmdFindDate 
      BackColor       =   &H00CED5B0&
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
      Left            =   6885
      TabIndex        =   19
      Top             =   1785
      Width           =   615
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
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1800
      Width           =   2325
   End
   Begin VB.TextBox txtParticulers 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      TabIndex        =   11
      Top             =   4110
      Width           =   3555
   End
   Begin VB.CommandButton cmdFindParty 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1050
      Picture         =   "frmStatementCheque.frx":09F4
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Search Party Name"
      Top             =   2745
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00CED5B0&
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
      Left            =   6315
      TabIndex        =   15
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00CED5B0&
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
      Height          =   330
      Left            =   2190
      TabIndex        =   17
      ToolTipText     =   "Search Record"
      Top             =   1950
      Width           =   915
   End
   Begin VB.TextBox txtCDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   1950
      Width           =   1995
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
      Left            =   6645
      TabIndex        =   9
      Top             =   3435
      Width           =   1695
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
      Left            =   3750
      TabIndex        =   12
      Top             =   4110
      Width           =   4020
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
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1290
      Width           =   1995
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00CED5B0&
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
      Height          =   360
      Left            =   9165
      TabIndex        =   13
      Top             =   4080
      Width           =   1140
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
      Left            =   8430
      TabIndex        =   10
      Top             =   3435
      Width           =   1890
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
      Left            =   90
      TabIndex        =   6
      Top             =   3435
      Width           =   1845
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
      Left            =   2040
      TabIndex        =   7
      Top             =   3435
      Width           =   1560
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
      Left            =   120
      TabIndex        =   2
      Top             =   2745
      Width           =   915
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00CED5B0&
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
      TabIndex        =   14
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00CED5B0&
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
      Left            =   7605
      TabIndex        =   16
      Top             =   8370
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00CED5B0&
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
      Left            =   143
      TabIndex        =   25
      Top             =   8370
      Width           =   1185
   End
   Begin MSFlexGridLib.MSFlexGrid flexDetails 
      Height          =   3495
      Left            =   60
      TabIndex        =   24
      ToolTipText     =   "Select && Press Delete To Remove a Row"
      Top             =   4560
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPFromDate 
      Height          =   345
      Left            =   6885
      TabIndex        =   22
      Top             =   1320
      Width           =   1605
      _ExtentX        =   2831
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
      Format          =   96993281
      CurrentDate     =   39189
   End
   Begin MSComCtl2.DTPicker DTPToDate 
      Height          =   345
      Left            =   8550
      TabIndex        =   18
      Top             =   1320
      Width           =   1605
      _ExtentX        =   2831
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
      Format          =   96993281
      CurrentDate     =   39189
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Report"
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
      Height          =   375
      Left            =   8910
      TabIndex        =   39
      Top             =   8400
      Width           =   1410
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Of Cheque  And Cash  Cheque"
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
      Index           =   4
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   5280
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   -240
      TabIndex        =   45
      Top             =   8160
      Width           =   10800
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000009&
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
      Left            =   8805
      TabIndex        =   44
      Top             =   2505
      Width           =   1110
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000009&
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
      Left            =   7170
      TabIndex        =   43
      Top             =   2505
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   6750
      Top             =   960
      Width           =   3585
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   8550
      TabIndex        =   41
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Left            =   6885
      TabIndex        =   40
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Statement No."
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
      Left            =   3360
      TabIndex        =   38
      Top             =   1020
      Width           =   1245
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PARTICULARS OF PURPOSE"
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
      Left            =   90
      TabIndex        =   37
      Top             =   3840
      Width           =   2595
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
      Left            =   90
      TabIndex        =   36
      Top             =   1710
      Width           =   1665
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MRR NO / Reference"
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
      Left            =   3750
      TabIndex        =   35
      Top             =   3195
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL DATE"
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
      Left            =   2130
      TabIndex        =   34
      Top             =   3195
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   33
      Top             =   2355
      Width           =   10425
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   32
      Top             =   8160
      Width           =   10455
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
      Left            =   8430
      TabIndex        =   31
      Top             =   3195
      Width           =   915
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL NO"
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
      Left            =   90
      TabIndex        =   30
      Top             =   3195
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
      Left            =   3750
      TabIndex        =   29
      Top             =   3870
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
      Left            =   6645
      TabIndex        =   28
      Top             =   3195
      Width           =   1245
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE PARTY / SUPPLIER "
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
      Left            =   120
      TabIndex        =   27
      Top             =   2475
      Width           =   2985
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
      Left            =   90
      TabIndex        =   26
      Top             =   1020
      Width           =   1845
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   12240
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
      Left            =   480
      TabIndex        =   46
      Top             =   120
      Width           =   5280
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Begin VB.Menu mnuDetailsPreview 
         Caption         =   "Details Preview"
      End
      Begin VB.Menu mnuSummaryPreview 
         Caption         =   "Summary Preview"
      End
      Begin VB.Menu mnuPartyStatement 
         Caption         =   "Party Statement"
      End
   End
End
Attribute VB_Name = "frmStatementCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytRowNo As Integer

Private Sub cmbCashAndCredit_Click()
'    txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
'    txtCDate.SetFocus
End Sub

Private Sub cmbCashAndCredit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub CmbMRRNo_Click()
'If Len(cmbMRRNO) <> 0 Then
'Call prcAddSuppMrrno(vmrid, cmbMRRNO)
'Else
'MsgBox "Select Mrr No", vbExclamation
'Exit Sub
'
'End If
'
'Set R = New ADODB.Recordset
'    If Len(txtNameOfParty) <> 0 Then
'    strRecord = " SELECT MRR_Challan_Main.compID, Company_Information.compName, Company_Information.compAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " FROM ((Company_Information INNER JOIN MRR_Challan_Main ON Company_Information.compID = MRR_Challan_Main.compID) INNER JOIN new_Supplier ON MRR_Challan_Main.suppID = new_Supplier.supplierID) INNER JOIN (new_Item_Type INNER JOIN ((MRR_Challan_Sub INNER JOIN new_Item_Name ON MRR_Challan_Sub.itemID = new_Item_Name.itemID) INNER JOIN New_PI ON MRR_Challan_Sub.PIID = New_PI.PIID) ON new_Item_Type.itemType = new_Item_Name.itemType) ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID " & _
'        " WHERE MRR_Challan_SUB.DELETEROW = 0 GROUP BY MRR_Challan_Main.compID, Company_Information.compName, Company_Information.compAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
'        " HAVING (((MRR_Challan_Main.compID)=" & bytCompID & ") AND ((MRR_Challan_Main.MRRID)='" & vmrid.Text & "'));"
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
'cmbMRRNO.Clear
'Call prcAddSuppMrrID(cmbMRRNO, txtNameOfParty.Text)
End Sub

Private Sub cmbMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbPreDate_Click()
    txtCDate.Text = Format(CDate(cmbPreDate.Text), cnstDtFrmtI)
    If cmbPreDate.Text = "" Then
    Else
        Set R = New Recordset
            sql1 = "Select CompanyID, Tran_Type, Tran_Date, SCNo From Statement_of_Cash_and_Credit Group By CompanyID,Tran_Type,Tran_Date,SCNo,Deleterow HAVING (((CompanyID)= " & bytCompID & ")And((Tran_Type)='" & cmbCashAndCredit.Text & "')And((Tran_Date)='" & Format(txtCDate.Text, "dd-mmm-yy") & "'))And Deleterow=0 Order By Tran_Date"
            R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            cmbStatementNo.Clear
            Do While Not R.EOF
                cmbStatementNo.AddItem R("SCNo")
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
    
    With flexDetails
        If cmdAdd.Caption = "&Add" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        Else
            .Row = bytRowNo
        End If
        .Col = 0:      .Text = flexDetails.Rows - 1                   'Count
        .Col = 1:      .Text = cmbCashAndCredit.Text                  'Transaction Type
        .Col = 2:      .Text = txtNameOfParty.Text                    'Name of the Party
        If cmdAdd.Caption = "&Add" Then
            .Col = 3:      .Text = FormatNumber(txtBill.Text, 0)      'Bill No
        Else
        End If
        .Col = 4:      .Text = txtBillDate.Text                       'Bill Date
        .Col = 5:      .Text = cmbMRRNo.Text                          'MRR No
        .Col = 6:      .Text = txtMRRDate.Text                        'MRR Date
        .Col = 7:      .Text = FormatNumber(txtAmount.Text, 2)        'Amount
        .Col = 8:      .Text = txtParticulers.Text                    'Particulers
        .Col = 9:      .Text = txtRemarks.Text                        'Remarks
'        .Col = 10:     .Text = txtStatementNo.Text                    'Statement No
        .Col = 10:     .Text = txtPartyID.Text                        'Party ID
        .Col = 11:     .Text = txtChalanDate.Text                     'Chalan Date
        .Col = 12:     .Text = txtChalanNo.Text                       'Chalan No
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
 Dim strPartyID As String, strChalanNo As String
 Dim intIDNO As Double
 Set MainComm = New ADODB.Command
 MainComm.ActiveConnection = MainConn

    On Error GoTo x
    MainConn.BeginTrans
        'Mysql = " Delete FROM Statement_of_Cash_and_Credit Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And StatementNo = " & txtStatementNo.Text & ";"
        Mysql = " Update Statement_of_Cash_and_Credit Set Deleterow=2,User_name='" & strUser & "',Pc_name='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & cmbStatementNo.Text & "' And Deleterow=0;"
        MainComm.CommandText = Mysql
        MainComm.Execute (Mysql)
        
'        Set RS = New ADODB.Recordset
'                    strRecord = "SELECT SCID FROM Statement_of_Cash_and_Credit Where CompanyID = " & bytCompID & " and Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And StatementNo = '" & txtSCID.Text & "' And DeleteRow=2; ;"
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
                .Col = 6: dtMRRDate = IIf(Len(.Text) = 0, "01-Jan-1900", .Text) ''"01-Jan-2006"
                .Col = 7: intAmount = .Text
                .Col = 8: strParticulers = .Text
                .Col = 9: strRemarks = .Text
'                .Col = 10: strStatementNo = .Text
                .Col = 10: strPartyID = .Text
                .Col = 11: dtChalanDate = IIf(Len(.Text) = 0, "01-Jan-1900", .Text) ''"01-Jan-2006"
                .Col = 12: strChalanNo = .Text
                intIDNO = fncNewID
                strRecord = " Insert Into Statement_of_Cash_and_Credit (Tr_No,CompanyID,Tran_Type,Tran_Date,Sl_No,NameOfTheParty,BillNo,BillDate,MRRNo,MRRDate,Amount,ParticulersOfPurchase,Remark,User_Name,PC_Name,Tran_Time,SCNo,SupplierID,Chalan_Date,Chalan_No,Entrydate,Entrytime)"
                strRecord = strRecord + " Values (" & intIDNO & "," & bytCompID & ",'" & strTranType & "','" & txtCDate.Text & "'," & intSLNO & ",'" & strPartyName & "','" & strBillNo & "','" & dtBillDate & "','" & strMRRNO & "','" & dtMRRDate & "'," & intAmount & ",'" & strParticulers & "','" & strRemarks & "','" & strUser & "','" & getComName & "','" & Time & "','" & cmbStatementNo & "','" & strPartyID & "','" & dtChalanDate & "','" & strChalanNo & "','" & Date & "','" & Time & "');"
                MainComm.CommandText = strRecord
                MainComm.Execute (strRecord)
            Next inti
        End With
    MainConn.CommitTrans
'    txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
    flexDetails.Rows = 1
    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
    con.RollbackTrans
End Sub

Private Sub cmdDelete_Click()

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn
On Error GoTo x
    If MsgBox("Are you confirm to delete ?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        MainConn.BeginTrans
            'Mysql = " Delete FROM Statement_of_Cash_and_Credit Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And StatementNo = " & txtStatementNo.Text & ";"
            Mysql = " Update Statement_of_Cash_and_Credit set Deleterow=1,User_name='" & strUser & "',pc_name='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' Where Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & cmbStatementNo.Text & "' and Deleterow=0;"
            MainComm.CommandText = Mysql
            MainComm.Execute
            
            Call prcBlank(1)
            MainConn.CommitTrans
            MsgBox "Deletion complete successfully. ", vbInformation, App.Title
'            txtStatementNo.Text = MaxStatementNo(cmbCashAndCredit.Text)
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
'            On Error GoTo finderror
            Set R = New Recordset
                sql1 = "Select Tran_Type,Sl_No,NameOfTheParty,BillNo,BillDate,MRRNo,MRRDate,Amount,ParticulersOfPurchase,Remark,SCNo,SupplierID,Chalan_Date,Chalan_No From Statement_of_Cash_and_Credit Where (((CompanyID) = " & bytCompID & ")And ((Tran_Date)='" & txtCDate.Text & "')And((Tran_Type)='" & cmbCashAndCredit.Text & "')And ((SCNo)= '" & cmbStatementNo.Text & "'))And Deleterow=0 Order By Tr_No"
                R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
                
                    With flexDetails
                        Dim i As Integer
                        i = 1
                        .Rows = 1
                        R.MoveFirst
                        Do While Not R.EOF
                            .Rows = .Rows + 1
                            .Row = i
                            .Col = 0: .Text = R("Sl_No").Value
                            .Col = 1: .Text = R("Tran_Type").Value
                            .Col = 2: .Text = R("NameOfTheParty").Value
                            .Col = 3: .Text = R("BillNo").Value
                            .Col = 4: .Text = IIf(Format(R("BillDate").Value, "DD-MMM-YYYY") = "01-Jan-1990", "", Format(R("BillDate").Value, "DD-MMM-YYYY"))
                            .Col = 5: .Text = R("MRRNo").Value
                            .Col = 6: .Text = IIf(Format(R("MRRDate").Value, "DD-MMM-YYYY") = "01-Jan-1990", "", Format(R("MRRDate").Value, "DD-MMM-YYYY"))
                            .Col = 7: .Text = FormatNumber(R("Amount").Value, 2)
                            .Col = 8: .Text = R("ParticulersOfPurchase").Value & ""
                            .Col = 9: .Text = R("Remark").Value & ""
'                            .Col = 10: .Text = R("StatementNo").Value & ""
                            .Col = 10: .Text = R("SupplierID").Value & ""
                            .Col = 11: .Text = R("Chalan_Date").Value & ""
                            .Col = 12: .Text = R("Chalan_No").Value & ""
                            i = i + 1
                            R.MoveNext
                        Loop
                    End With
            Set R = Nothing
            cmdFind.Caption = "&Clear"
'            txtStatementNo.Locked = True
            cmdUpdate.Enabled = False
            cmdEdit.Enabled = True
'finderror:
'            Exit Sub
        ElseIf cmdFind.Caption = "&Clear" Then
            flexDetails.Rows = 1
            cmdFind.Caption = "&Find"
            txtSCID.Text = ""
'            txtStatementNo.Locked = False
            cmdUpdate.Enabled = True
            cmdEdit.Enabled = False
        End If
    End If
End Sub

Private Sub cmdFindDate_Click()
    If cmbCashAndCredit.Text = "" Then
        MsgBox "Please type Transaction Type.", vbCritical, App.Title
    Else
        Set R = New Recordset
            sql1 = "Select CompanyID, Tran_Type, Tran_Date From Statement_of_Cash_and_Credit where Deleterow=0 Group By CompanyID,Tran_Type,Tran_Date HAVING (((CompanyID)= " & bytCompID & ")And((Tran_Type)='" & cmbCashAndCredit.Text & "')And((Tran_Date)>='" & DTPFromDate.Value & "')And((Tran_Date)<='" & DTPToDate.Value & "')) Order By Tran_Date"
            R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            cmbPreDate.Clear
            cmbStatementNo.Clear
            Do While Not R.EOF
                cmbPreDate.AddItem R("Tran_Date")
                R.MoveNext
            Loop
        Set R = Nothing
    End If
End Sub

Private Sub cmdFindParty_Click()
    frmSerchForm.txtFormName.Text = "Supplier ID and Name for Statement of Cheque"
    frmSerchForm.Show 1
End Sub

Private Sub cmdNewParty_Click()
    frmNewParty.Caption = "New Supplier (Cheque)"
    frmNewParty.cmdClose.Visible = False
    frmNewParty.Show 1
End Sub

Private Sub cmdUpdate_Click()

Dim PC_Name1 As String, Unser_Name1 As String
Dim intSLNO As Integer
Dim strTranType As String
Dim strPartyName As String, strMRRNO As String
Dim strBillNo As String
Dim intAmount As Double
Dim dtBillDate As Date, dtMRRDate As Date, dtChalanDate As Date
Dim strParticulers As String, strRemarks As String, strStatementNo As String
Dim strPartyID As String, strChalanNo As String
Dim intIDNO As Double
Dim intSCID As Integer
Dim intSC As String
Dim intSCNo As Integer
    
    
    On Error GoTo x
    If MsgBox("Do you want to Save?", vbInformation + vbYesNo, App.Title) = vbNo Then Exit Sub
        Set MainComm = New ADODB.Command
        MainConn.BeginTrans
        
        intYear = Year(txtCDate.Text)
             
        Set R = New ADODB.Recordset
                    strRecord = "SELECT MAX(SC_NO) as SC_NO FROM New_SCID Where ComID = " & bytCompID & " And SC_Year = " & intYear & " AND TransType='" & cmbCashAndCredit & "';"
                    MainComm.CommandText = strRecord
                R.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(R![SC_No]) Then
                        intSCID = R.Fields(0) + 1
                    Else
                        intSCID = 1
                    End If
                Set R = Nothing
            intSC = fncMRR(intSCID)
            txtSCID.Text = "SC" & bytCompID & intSC & "-" & intYear
            
            Set RS = New ADODB.Recordset
                    strRecord = "SELECT MAX(SCID) as SCID FROM New_SCID;"
                    MainComm.CommandText = strRecord
                RS.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(RS![SCID]) Then
                        intSCNo = RS.Fields(0) + 1
                        
                    Else
                        intSCNo = 1
                    End If
                Set RS = Nothing
            
        With flexDetails
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0:  intSLNO = .Text
                .Col = 1:  strTranType = .Text
                .Col = 2:  strPartyName = .Text
                .Col = 3:  strBillNo = .Text
                .Col = 4:  dtBillDate = .Text
                .Col = 5:  strMRRNO = .Text
                .Col = 6:  dtMRRDate = IIf(Len(.Text) = 0, "01-Jan-2006", .Text)
                .Col = 7:  intAmount = .Text
                .Col = 8:  strParticulers = .Text
                .Col = 9:  strRemarks = .Text
'                .Col = 10: strStatementNo = .Text
                .Col = 10: strPartyID = .Text
                .Col = 11: dtChalanDate = IIf(Len(.Text) = 0, "01-Jan-2006", .Text)
                .Col = 12: strChalanNo = .Text
                intIDNO = fncNewID
                
                Mysql = " Insert Into Statement_of_Cash_and_Credit (Tr_No,CompanyID,Tran_Type,Tran_Date,Sl_No,NameOfTheParty,BillNo,BillDate,MRRNo,MRRDate,Amount,ParticulersOfPurchase,Remark,User_Name,PC_Name,Tran_Time,SCNo,SupplierID,Chalan_Date,Chalan_No,Entrydate,Entrytime)"
                Mysql = Mysql + " Values (" & intIDNO & "," & bytCompID & ",'" & strTranType & "','" & txtCDate.Text & "'," & intSLNO & ",'" & strPartyName & "','" & strBillNo & "','" & dtBillDate & "','" & strMRRNO & "','" & dtMRRDate & "'," & intAmount & ",'" & strParticulers & "','" & strRemarks & "','" & strUser & "','" & getComName & "','" & Time & "','" & txtSCID & "','" & strPartyID & "','" & dtChalanDate & "','" & strChalanNo & "','" & Date & "','" & Time & "');"
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
    MsgBox "Data saved successfully", vbInformation, App.Title
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
            '.Col = 0:    lblRemark.Caption = .Text
            .Col = 1:    cmbCashAndCredit.Text = .Text
            .Col = 2:    txtNameOfParty.Text = .Text
            .Col = 3:    txtBill.Text = .Text
            .Col = 4:    txtBillDate.Text = .Text
            .Col = 5:    cmbMRRNo.Text = .Text
            .Col = 6:    txtMRRDate.Text = .Text
            .Col = 7:    txtAmount.Text = .Text
            .Col = 8:    txtParticulers.Text = .Text
            .Col = 9:    txtRemarks.Text = .Text
'            .Col = 10:   txtStatementNo.Text = .Text
            .Col = 10:   txtPartyID.Text = .Text
            .Col = 11:   txtChalanDate.Text = .Text
            .Col = 12:   txtChalanNo.Text = .Text
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
        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
'        Call prcConReport(CrystalReport1)
        DTPFromDate.Value = Date
        DTPToDate.Value = Date
        mnuFile.Visible = False
        cmbCashAndCredit.AddItem "Cheque"
        'cmbCashAndCredit.AddItem "Cheque/Cash"
        cmbCashAndCredit.AddItem "Cash Cheque"
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
    txtChalanDate.Text = ""
    txtChalanNo.Text = ""
    txtNameOfParty.Text = ""
    txtBill.Text = ""
    txtBillDate.Text = ""
'    txtDiscount.Text = "0.00"
'    txtTotal.Text = ""
    cmbMRRNo.Text = ""
    txtMRRDate.Text = ""
    txtAmount.Text = ""
    txtParticulers.Text = ""
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
            On Error GoTo error2
            Set R = New Recordset
                sql1 = "SELECT BillNo, SupplierID FROM Statement_of_Cash_and_Credit WHERE (((SupplierID)='" & txtPartyID.Text & "')And((BillNo)='" & txtBill.Text & "'))"
                R.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
error2:
                If err.Number = 3021 Then
                    Set R = Nothing
                Else
                    txtBill.Text = R("BillNo").Value
                    txtBill.ForeColor = &HFF&
                    Set R = Nothing
                    'MsgBox "Duplicate Bill Number", 16
                    If MsgBox("Duplicate Bill Number, If you want to change then press Yes.", vbInformation + vbYesNo, App.Title) = vbYes Then
                    
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
        .Cols = 13
        
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
        .ColAlignment(12) = flexAlignCenterCenter
'        .ColAlignment(13) = flexAlignCenterCenter
        
        .Col = 0:   .Text = "SL NO."
        .Col = 1:   .Text = "TYPE"
        .Col = 2:   .Text = "NAME OF THE PARTY/SUPP.."
        .Col = 3:   .Text = "Bill/Ref.."
        .Col = 4:   .Text = "BILL DATE"
        .Col = 5:   .Text = "MRR"
        .Col = 6:   .Text = "MRR DATE"
        .Col = 7:   .Text = "AMOUNT"
        .Col = 8:   .Text = "Particulers of Purchase"
        .Col = 9:   .Text = "REMARKS"
'        .Col = 10:   .Text = "St. No"
        .Col = 10:   .Text = "P. ID"
        .Col = 11:   .Text = "Chalan Date"
        .Col = 12:   .Text = "Chalan No"
        
        .ColWidth(0) = 600:      .ColWidth(1) = 1000
        .ColWidth(2) = 2200:     .ColWidth(3) = 800
        .ColWidth(4) = 1100:     .ColWidth(5) = 1300
        .ColWidth(6) = 1100:     .ColWidth(7) = 1400
        .ColWidth(8) = 2500:     .ColWidth(9) = 1500
        .ColWidth(10) = 800:     .ColWidth(11) = 1000
        .ColWidth(12) = 1000:
'        .ColWidth(13) = 800
    End With
End Sub

Private Function fncFindData() As Boolean
'    Set R = New ADODB.Recordset
'        Mysql = "Select Count(*) AS RC From PR_Main Where PRID = " & cmbPRNo.Tag & " AND Compid = " & CompID & ";"
'        MainComm.CommandText = Mysql
'    Set R = MainComm.Execute
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
''    Mysql = "Select PRID from PR_Main Where PR_NO = " & cmbPRNo.Text & " and Compid = " & CompID & ";"
''    MainComm.CommandText = Mysql
''    Set R = MainComm.Execute
''
''    If R.EOF = False And R.BOF = False Then
''        fncFindID = R("IDNO")
''    End If
''
''    Set R = Nothing
''End Function

Private Function fncNewID() As Double
    Set R = New ADODB.Recordset
        Mysql = "Select max(Tr_No) as maxID from Statement_of_Cash_and_Credit Where Companyid = " & bytCompID & ";"
       R.Open Mysql, MainConn, adOpenStatic

    
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

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PopupMenu mnuFile, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub mnuDetailsPreview_Click()
    
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
'    With CrystalReport1
'        .ReportFileName = findReportPath & "\Supply Chain\rptStatementChequeDetails.rpt"
'        .PrintFileName = findReportPath & "\Supply Chain\rptStatementChequeDetails.rpt"
'
'        .SQLQuery = "Select * from viewStatement_Cash_Credit where CompanyID = " & bytCompID & " and Tran_Type = '" & cmbCashAndCredit.Text & "' and Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & txtSCID.Text & "' Order By Tr_No;"
'        .PrintReport
'    End With
    
    Set REPORT = Appl.OpenReport(findReportPath & "Supply Chain\rptStatementChequeDetails.rpt")
    strRecord = "Select * from viewStatement_Cash_Credit where CompanyID = " & bytCompID & " and Tran_Type = '" & cmbCashAndCredit.Text & "' and Tran_Date = '" & txtCDate.Text & "' And SCNo = '" & txtSCID.Text & "' Order By Tr_No;"
           
    Set R = New ADODB.Recordset
    R.Open strRecord, MainConn, adOpenStatic
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = REPORT
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    R.Close
       
End Sub

Private Sub mnuPartyStatement_Click()
    If txtPartyID.Text = "" Then
        MsgBox "You should provide Supply ID", vbCritical, App.Title
        txtPartyID.SetFocus
        Exit Sub
    End If

'    With CrystalReport1
'        .Formulas(0) = "":  .Formulas(1) = "":  .Formulas(2) = ""
'
'        .ReportFileName = findReportPath & "\Supply Chain\rptPartyStatement.rpt"
'        .PrintFileName = findReportPath & "\Supply Chain\rptPartyStatement.rpt"
'        If cmbCashAndCredit.Text = "" Then
'            .SQLQuery = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPFromDate.Value & "' And Tran_Date <= '" & DTPToDate.Value & "' Order By Tr_No;"
'            .Formulas(2) = "StatementType = 'All Transaction'"
'        Else
'            .SQLQuery = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date >= '" & DTPFromDate.Value & "' And Tran_Date <= '" & DTPToDate.Value & "' Order By Tr_No;"
'        End If
'        .Formulas(0) = "FromDate = '" & Format(DTPFromDate.Value, "DD-MMM-YYYY") & "'"
'        .Formulas(1) = "ToDate = '" & Format(DTPToDate.Value, "DD-MMM-YYYY") & "'"
'        .PrintReport
'
'        .Formulas(0) = "": .Formulas(1) = "":  .Formulas(2) = ""
'    End With


        
        Set REPORT = Appl.OpenReport(findReportPath & "Supply Chain\rptPartyStatement.rpt")
        
        If Len(txtSCID.Text) = 0 Then
        If cmbCashAndCredit.Text = "" Then
        strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPFromDate.Value & "' And Tran_Date <= '" & DTPToDate.Value & "' Order By Tr_No;"
            '.SQLQuery = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Date >= '" & DTPFromDate.Value & "' And Tran_Date <= '" & DTPToDate.Value & "' Order By Tr_No;"
            '.Formulas(2) = "StatementType = 'All Transaction'"
        Else
            strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbCashAndCredit.Text & "' And Tran_Date >= '" & DTPFromDate.Value & "' And Tran_Date <= '" & DTPToDate.Value & "' Order By Tr_No;"
        End If
        Else
        strRecord = "Select * From viewStatement_Cash_Credit Where CompanyID = " & bytCompID & " And SupplierID = '" & txtPartyID.Text & "' And Tran_Type = '" & cmbCashAndCredit.Text & "' And SCNO = '" & txtSCID.Text & "' Order By Tr_No;"
        End If
        '.Formulas(0) = "FromDate = '" & Format(DTPFromDate.Value, "DD-MMM-YYYY") & "'"
        '.Formulas(1) = "ToDate = '" & Format(DTPToDate.Value, "DD-MMM-YYYY") & "'"
        '.PrintReport

        '.Formulas(0) = "": .Formulas(1) = "":  .Formulas(2) = ""
        
    Set R = New ADODB.Recordset
    R.Open strRecord, MainConn, adOpenStatic
    
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = REPORT
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    R.Close
    
End Sub

Private Sub mnuSummaryPreview_Click()
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

    Set REPORT = Appl.OpenReport(findReportPath & "Supply Chain\rptStatementChequeSummary.rpt")
    strRecord = "Select * From viewStatement_Cheque_Summary Where CompanyID = " & bytCompID & " and Tran_Type = '" & cmbCashAndCredit.Text & "' and Tran_Date = '" & txtCDate.Text & "';"
             
    
    Set R = New ADODB.Recordset
    R.Open strRecord, MainConn, adOpenStatic
    
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = REPORT
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    R.Close
End Sub

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
                If txtBill.Text = .TextMatrix(inti, 3) And txtPartyID.Text = .TextMatrix(inti, 11) Then
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
    txtChalanNo.SelStart = 0
    txtChalanNo.SelLength = Len(txtChalanNo.Text)
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

'Private Sub txtMRRNO_GotFocus()
'    txtMRRNo.SelStart = 0
'    txtMRRNo.SelLength = Len(txtMRRNo.Text)
'End Sub

Private Sub txtNameOfParty_GotFocus()
    txtNameOfParty.SelStart = 0
    txtNameOfParty.SelLength = Len(txtNameOfParty.Text)
End Sub

Private Sub txtParticulers_GotFocus()
    txtParticulers.SelStart = 0
    txtParticulers.SelLength = Len(txtParticulers.Text)
End Sub

Private Sub txtParticulers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
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
'        , GoodsName
        Set RS = New Recordset
        
        sql1 = "Select SupplierID, SupplierName, SupplierAddress, GoodsName From Supplier_Information Where (((SupplierID)='" & txtPartyID.Text & "'))"
            RS.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
            
'            sql1 = "Select supplierID, SuppName, SuppAddress From New_Supplier Where (((SupplierID)='" & txtPartyID.Text & "'))"
'            RS.Open sql1, MainConn, adOpenDynamic, adLockPessimistic, adCmdText
error1:
        If err.Number = 3021 Then
            Set RS = Nothing
            MsgBox "This Is Not Valid Number", 16
            txtPartyID.SetFocus
        Else
            txtNameOfParty.Text = RS("SupplierName").Value
'''''''''''            txtParticulers.Text = RS("GoodsName").Value & ""
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


