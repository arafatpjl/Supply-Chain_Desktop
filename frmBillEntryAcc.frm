VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillEntryAcc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9269
   ClientLeft      =   13
   ClientTop       =   -39
   ClientWidth     =   9672
   ControlBox      =   0   'False
   Icon            =   "frmBillEntryAcc.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   713
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   744
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cboTaxPer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   1534
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1898
      Width           =   1729
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   156
      TabIndex        =   83
      Tag             =   "W"
      Top             =   910
      Width           =   2444
      Begin VB.OptionButton optS 
         BackColor       =   &H8000000E&
         Caption         =   "Supply Chain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   65
         TabIndex        =   85
         Top             =   -26
         Value           =   -1  'True
         Width           =   1404
      End
      Begin VB.OptionButton optM 
         BackColor       =   &H8000000E&
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   1521
         TabIndex        =   84
         Top             =   -26
         Width           =   988
      End
   End
   Begin VB.Frame framM 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   299
      Left            =   3250
      TabIndex        =   79
      Top             =   1898
      Width           =   6305
      Begin VB.TextBox txtTrDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   4706
         MaxLength       =   30
         TabIndex        =   7
         Top             =   0
         Width           =   1443
      End
      Begin VB.TextBox txtTrID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   1664
         MaxLength       =   30
         TabIndex        =   6
         Top             =   0
         Width           =   2301
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   3952
         TabIndex        =   81
         Top             =   0
         Width           =   767
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transaction ID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   52
         TabIndex        =   80
         Top             =   0
         Width           =   1625
      End
   End
   Begin VB.Frame framNature 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   494
      Left            =   1040
      TabIndex        =   50
      Tag             =   "W"
      Top             =   1300
      Width           =   1118
      Begin VB.OptionButton optCr 
         BackColor       =   &H8000000E&
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   91
         TabIndex        =   52
         Top             =   -26
         Value           =   -1  'True
         Width           =   884
      End
      Begin VB.OptionButton optCash 
         BackColor       =   &H8000000E&
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   91
         TabIndex        =   51
         Top             =   260
         Width           =   754
      End
   End
   Begin VB.Frame framST 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supply Chain Statement Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   6253
      Left            =   130
      TabIndex        =   47
      Top             =   2236
      Width           =   9425
      Begin VB.TextBox txtBillValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   130
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1612
         Width           =   1625
      End
      Begin VB.TextBox txtBillNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   130
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1092
         Width           =   1625
      End
      Begin VB.ComboBox cboSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   1794
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   520
         Width           =   4381
      End
      Begin VB.TextBox txtTotalTax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   7826
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   520
         Width           =   1456
      End
      Begin VB.TextBox txtTotalBill 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   6214
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   520
         Width           =   1625
      End
      Begin VB.ComboBox cboSource 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   1794
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2106
         Width           =   3848
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   7826
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2132
         Width           =   1456
      End
      Begin VB.ComboBox cboHead 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   182
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2678
         Width           =   4576
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   8034
         TabIndex        =   24
         Top             =   2522
         Width           =   1118
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   4784
         MaxLength       =   30
         TabIndex        =   23
         Top             =   2704
         Width           =   3052
      End
      Begin VB.TextBox txtTaxVal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   6630
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   29
         Top             =   2132
         Width           =   1222
      End
      Begin VB.TextBox txtTaxR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   5668
         MaxLength       =   30
         TabIndex        =   21
         Top             =   2132
         Width           =   988
      End
      Begin VB.TextBox txtVatVal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   7826
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   28
         Top             =   1612
         Width           =   1456
      End
      Begin VB.TextBox txtVateR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   6214
         MaxLength       =   30
         TabIndex        =   18
         Top             =   1612
         Width           =   1625
      End
      Begin VB.TextBox txtIncDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   299
         Left            =   130
         MaxLength       =   30
         TabIndex        =   19
         Top             =   2106
         Width           =   1625
      End
      Begin VB.TextBox txtNetBill 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   4784
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1612
         Width           =   1443
      End
      Begin VB.TextBox txtOths 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   3172
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1612
         Width           =   1638
      End
      Begin VB.TextBox txtDisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   1742
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1612
         Width           =   1456
      End
      Begin VB.TextBox txtMrrDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   7826
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1092
         Width           =   1456
      End
      Begin VB.TextBox txtMrrNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   6214
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1092
         Width           =   1638
      End
      Begin VB.TextBox txtChalDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   4784
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1092
         Width           =   1456
      End
      Begin VB.TextBox txtChalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   3172
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1092
         Width           =   1638
      End
      Begin VB.TextBox txtBillDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   1742
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1092
         Width           =   1456
      End
      Begin VB.TextBox txtSupplier 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   273
         Left            =   1742
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   234
         Width           =   4485
      End
      Begin MSFlexGridLib.MSFlexGrid flexBill 
         Height          =   3275
         Left            =   104
         TabIndex        =   53
         Top             =   3042
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   5775
         _Version        =   393216
         Rows            =   1
         Cols            =   20
         FixedCols       =   0
         RowHeightMin    =   275
         BackColor       =   16777215
         BackColorFixed  =   -2147483624
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483642
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         GridColor       =   -2147483630
         GridColorFixed  =   -2147483630
         Redraw          =   -1  'True
         AllowBigSelection=   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bill Incorp. Date"
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
         Height          =   247
         Index           =   5
         Left            =   130
         TabIndex        =   66
         Top             =   1872
         Width           =   1625
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bill Amount"
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
         Height          =   273
         Left            =   130
         TabIndex        =   62
         Top             =   1352
         Width           =   1625
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bill/Ref. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   130
         TabIndex        =   54
         Top             =   832
         Width           =   1625
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acc. Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   299
         Left            =   130
         TabIndex        =   77
         Top             =   494
         Width           =   1625
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Tax Amnt."
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
         Height          =   299
         Left            =   7826
         TabIndex        =   76
         Top             =   234
         Width           =   1456
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Bill Amnt."
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
         Height          =   299
         Left            =   6214
         TabIndex        =   75
         Top             =   234
         Width           =   1625
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   4784
         TabIndex        =   72
         Top             =   2444
         Width           =   3055
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rate (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   5668
         TabIndex        =   70
         Top             =   1872
         Width           =   975
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tax Deduction Source"
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
         Height          =   533
         Index           =   1
         Left            =   1742
         TabIndex        =   74
         Top             =   1872
         Width           =   3952
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Head of Accounts"
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
         Height          =   533
         Index           =   6
         Left            =   130
         TabIndex        =   67
         Top             =   2444
         Width           =   4680
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Cost"
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
         Height          =   273
         Left            =   7826
         TabIndex        =   73
         Top             =   1872
         Width           =   1456
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tax Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   6630
         TabIndex        =   71
         Top             =   1872
         Width           =   1222
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VAT Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   7826
         TabIndex        =   69
         Top             =   1352
         Width           =   1456
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VAT Rate (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   6214
         TabIndex        =   68
         Top             =   1352
         Width           =   1625
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Net Bill Amnt."
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
         Height          =   273
         Left            =   4784
         TabIndex        =   65
         Top             =   1352
         Width           =   1456
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(-) Others"
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
         Height          =   286
         Left            =   3172
         TabIndex        =   64
         Top             =   1352
         Width           =   1638
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(-) Discount"
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
         Height          =   286
         Left            =   1742
         TabIndex        =   63
         Top             =   1352
         Width           =   1456
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SC. Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   273
         Left            =   130
         TabIndex        =   60
         Top             =   234
         Width           =   1625
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MRR Date"
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
         Height          =   286
         Index           =   8
         Left            =   7826
         TabIndex        =   59
         Top             =   832
         Width           =   1456
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MRR No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   286
         Left            =   6214
         TabIndex        =   58
         Top             =   832
         Width           =   1625
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Challan Date"
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
         Height          =   286
         Index           =   7
         Left            =   4784
         TabIndex        =   57
         Top             =   832
         Width           =   1456
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Challan No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   286
         Left            =   3172
         TabIndex        =   56
         Top             =   832
         Width           =   1625
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bill Date"
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
         Height          =   286
         Index           =   0
         Left            =   1742
         TabIndex        =   55
         Top             =   832
         Width           =   1456
      End
   End
   Begin VB.Frame framSelect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Statement or Find Manual Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1092
      Left            =   2834
      TabIndex        =   40
      Top             =   728
      Width           =   6721
      Begin VB.CommandButton cmdFind 
         Caption         =   "Show Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   5018
         TabIndex        =   4
         Top             =   650
         Width           =   1534
      End
      Begin VB.ComboBox cboSTNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   1508
         TabIndex        =   3
         Top             =   676
         Width           =   3419
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Statement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   5018
         TabIndex        =   2
         Top             =   260
         Width           =   1534
      End
      Begin VB.Frame framDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   403
         Left            =   130
         TabIndex        =   41
         Top             =   208
         Width           =   4823
         Begin MSComCtl2.DTPicker txtDateFr 
            Height          =   299
            Left            =   1196
            TabIndex        =   0
            Top             =   78
            Width           =   1560
            _ExtentX        =   2756
            _ExtentY        =   527
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.1509
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125239297
            CurrentDate     =   40544
            MaxDate         =   44196
            MinDate         =   40544
         End
         Begin MSComCtl2.DTPicker txtDateTo 
            Height          =   299
            Left            =   3276
            TabIndex        =   1
            Top             =   78
            Width           =   1560
            _ExtentX        =   2756
            _ExtentY        =   527
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.1509
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125239297
            CurrentDate     =   40544
            MaxDate         =   44196
            MinDate         =   40544
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   325
            Left            =   3250
            TabIndex        =   45
            Top             =   52
            Width           =   1534
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   325
            Left            =   1170
            TabIndex        =   43
            Top             =   52
            Width           =   1534
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date From"
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
            Height          =   273
            Left            =   0
            TabIndex        =   42
            Top             =   78
            Width           =   1196
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "To"
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
            Height          =   273
            Left            =   2678
            TabIndex        =   44
            Top             =   78
            Width           =   598
         End
      End
      Begin VB.Label lblSTNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select ST. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   286
         Left            =   130
         TabIndex        =   46
         Top             =   676
         Width           =   1326
      End
   End
   Begin VB.Frame framS 
      BackColor       =   &H00E26101&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   481
      Left            =   2782
      TabIndex        =   35
      Top             =   8684
      Visible         =   0   'False
      Width           =   3796
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         Enabled         =   0   'False
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
         Left            =   1365
         TabIndex        =   32
         ToolTipText     =   "Preview"
         Top             =   52
         Width           =   936
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
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
         Left            =   65
         TabIndex        =   33
         ToolTipText     =   "Delete Record"
         Top             =   52
         Visible         =   0   'False
         Width           =   936
      End
      Begin VB.CommandButton cmdSave 
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
         Height          =   375
         Left            =   2665
         TabIndex        =   31
         ToolTipText     =   "Save Record"
         Top             =   52
         Width           =   936
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E26101&
      BorderStyle     =   0  'None
      Height          =   481
      Left            =   8086
      TabIndex        =   38
      Top             =   8684
      Width           =   1222
      Begin VB.CommandButton cmdClose 
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
         Height          =   375
         Left            =   130
         TabIndex        =   34
         ToolTipText     =   "Form Close"
         Top             =   52
         Width           =   1014
      End
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Height          =   325
      Left            =   130
      TabIndex        =   86
      Top             =   858
      Width           =   2587
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Load Supplier Bills From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   286
      Left            =   130
      TabIndex        =   82
      Top             =   598
      Width           =   2587
   End
   Begin VB.Label lblTaxPer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In.Tax Period"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   273
      Left            =   130
      TabIndex        =   78
      Top             =   1898
      Width           =   1365
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
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
      Height          =   637
      Left            =   1014
      TabIndex        =   49
      Top             =   1222
      Width           =   1261
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Nature of Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   468
      Left            =   130
      TabIndex        =   48
      Top             =   1300
      Width           =   897
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E26101&
      Height          =   689
      Index           =   0
      Left            =   0
      TabIndex        =   39
      Top             =   8606
      Width           =   9815
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      BackColor       =   &H00E26101&
      Caption         =   "Update Supplier Bills for Accounts - Credit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   16.3
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   442
      Left            =   104
      TabIndex        =   37
      Top             =   52
      Width           =   9477
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E26101&
      Height          =   559
      Index           =   1
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   9919
   End
End
Attribute VB_Name = "frmBillEntryAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytRow As Byte
Private Sub cboHead_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub cboSource_Click()
  txtTaxR.Text = "0.00"
  txtTaxVal.Text = "0.00"
  If Len(cboSource) = 0 Or Len(cboTaxPer) = 0 Then Exit Sub
  
  Set R = New ADODB.Recordset
  If Len(txtTotalBill) = 0 Then
    strRecord = "SELECT TaxRate FROM Acc_Income_Tax_Rate WHERE TaxPerID = " & cboTaxPer.Tag & "" _
              & " AND SourceName = '" & cboSource.Text & "' AND 100 BETWEEN RangeFr AND RangeTo;"
  Else
    strRecord = "SELECT TaxRate FROM Acc_Income_Tax_Rate WHERE TaxPerID = " & cboTaxPer.Tag & "" _
              & " AND SourceName = '" & cboSource.Text & "' AND " & CCur(txtTotalBill.Text) & " BETWEEN RangeFr AND RangeTo;"
  End If
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then txtTaxR.Text = Format(R![TaxRate], "##0.00")
  R.Close
  Set R = Nothing
End Sub
Private Sub cboSource_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub cboTaxPer_Click()
  cboTaxPer.Tag = 0
  If Len(cboTaxPer) = 0 Then Exit Sub
  cboTaxPer.Tag = cboTaxPer.ItemData(cboTaxPer.ListIndex)
  
  If cboTaxPer.Text <> lblTaxPer.Tag Then
    Call addTexSource(cboSource, cboTaxPer.Tag)
    lblTaxPer.Tag = cboTaxPer.Text
  End If
End Sub
Private Sub cboTaxPer_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub cmdAdd_Click()
  
  If Len(cboTaxPer) = 0 Then
    MsgBox "Select Income Tax Period", vbInformation, cnstMsgInfo
    cboTaxPer.SetFocus
    Exit Sub
  End If
  If optS.Value = True Then
    If Len(txtTrDate) = 0 Then
      MsgBox "Enter Transaction Date", vbInformation, cnstMsgInfo
      txtTrDate.SetFocus
      Exit Sub
    End If
  Else
    If Len(txtIncDate) = 0 Then
      MsgBox "Enter Bill Incorporate Date", vbInformation, cnstMsgInfo
      txtIncDate.SetFocus
      Exit Sub
    End If
  End If
  If Len(cboSource) = 0 Then
    MsgBox "Select Head of Accounts", vbInformation, cnstMsgInfo
    cboHead.SetFocus
    Exit Sub
  End If
  If Len(cboHead) = 0 Then
    MsgBox "Select Head of Accounts", vbInformation, cnstMsgInfo
    cboHead.SetFocus
    Exit Sub
  End If
    
  If bytRow = 0 Then
    MsgBox "No Selected Row Found", vbInformation, cnstMsgInfo
    Exit Sub
  End If

  With flexBill
    .Row = bytRow
    
    For inti = 0 To .Cols - 1
      .Col = inti
      .CellFontBold = False
    Next inti

    .Col = 5:   .Text = txtDisc.Text
    .Col = 6:   .Text = txtOths.Text
    .Col = 7:   .Text = txtNetBill.Text
    .Col = 8:   .Text = txtVateR.Text
    .Col = 9:   .Text = txtVatVal.Text
    .Col = 10:  .Text = cboTaxPer.Text
    .Col = 11:  .Text = txtTaxR.Text
    .Col = 12:  .Text = txtTaxVal.Text
    .Col = 13:  .Text = txtTotalCost.Text
    .Col = 14:  .Text = cboHead.Text
    If Len(txtRemarks) = 0 Then
      .Col = 20:  .Text = ""
    Else
      .Col = 19
      If txtRemarks <> .Text Then
        .Col = 20:  .Text = txtRemarks.Text
      End If
    End If
    .Col = 21:  .Text = txtIncDate.Text
    
  End With
  bytRow = 0
  Call nullTextBox
End Sub
Private Sub cmdClose_Click()
  If MsgBox("Are you sure to Close", vbInformation + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  Unload Me
End Sub
Private Sub cmdFind_Click()
  ''MsgBox Right(cboSTNo.Text, 11)
  ''MsgBox Left(cboSTNo.Text, Len(cboSTNo) - 15)
  
  flexBill.Rows = 1
  Screen.MousePointer = vbHourglass
  
  If optS.Value = True Then
    strRecord = "SELECT Tran_Type,Sl_No,NameOfTheParty,PartyName,BillNo,BillDate,MRRNo,MRRDate,Amount,Remark AS SCRemark,SCNO,SupplierID,Chalan_Date,Chalan_No FROM Statement_of_Cash_and_Credit A LEFT JOIN Party B ON A.NameOfTheParty = B.PartyName" _
              & " WHERE A.CompanyID = " & bytCompID & " AND A.Tran_Date = '" & Right(cboSTNo.Text, 11) & "' AND A.Tran_Type = '" & lblExp.Tag & "' AND A.SCNO = '" & Left(cboSTNo.Text, Len(cboSTNo) - 15) & "' AND A.Deleterow = 0 ORDER BY Tr_No;"
  Else
    strRecord = "SELECT Tran_Type,Sl_No,NameOfTheParty,PartyName,BillNo,BillDate,MRRNo,MRRDate,Amount,Remark AS SCRemark,SCNO,SupplierID,Chalan_Date,Chalan_No FROM Statement_of_Cash_and_Credit A LEFT JOIN Party B ON A.NameOfTheParty = B.PartyName" _
              & " WHERE A.CompanyID = " & bytCompID & " AND A.Tran_Date = '" & Right(cboSTNo.Text, 11) & "' AND A.Tran_Type = '" & lblExp.Tag & "' AND A.SCNO = '" & Left(cboSTNo.Text, Len(cboSTNo) - 15) & "' AND A.Deleterow = 0 ORDER BY Tr_No;"
  End If
  
  Set R = New ADODB.Recordset
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    With flexBill
      .Redraw = False
      For inti = 1 To R.RecordCount
        .Rows = .Rows + 1
        .Row = .Rows - 1

        .Col = 0:   .Text = R![Sl_No]
        .Col = 1:   .Text = R![NameOfTheParty]
        .Col = 2:   .Text = R![BillNo]
        .Col = 3:   .Text = Format(R![BillDate], "DD-MMM-YY")
        .Col = 4:   .Text = Format(R![Amount], "##,##0.00")
        .Col = 5:   .Text = "0.00" ''Format(R![], "##,##0.00")
        .Col = 6:   .Text = "0.00" ''Format(R![], "##,##0.00")
        .Col = 7:   .Text = Format(R![Amount], "##,##0.00")
        .Col = 8:   .Text = "0.00" ''Format(R![], "##,##0.00")
        .Col = 9:   .Text = "0.00" ''Format(R![], "##,##0.00")
        .Col = 10:  .Text = "" ''R![]
        .Col = 11:  .Text = "0.00" ''Format(R![], "##,##0.00")
        .Col = 12:  .Text = "0.00" ''Format(R![], "##,##0.00")
        .Col = 13:  .Text = Format(R![Amount], "##,##0.00")
        .Col = 14:  .Text = "" ''R![]
        .Col = 15:  .Text = R![MRRNo]
        .Col = 16:  .Text = Format(R![mrrdate], "DD-MMM-YY")
        .Col = 17:  .Text = R![Chalan_No]
        .Col = 18:  .Text = Format(R![Chalan_Date], "DD-MMM-YY")
        .Col = 19:  .Text = R![SCRemark]
        .Col = 20:  .Text = "" ''R![ACCRemark]
        .Col = 21:  .Text = "" ''Format(R![], "DD-MMM-YY")
        .Col = 22:  .Text = IIf(IsNull(R![PartyName]), "", R![PartyName])
        R.MoveNext
      Next inti
      .Row = 0
      .Redraw = True
    End With
  End If
  R.Close
  Set R = Nothing
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdLoad_Click()
  cboSTNo.Clear
  Screen.MousePointer = vbHourglass
  
  If optS.Value = True Then
    strRecord = "SELECT DISTINCT Tran_Date,SCNo FROM Statement_of_Cash_and_Credit WHERE CompanyID = " & bytCompID & " AND Tran_Type = '" & lblExp.Tag & "'" _
              & " AND Tran_Date BETWEEN '" & txtDateFr.Value & "' AND '" & txtDateTo.Value & "' ORDER BY Tran_Date,SCNo;"
  Else
    strRecord = "SELECT DISTINCT Tran_Date,SCNo FROM Statement_of_Cash_and_Credit WHERE CompanyID = " & bytCompID & " AND Tran_Type = '" & lblExp.Tag & "'" _
              & " AND Tran_Date BETWEEN '" & txtDateFr.Value & "' AND '" & txtDateTo.Value & "' ORDER BY Tran_Date,SCNo;"
  End If
  
  Set R = New ADODB.Recordset
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    Do Until R.EOF
      cboSTNo.AddItem R![SCNO] & " DT." & Format(R![Tran_Date], cnstDtFrmtI)
      R.MoveNext
    Loop
  End If
  R.Close
  Set R = Nothing
  Screen.MousePointer = vbDefault
End Sub
Private Sub flexBill_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call nullTextBox
    
    With flexBill
      If .RowSel = 0 Then Exit Sub
      .Col = 0:   .CellFontBold = True
      .Col = 1:   .CellFontBold = True:    txtSupplier.Text = .Text
      .Col = 2:   .CellFontBold = True:    txtBillNo.Text = .Text
      .Col = 3:   .CellFontBold = True:    txtBillDate.Text = .Text
      .Col = 4:   .CellFontBold = True:    txtBillValue.Text = .Text
      .Col = 5:   .CellFontBold = True:    txtDisc.Text = .Text
      .Col = 6:   .CellFontBold = True:    txtOths.Text = .Text
      .Col = 7:   .CellFontBold = True:    txtNetBill.Text = .Text
      .Col = 8:   .CellFontBold = True:    txtVateR.Text = .Text
      .Col = 9:   .CellFontBold = True:    txtVatVal.Text = .Text
      ''.Col = 10:  .CellFontBold = True:    If Len(.Text) > 0 Then cboTaxPer.Text = .Text
      .Col = 11:  .CellFontBold = True:    txtTaxR.Text = .Text
      .Col = 12:  .CellFontBold = True:    txtTaxVal.Text = .Text
      .Col = 13:  .CellFontBold = True:    txtTotalCost.Text = .Text
      .Col = 14:  .CellFontBold = True:    If Len(.Text) > 0 Then cboHead.Text = .Text
      .Col = 15:  .CellFontBold = True:    txtMRRNo.Text = .Text
      .Col = 16:  .CellFontBold = True:    txtMRRDate.Text = .Text
      .Col = 17:  .CellFontBold = True:    txtChalNo.Text = .Text
      .Col = 18:  .CellFontBold = True:    txtChalDate.Text = .Text
      .Col = 19:  .CellFontBold = True:    txtRemarks.Text = .Text
      .Col = 21:  .CellFontBold = True:    txtIncDate.Text = .Text
      .Col = 22:  .CellFontBold = True:    If Len(.Text) > 0 Then cboSupplier.Text = .Text
      bytRow = .Row

    End With
    txtDisc.SetFocus
  End If
End Sub

Private Sub Form_Load()

Screen.MousePointer = vbHourglass

Call prcMakeCenter(Me)

txtDateFr.Value = Date
txtDateTo.Value = Date

Call addTexPeriod(cboTaxPer)
Call addAccHead(cboHead)
Call addAccSupplier(cboSupplier)

With flexBill
  .Width = 9645
  .Height = 3275
  
  .Rows = 1
  .Cols = 23
  .Row = 0
  .ColAlignment(0) = flexAlignCenterCenter
  .ColAlignment(1) = flexAlignLeftCenter
  .ColAlignment(2) = flexAlignLeftCenter
  .ColAlignment(3) = flexAlignCenterCenter
  .ColAlignment(4) = flexAlignRightCenter
  .ColAlignment(5) = flexAlignRightCenter
  .ColAlignment(6) = flexAlignRightCenter
  .ColAlignment(7) = flexAlignRightCenter
  .ColAlignment(8) = flexAlignCenterCenter
  .ColAlignment(9) = flexAlignRightCenter
  .ColAlignment(10) = flexAlignCenterCenter
  .ColAlignment(11) = flexAlignCenterCenter
  .ColAlignment(12) = flexAlignRightCenter
  .ColAlignment(13) = flexAlignRightCenter
  .ColAlignment(14) = flexAlignLeftCenter
  .ColAlignment(15) = flexAlignLeftCenter
  .ColAlignment(16) = flexAlignCenterCenter
  .ColAlignment(17) = flexAlignLeftCenter
  .ColAlignment(18) = flexAlignCenterCenter
  .ColAlignment(19) = flexAlignLeftCenter
  .ColAlignment(20) = flexAlignLeftCenter
  .ColAlignment(21) = flexAlignCenterCenter
  .ColAlignment(22) = flexAlignCenterCenter
  
  .Col = 0:   .CellFontBold = True:  .Text = "SL#"
  .Col = 1:   .CellFontBold = True:  .Text = "Supplier"
  .Col = 2:   .CellFontBold = True:  .Text = "Bill/Ref."
  .Col = 3:   .CellFontBold = True:  .Text = "Bill DT"
  .Col = 4:   .CellFontBold = True:  .Text = "Bill Amnt"
  .Col = 5:   .CellFontBold = True:  .Text = "Discount"
  .Col = 6:   .CellFontBold = True:  .Text = "Others"
  .Col = 7:   .CellFontBold = True:  .Text = "Net Amnt"
  .Col = 8:   .CellFontBold = True:  .Text = "VatRate"
  .Col = 9:   .CellFontBold = True:  .Text = "Vat Amnt"
  ''.Col = 10:  .CellFontBold = True:  .Text = ""
  .Col = 11:  .CellFontBold = True:  .Text = "Tax Rate"
  .Col = 12:  .CellFontBold = True:  .Text = "Tax Amnt"
  .Col = 13:  .CellFontBold = True:  .Text = "TTL Cost"
  .Col = 14:  .CellFontBold = True:  .Text = "A/C Head"
  .Col = 15:  .CellFontBold = True:  .Text = "MRR No"
  .Col = 16:  .CellFontBold = True:  .Text = "MRR DT"
  .Col = 17:  .CellFontBold = True:  .Text = "Challan"
  .Col = 18:  .CellFontBold = True:  .Text = "Chal.DT"
  .Col = 19:  .CellFontBold = True:  .Text = "SC. Remarks"
  .Col = 20:  .CellFontBold = True:  .Text = "ACC. Remarks"
  .Col = 21:  .CellFontBold = True:  .Text = "Bill Inc.DT"
  .Col = 22:  .CellFontBold = True:  .Text = "Acc.Party"

  .ColWidth(0) = 450:      .ColWidth(1) = 1800
  .ColWidth(2) = 1000:     .ColWidth(3) = 900
  .ColWidth(4) = 950:      .ColWidth(5) = 800
  .ColWidth(6) = 800:      .ColWidth(7) = 950
  .ColWidth(8) = 800:      .ColWidth(9) = 900
  .ColWidth(10) = 0:       .ColWidth(11) = 800
  .ColWidth(12) = 900:     .ColWidth(13) = 950
  .ColWidth(14) = 1500:    .ColWidth(15) = 1500
  .ColWidth(16) = 900:     .ColWidth(17) = 1500
  .ColWidth(18) = 900:     .ColWidth(19) = 1500
  .ColWidth(20) = 1500:    .ColWidth(21) = 1000
  .ColWidth(22) = 1500
End With

cboTaxPer.ListIndex = 0
lblExp.Tag = "Credit"
Screen.MousePointer = vbDefault
If strUserName = "PACIFIC" Then cmdDelete.Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmBillEntryAcc = Nothing
End Sub
Private Sub optCash_Click()
  If optCash.Value = True Then
    lblExp.Tag = "Cash"
    lblTop.Caption = "Update Supplier Bills for Accounts - Cash"
  Else
    lblExp.Tag = "Credit"
    lblTop.Caption = "Update Supplier Bills for Accounts - Credit"
  End If
  cboSTNo.Clear
  Call nullTextBox
  flexBill.Rows = 1
  txtIncDate.Text = ""
End Sub
Private Sub optCr_Click()
  optCash_Click
End Sub
Private Sub optM_Click()
  Call optS_Click
End Sub
Private Sub optS_Click()
  txtTrID.Text = ""
  txtTrDate.Text = ""
  If optS.Value = True Then
    cmdLoad.Caption = "Load Statement"
    lblSTNo.Caption = "Select ST. No."
    EnableFrame framM, False
    
    txtBillNo.BackColor = &H80000016:     txtBillNo.Locked = True
    txtBillDate.BackColor = &H80000016:   txtBillDate.Locked = True
    txtChalNo.BackColor = &H80000016:     txtChalNo.Locked = True
    txtChalDate.BackColor = &H80000016:   txtChalDate.Locked = True
    txtMRRNo.BackColor = &H80000016:      txtMRRNo.Locked = True
    txtMRRDate.BackColor = &H80000016:    txtMRRDate.Locked = True
    txtBillValue.BackColor = &H80000016:  txtBillValue.Locked = True
    
  Else
    cmdLoad.Caption = "Load Transaction"
    lblSTNo.Caption = "Select TR. No."
    EnableFrame framM, True
    
    txtBillNo.BackColor = &H8000000E:     txtBillNo.Locked = False
    txtBillDate.BackColor = &H8000000E:   txtBillDate.Locked = False
    txtChalNo.BackColor = &H8000000E:     txtChalNo.Locked = False
    txtChalDate.BackColor = &H8000000E:   txtChalDate.Locked = False
    txtMRRNo.BackColor = &H8000000E:      txtMRRNo.Locked = False
    txtMRRDate.BackColor = &H8000000E:    txtMRRDate.Locked = False
    txtBillValue.BackColor = &H8000000E:  txtBillValue.Locked = False
  End If
End Sub
Private Sub txtBillDate_GotFocus()
  txtGotFocus txtBillDate
End Sub
Private Sub txtBillDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtBillDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtBillDate_LostFocus()
  If optM.Value = True Then txtBillDate.Text = FormatDate(txtBillDate)
End Sub
Private Sub txtBillNo_GotFocus()
  txtGotFocus txtBillNo
End Sub
Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtBillNo_LostFocus()
  txtBillNo.Text = Trim(UCase(txtBillNo.Text))
End Sub
Private Sub txtBillValue_GotFocus()
  txtGotFocus txtBillValue
End Sub
Private Sub txtBillValue_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtBillValue_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtBillValue_LostFocus()
  txtBillValue.Text = Format(txtBillValue.Text, "##,##0.00")
End Sub
Private Sub txtChalDate_GotFocus()
  txtGotFocus txtChalDate
End Sub
Private Sub txtChalDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtChalDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtChalDate_LostFocus()
  If optM.Value = True Then txtChalDate.Text = FormatDate(txtChalDate)
End Sub
Private Sub txtChalNo_GotFocus()
  txtGotFocus txtChalNo
End Sub
Private Sub txtChalNo_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtChalNo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtChalNo_LostFocus()
  txtChalNo.Text = Trim(UCase(txtChalNo.Text))
End Sub
Private Sub txtDisc_GotFocus()
  txtGotFocus txtDisc
End Sub
Private Sub txtDisc_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtDisc_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtDisc_LostFocus()
  txtDisc.Text = Format(txtDisc.Text, "##,##0.00")
End Sub
Private Sub txtIncDate_GotFocus()
  txtGotFocus txtIncDate
End Sub
Private Sub txtIncDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtIncDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtIncDate_LostFocus()
  txtIncDate.Text = FormatDate(txtIncDate)
End Sub
Private Sub txtMRRDate_GotFocus()
  txtGotFocus txtMRRDate
End Sub
Private Sub txtMRRDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtMRRDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtMRRDate_LostFocus()
  If optM.Value = True Then txtMRRDate.Text = FormatDate(txtMRRDate)
End Sub
Private Sub txtMRRNO_GotFocus()
  txtGotFocus txtMRRNo
End Sub
Private Sub txtMRRNO_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtMrrNo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtMrrNo_LostFocus()
  txtMRRNo.Text = Trim(UCase(txtMRRNo.Text))
End Sub
Private Sub txtNetBill_GotFocus()
  txtGotFocus txtNetBill
End Sub
Private Sub txtNetBill_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtOths_GotFocus()
  txtGotFocus txtOths
End Sub
Private Sub txtOths_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtOths_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtOths_LostFocus()
  txtOths.Text = Format(txtOths.Text, "##,##0.00")
End Sub
Private Sub txtRemarks_GotFocus()
  txtGotFocus txtRemarks
End Sub
Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtRemarks_LostFocus()
  txtRemarks.Text = Trim(UCase(txtRemarks.Text))
End Sub
Private Sub txtTaxR_GotFocus()
  txtGotFocus txtTaxR
End Sub
Private Sub txtTaxR_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtTaxR_LostFocus()
  txtTaxR.Text = Format(txtTaxR.Text, "##,##0.00")
End Sub
Private Sub txtTaxVal_GotFocus()
  txtGotFocus txtTaxVal
End Sub
Private Sub txtTaxVal_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtTotalCost_GotFocus()
  txtGotFocus txtTotalCost
End Sub
Private Sub txtTotalCost_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtTrDate_GotFocus()
  txtGotFocus txtTrDate
End Sub
Private Sub txtTrDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtTrDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtTrDate_LostFocus()
  txtTrDate.Text = FormatDate(txtTrDate)
End Sub
Private Sub txtTrID_GotFocus()
  txtGotFocus txtTrID
End Sub
Private Sub txtTrID_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtTrID_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtTrID_LostFocus()
  txtTrID.Text = Trim(UCase(txtTrID.Text))
End Sub
Private Sub txtVateR_GotFocus()
  txtGotFocus txtVateR
End Sub
Private Sub txtVateR_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtVateR_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtVateR_LostFocus()
  txtVateR.Text = Format(txtVateR.Text, "##,##0.00")
End Sub
Private Sub txtVatVal_GotFocus()
  txtGotFocus txtVatVal
End Sub
Private Sub txtVatVal_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub

Sub nullTextBox()
  txtSupplier.Text = ""
  cboSupplier.ListIndex = -1
  txtTotalBill.Text = ""
  txtTotalTax.Text = ""
  txtBillNo.Text = ""
  txtBillDate.Text = ""
  txtBillValue.Text = ""
  txtDisc.Text = ""
  txtOths.Text = ""
  txtNetBill.Text = ""
  txtVateR.Text = ""
  txtVatVal.Text = ""
  txtTaxR.Text = ""
  txtTaxVal.Text = ""
  txtTotalCost.Text = ""
  txtMRRNo.Text = ""
  txtMRRDate.Text = ""
  txtChalNo.Text = ""
  txtChalDate.Text = ""
  cboHead.ListIndex = -1
  cboSource.ListIndex = -1
  txtRemarks.Text = ""
End Sub

Sub addAccHead(cboMain As ComboBox)
  cboMain.Clear
  Set R = New ADODB.Recordset
  strRecord = "SELECT HeadID,HeadName FROM InventoryPJL.dbo.AccountsHead ORDER BY HeadName;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    Do Until R.EOF
      cboMain.AddItem R![HeadName]
      cboMain.ItemData(cboMain.NewIndex) = R![HeadID]
      R.MoveNext
    Loop
  End If
  R.Close
  Set R = Nothing
End Sub

Sub addTexPeriod(cboMain As ComboBox)
  cboMain.Clear
  Set R = New ADODB.Recordset
  strRecord = "SELECT TaxPerID,TaxPeriod FROM InventoryPJL.dbo.Acc_New_Tax_Period ORDER BY TaxPerID DESC;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    Do Until R.EOF
      cboMain.AddItem R![TaxPeriod]
      cboMain.ItemData(cboMain.NewIndex) = R![TaxPerID]
      R.MoveNext
    Loop
  End If
  R.Close
  Set R = Nothing
End Sub

Sub addTexSource(cboMain As ComboBox, parTaxPerID As Integer)
  cboMain.Clear
  Set R = New ADODB.Recordset
  strRecord = "SELECT DISTINCT SourceName FROM InventoryPJL.dbo.Acc_Income_Tax_Rate WHERE TaxPerID = " & parTaxPerID & " ORDER BY SourceName;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    Do Until R.EOF
      cboMain.AddItem R![SourceName]
      R.MoveNext
    Loop
  End If
  R.Close
  Set R = Nothing
End Sub
Sub addAccSupplier(cboMain As ComboBox)
  cboMain.Clear
  Set R = New ADODB.Recordset
  strRecord = "SELECT PartyID,PartyName FROM InventoryPJL.dbo.Party ORDER BY PartyName;"
  R.Open strRecord, MainConn, adOpenStatic
  If R.RecordCount > 0 Then
    Do Until R.EOF
      cboMain.AddItem R![PartyName]
      cboMain.ItemData(cboMain.NewIndex) = R![PartyID]
      R.MoveNext
    Loop
  End If
  R.Close
  Set R = Nothing
End Sub

