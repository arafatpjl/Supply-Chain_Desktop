VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMrrReceiveChallan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9840
   ClientLeft      =   15
   ClientTop       =   -150
   ClientWidth     =   15180
   ControlBox      =   0   'False
   Icon            =   "frmMrrReceiveChallan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   15180
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PO Item Information "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1768
      Left            =   6266
      TabIndex        =   94
      Top             =   923
      Width           =   8866
      Begin MSFlexGridLib.MSFlexGrid flexItem 
         Height          =   1599
         Left            =   26
         TabIndex        =   95
         Top             =   208
         Width           =   9100
         _ExtentX        =   16060
         _ExtentY        =   2831
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   0
         RowHeightMin    =   275
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   -2147483632
         ForeColorFixed  =   16776960
         BackColorSel    =   -2147483642
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         GridColor       =   -2147483630
         GridColorFixed  =   -2147483630
         Redraw          =   -1  'True
         AllowBigSelection=   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   6125
      Left            =   11760
      TabIndex        =   76
      Top             =   2756
      Width           =   3375
      Begin VB.TextBox Name5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   5603
         Width           =   3135
      End
      Begin VB.TextBox Name4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   4459
         Width           =   3135
      End
      Begin VB.TextBox Name3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   3289
         Width           =   3135
      End
      Begin VB.TextBox Name2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   2054
         Width           =   3135
      End
      Begin VB.TextBox Name1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   884
         Width           =   3135
      End
      Begin VB.TextBox Empcode5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   85
         Top             =   5070
         Width           =   615
      End
      Begin VB.TextBox Empcode4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   84
         Top             =   4030
         Width           =   1095
      End
      Begin VB.TextBox Empcode3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   83
         Top             =   2860
         Width           =   1095
      End
      Begin VB.TextBox Empcode2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   82
         Top             =   1625
         Width           =   1095
      End
      Begin VB.TextBox Empcode1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   81
         Top             =   468
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ASST.GENERAL MANAGER (MCD)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   5115
         Width           =   2400
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MANAGER (MCD)"
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
         Height          =   260
         Left            =   117
         TabIndex        =   80
         Top             =   4030
         Width           =   1820
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QUANTITY CHECKED BY"
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
         Height          =   260
         Left            =   117
         TabIndex        =   79
         Top             =   2860
         Width           =   2054
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PREPARED BY"
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
         Height          =   260
         Left            =   117
         TabIndex        =   78
         Top             =   1625
         Width           =   1820
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QUALITY CERTIFIED BY"
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
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   468
         Width           =   1935
      End
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
      Left            =   5616
      TabIndex        =   72
      ToolTipText     =   "Delete Record"
      Top             =   9120
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   9450
      TabIndex        =   71
      ToolTipText     =   "Save/Update Record"
      Top             =   9120
      Width           =   1185
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0C000&
      Caption         =   " &Refresh"
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
      Left            =   6903
      TabIndex        =   70
      ToolTipText     =   "Refresh"
      Top             =   9120
      UseMaskColor    =   -1  'True
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
      Left            =   10695
      TabIndex        =   69
      ToolTipText     =   "Close Form"
      Top             =   9120
      Width           =   1185
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
      Left            =   4329
      TabIndex        =   68
      ToolTipText     =   "Preview PI"
      Top             =   9120
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
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
      Left            =   8177
      TabIndex        =   67
      ToolTipText     =   "New/Edit Record"
      Top             =   9120
      Width           =   1185
   End
   Begin VB.TextBox txtMode 
      Alignment       =   2  'Center
      BackColor       =   &H00E26101&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   405
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   66
      ToolTipText     =   "Mode of Data Entry"
      Top             =   9210
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   150
      TabIndex        =   40
      Top             =   960
      Width           =   6084
      Begin VB.CommandButton cmdReload 
         BackColor       =   &H00C0C000&
         Caption         =   "Reload PO Item"
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
         Left            =   3302
         TabIndex        =   93
         ToolTipText     =   "Show Record"
         Top             =   1248
         Width           =   1547
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Show"
         Height          =   255
         Left            =   4966
         TabIndex        =   92
         Top             =   1287
         Width           =   767
      End
      Begin VB.TextBox vmrrid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   117
         TabIndex        =   60
         Top             =   1534
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbMRRID 
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
         Height          =   286
         Left            =   1530
         TabIndex        =   1
         Top             =   1530
         Visible         =   0   'False
         Width           =   3354
      End
      Begin VB.CommandButton cmdNewSupplier 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5670
         Picture         =   "frmMrrReceiveChallan.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Add New Supplier"
         Top             =   420
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cmbSupplier 
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
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   5445
      End
      Begin VB.CommandButton cmdShow 
         BackColor       =   &H00C0C000&
         Caption         =   "&Show"
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
         Left            =   4992
         TabIndex        =   2
         ToolTipText     =   "Show Record"
         Top             =   832
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMRRID 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   52
         Top             =   840
         Width           =   3328
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
         Left            =   1530
         TabIndex        =   3
         Top             =   1248
         Width           =   1651
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier Name : "
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
         Height          =   255
         Left            =   150
         TabIndex        =   46
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR No."
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
         Height          =   260
         Left            =   156
         TabIndex        =   45
         Top             =   871
         Width           =   949
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MRR Date"
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
         Height          =   260
         Index           =   0
         Left            =   156
         TabIndex        =   44
         Top             =   1287
         Width           =   1092
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
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
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
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
         Height          =   260
         Left            =   1326
         TabIndex        =   42
         Top             =   1287
         Width           =   195
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mrr Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E26101&
      Height          =   6135
      Left            =   150
      TabIndex        =   29
      Top             =   2745
      Width           =   11595
      Begin VB.ComboBox txtMRRQty 
         Height          =   315
         Left            =   5400
         TabIndex        =   96
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cmbDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5895
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   2460
      End
      Begin VB.TextBox txtCont 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   9435
         MaxLength       =   30
         TabIndex        =   8
         Top             =   420
         Width           =   1980
      End
      Begin VB.ComboBox cmbprno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   2160
         TabIndex        =   14
         Top             =   1560
         Width           =   3135
      End
      Begin VB.ComboBox cmbPINO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6495
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   420
         Width           =   2865
      End
      Begin VB.CommandButton cmdNewMainPF 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6150
         Picture         =   "frmMrrReceiveChallan.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Add New Main PF"
         Top             =   78
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cmbMainPF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   4560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   1859
      End
      Begin VB.TextBox txtChallanDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   286
         Left            =   2550
         TabIndex        =   5
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox txtChallanNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   286
         Left            =   180
         TabIndex        =   4
         Top             =   420
         Width           =   2280
      End
      Begin VB.TextBox txtShortExcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   8580
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1530
      End
      Begin VB.TextBox txtReceivedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   7080
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1560
         Width           =   1440
      End
      Begin VB.TextBox txtRemarks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   4260
         MaxLength       =   500
         TabIndex        =   17
         Top             =   2130
         Width           =   6045
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
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
         Left            =   3960
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   5730
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtCharges 
         Alignment       =   2  'Center
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
         Left            =   6870
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   5730
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cmbItemType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   990
         Width           =   2340
      End
      Begin VB.ComboBox cmbItem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   286
         Left            =   2580
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   990
         Width           =   2860
      End
      Begin VB.CommandButton cmdAddNewItem 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5505
         Picture         =   "frmMrrReceiveChallan.frx":0626
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Add New Item"
         Top             =   990
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cmbCurrency 
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
         Left            =   1260
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   5730
         Width           =   1050
      End
      Begin VB.TextBox txtmatcode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   208
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1560
         Width           =   1898
      End
      Begin VB.CommandButton cmdAdd 
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
         Height          =   345
         Left            =   10410
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add to List"
         Top             =   2100
         Width           =   1035
      End
      Begin VB.TextBox txtMRRQty_Or 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2130
         Width           =   2070
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   286
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1815
      End
      Begin VB.ComboBox cmbUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10230
         Locked          =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         MaxLength       =   255
         TabIndex        =   12
         ToolTipText     =   "Item Description"
         Top             =   990
         Width           =   3015
      End
      Begin VB.TextBox txtMRRValue 
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
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   9750
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   5760
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   3198
         Left            =   117
         TabIndex        =   73
         Top             =   2522
         Width           =   11817
         _ExtentX        =   20849
         _ExtentY        =   5636
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   285
         BackColor       =   -2147483626
         ForeColor       =   0
         BackColorFixed  =   14737632
         ForeColorFixed  =   14835969
         BackColorSel    =   -2147483642
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         GridColor       =   -2147483630
         GridColorFixed  =   -2147483630
         Enabled         =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "PO Description :"
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
         Left            =   5895
         TabIndex        =   75
         Top             =   765
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Container No :"
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
         Height          =   255
         Index           =   0
         Left            =   9435
         TabIndex        =   74
         Top             =   210
         Width           =   1875
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unit :"
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
         Height          =   255
         Left            =   10260
         TabIndex        =   59
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO No. :"
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
         Height          =   255
         Left            =   6495
         TabIndex        =   57
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label lblActPF 
         BackColor       =   &H00FFFFFF&
         Caption         =   "File No :"
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
         Height          =   255
         Left            =   4560
         TabIndex        =   56
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Challan/Invoice Date :"
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
         Height          =   255
         Left            =   2550
         TabIndex        =   55
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Challan/Invoice No :"
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
         Height          =   255
         Left            =   210
         TabIndex        =   54
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Short/Excess : "
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
         Height          =   255
         Left            =   8640
         TabIndex        =   53
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Received Qty : "
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
         Height          =   255
         Left            =   7080
         TabIndex        =   51
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remarks : "
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
         Left            =   4290
         TabIndex        =   50
         Top             =   1890
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Discount (Less) :"
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
         Index           =   3
         Left            =   2490
         TabIndex        =   49
         Top             =   5790
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Charges (Add) :"
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
         Index           =   2
         Left            =   5430
         TabIndex        =   48
         Top             =   5790
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item Group : "
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
         Height          =   255
         Left            =   180
         TabIndex        =   47
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Currency"
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
         Height          =   255
         Left            =   150
         TabIndex        =   39
         Top             =   5730
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
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
         Height          =   255
         Left            =   1020
         TabIndex        =   38
         Top             =   5730
         Width           =   195
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Amount : "
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
         Left            =   2310
         TabIndex        =   37
         Top             =   1890
         Width           =   1290
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unit Price : "
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
         Height          =   255
         Left            =   195
         TabIndex        =   36
         Top             =   1890
         Width           =   945
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P.O. Qty : "
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
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label lblSize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PR NO.: "
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
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Description : "
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
         Height          =   255
         Index           =   1
         Left            =   8385
         TabIndex        =   33
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label lblItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name of Item :"
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
         Height          =   255
         Left            =   2580
         TabIndex        =   32
         Top             =   750
         Width           =   1425
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Materials/Store Code: "
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
         Height          =   255
         Left            =   210
         TabIndex        =   31
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total MRR Value :"
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
         Height          =   182
         Index           =   0
         Left            =   8281
         TabIndex        =   30
         Top             =   5785
         Width           =   1417
      End
   End
   Begin VB.TextBox txtEditMRRid 
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
      Left            =   1185
      MaxLength       =   30
      TabIndex        =   28
      Top             =   9225
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   -15
      TabIndex        =   65
      Top             =   8985
      Width           =   15240
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   -30
      TabIndex        =   64
      Top             =   8925
      Width           =   15240
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
      Index           =   1
      Left            =   0
      TabIndex        =   62
      Top             =   855
      Width           =   15240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MRR Information"
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
      Left            =   90
      TabIndex        =   61
      Top             =   210
      Width           =   5400
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   15240
   End
End
Attribute VB_Name = "frmMrrReceiveChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intmrrid As Double
Dim sl As String
Dim intYear As Integer
Dim intMRR As Integer
Dim strMRR As String
Dim strGroupNameShort As String
Dim boEdit As Boolean, bytRowNo As Byte
Dim addslno As Integer
Dim MRRPCID As Integer
Dim MRRUSERID As Integer
Dim strMRRID As String
Dim TempMRRVal As Currency
Dim tempRecQtyEdit As Currency
Dim ITEM1 As String
Dim DES1 As String
Dim Varpo1 As String
Dim PRNo1 As String
Dim CID As Integer

Private Sub ChkShow_Click()
If ChkShow.Value = 1 Then
cmbMRRID.Visible = True
txtMRRID.Visible = False
txtMRRID.Text = ""
cmbMRRID.Top = txtMRRID.Top
If Len(cmbSupplier) = 0 Then Exit Sub
    Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
    Call addPINO(cmbPINO, cmbSupplier.Text)
End If
If ChkShow.Value = 0 Then
cmbMRRID.Visible = False
txtMRRID.Visible = True
cmbMRRID.Text = ""
vmrrid.Text = ""
End If
End Sub

Private Sub cmbDescription_Click()

    cmbprno.ListIndex = -1
    txtMRRQty.Text = ""
    txtReceivedQty.Text = ""
    txtShortExcess.Text = ""
    cmbUnit.ListIndex = -1
    txtUnitPrice.Text = ""
    txtAmount.Text = ""
    txtRemarks.Text = ""

    DES1 = cmbDescription
End Sub

Private Sub cmbDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbItem_Click()
    txtDesc.Text = ""
    txtmatcode.Text = ""
    cmbprno.ListIndex = -1
    txtMRRQty.Text = ""
    txtReceivedQty.Text = ""
    txtShortExcess.Text = ""
    cmbUnit.ListIndex = -1
    txtUnitPrice.Text = ""
    txtAmount.Text = ""
    txtRemarks.Text = ""
    cmbItem.Tag = ""
    
    If Len(cmbItem) = 0 Then Exit Sub
    cmbItem.Tag = cmbItem.ItemData(cmbItem.ListIndex)
    
    Set R = New ADODB.Recordset
'    strRecord = "SELECT ItDesc,MATCODE FROM dbo.Pro_Inv_Sub where piID=" & findPIID(cmbPINO.Text) & " AND ItemID='" & findItemID(cmbItem) & "' AND DeleteRow= 0 ;"
  ''04-MAY-2016
     strRecord = " SELECT DISTINCT dbo.Pro_Inv_Sub.ItDesc " _
              & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID INNER JOIN dbo.new_Item_Name ON " _
              & " dbo.Pro_Inv_Sub.itemID = dbo.new_Item_Name.itemID INNER JOIN dbo.Pro_Inv_Main ON dbo.New_PI.PIID = dbo.Pro_Inv_Main.PIID INNER JOIN " _
              & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID WHERE dbo.Pro_Inv_Sub.DeleteRow=0 AND dbo.new_Supplier.suppName='" & cmbSupplier & "' AND  New_PI.PINO='" & cmbPINO & "' AND Pro_Inv_Sub.ItemID = " & cmbItem.Tag & " "
    R.Open strRecord, MainConn, adOpenStatic
    cmbDescription.Clear
    If R.RecordCount > 0 Then
        While Not R.EOF
        'txtDesc.Text = R![ItDesc]
        'txtmatcode.Text = R![MatCode]
        cmbDescription.AddItem R![ItDesc]
        R.MoveNext
        Wend
    End If
  Set R = Nothing
    ITEM1 = cmbItem
End Sub

Private Sub cmbItemType_Click()
    If Len(cmbItemType) = 0 Then Exit Sub

    ''Call addItemName(cmbItem, cmbItemType.Text)
    cmbItem.Clear
    Set R1 = New ADODB.Recordset
    strRecord = "SELECT DISTINCT Pro_Inv_Sub.itemID,dbo.new_Item_Name.itemName" _
              & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID INNER JOIN dbo.new_Item_Name ON " _
              & " dbo.Pro_Inv_Sub.itemID = dbo.new_Item_Name.itemID INNER JOIN dbo.Pro_Inv_Main ON dbo.New_PI.PIID = dbo.Pro_Inv_Main.PIID INNER JOIN " _
              & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID" _
              & " WHERE dbo.new_Supplier.suppName='" & cmbSupplier.Text & "' AND dbo.New_PI.PINo ='" & cmbPINO & "' AND dbo.new_Item_Name.itemType ='" & cmbItemType.Text & "' AND dbo.Pro_Inv_Sub.DeleteRow=0"
    R1.Open strRecord, MainConn, adOpenStatic
    If R1.RecordCount > 0 Then
        Do Until R1.EOF
            cmbItem.AddItem R1![itemName]
            cmbItem.ItemData(cmbItem.NewIndex) = R1![ITEMID]
        R1.MoveNext
        Loop
    End If
    R1.Close
    Set R1 = Nothing
End Sub

Private Sub cmbPINO_Click()
    
    'Call addPrNO(cmbprno, cmbPINO.Text)
    
    txtDesc.Text = ""
    txtmatcode.Text = ""
    cmbprno.ListIndex = -1
    cmbDescription.ListIndex = -1
    txtMRRQty.Text = ""
    txtReceivedQty.Text = ""
    txtShortExcess.Text = ""
    cmbUnit.ListIndex = -1
    txtUnitPrice.Text = ""
    txtAmount.Text = ""
    txtRemarks.Text = ""
    
'    Call addPOITEM(cmbItem, cmbPINO.Text)
'    Call addPrNO(cmbprno, cmbPINO.Text)
    cmbItemType.Clear
    If Len(cmbPINO) = 0 Then Exit Sub
    
    Set R = New ADODB.Recordset
    
'    findPIID(cmbPINO.Text)
'       strRecord = "SELECT groupname FROM dbo.New_PI where pino='" & cmbPINO & "'  ;"
'        strRecord = "SELECT groupname FROM dbo.New_PI where piID='" & findPIID(cmbPINO.Text) & "'  ;"


''''''''''''''''''''''''GROUP 04-MAY-2016
cmbItemType.Clear

    strRecord = "SELECT dbo.New_PI.groupname FROM dbo.Pro_Inv_Main INNER JOIN" _
                      & " dbo.New_PI ON dbo.Pro_Inv_Main.PIID = dbo.New_PI.PIID AND dbo.Pro_Inv_Main.compID = dbo.New_PI.ComID INNER JOIN" _
                      & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID where dbo.Pro_Inv_Main.DeleteRow=0 AND dbo.new_Supplier.suppName='" & cmbSupplier & "' AND pino='" & cmbPINO & "';"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      Do Until R.EOF
        cmbItemType.AddItem R![GroupName]
        R.MoveNext
      Loop
    End If
    R.Close
    
    flexItem.Rows = 1
    strRecord = "SELECT DISTINCT B.itemID,C.itemType,C.itemName,B.ItDesc,B.PRNo,B.PIQty,B.QtyUnit,B.Currency,B.UnitPrice,B.ItTotal FROM New_PI A INNER JOIN Pro_Inv_Sub B ON A.PIID = B.PIID" _
              & " INNER JOIN new_Item_Name C ON B.itemID = C.itemID INNER JOIN Pro_Inv_Main D ON A.PIID = D.PIID INNER JOIN new_Supplier E ON D.suppID = E.supplierID" _
              & " WHERE E.suppName = '" & cmbSupplier.Text & "' AND A.PINo = '" & cmbPINO.Text & "' AND B.DeleteRow = 0 ORDER BY C.itemType,C.itemName,B.PRNo;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
      With flexItem
        For inti = 1 To R.RecordCount
          .Rows = .Rows + 1
          .Row = .Rows - 1
          .Col = 0:   .Text = R![ITEMID]
          .Col = 1:   .Text = R![itemType]
          .Col = 2:   .Text = R![itemName]
          .Col = 3:   .Text = R![ItDesc]
          .Col = 4:   .Text = R![PRNo]
          .Col = 5:   .Text = R![PIQTY]
          .Col = 6:   .Text = R![QtyUnit]
          .Col = 7:   .Text = Format(R![UnitPrice], "##,##0.00")
          .Col = 8:   .Text = Format(R![ItTotal], "##,##0.00")
          .Col = 9:   .Text = R![Currency]
          R.MoveNext
        Next inti
        .Row = 0
      End With
    End If
    R.Close
    
    cmbprno.Clear
    
    Set R2 = New ADODB.Recordset
    
    strRecord = " SELECT DISTINCT dbo.Pro_Inv_Sub.PRNo " _
              & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID INNER JOIN dbo.new_Item_Name ON " _
              & " dbo.Pro_Inv_Sub.itemID = dbo.new_Item_Name.itemID INNER JOIN dbo.Pro_Inv_Main ON dbo.New_PI.PIID = dbo.Pro_Inv_Main.PIID INNER JOIN " _
              & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID WHERE dbo.new_Supplier.suppName='" & cmbSupplier & "' AND  dbo.New_PI.PINo ='" & cmbPINO & "' AND dbo.Pro_Inv_Sub.DeleteRow=0"
    
    R2.Open strRecord, MainConn, adOpenStatic
    If R2.RecordCount > 0 Then
        Do Until R2.EOF
            cmbprno.AddItem R2![PRNo]
        R2.MoveNext
        Loop
    End If
    R2.Close
    Set R = Nothing
    Set R2 = Nothing

    Varpo1 = cmbPINO
    
    If flexItem.Rows > 1 Then cmbItemType.ListIndex = 0

End Sub

Private Sub cmbPINO_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub cmbCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
   Call prcMoveTab(KeyCode)
End Sub
Private Sub cmbItemType_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbMainPF_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbMrrID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbMRRID_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub cmbMRRID_LostFocus()
    cmbMRRID.Text = Trim(UCase(cmbMRRID.Text))
End Sub

Private Sub cmbPINO_LostFocus()
  cmbPINO.Tag = findPIID(cmbPINO.Text)
End Sub

Private Sub cmbprno_Click()
    txtMRRQty.Text = ""
    cmbUnit.Text = ""
    txtUnitPrice.Text = ""
    txtReceivedQty.Text = ""
    txtAmount.Text = ""
    txtShortExcess.Text = ""
    
    PRNo1 = cmbprno
    
    Set R = New ADODB.Recordset
    
'        strRecord = "SELECT DISTINCT PIQTY,QTYUNIT,UNITPRICE,CURRENCY FROM dbo.Pro_Inv_Sub where piID=" & findPIID(cmbPINO.Text) & " AND PRNO='" & cmbprno.Text & "' AND ItemID='" & findItemID(cmbItem) & "' AND ItDesc= '" & cmbDescription.Text & "' And DeleteRow=0  ;"
        ''04-MAY-2016
         strRecord = " SELECT DISTINCT dbo.Pro_Inv_Sub.PIQTY,dbo.Pro_Inv_Sub.QTYUNIT,dbo.Pro_Inv_Sub.UNITPRICE,dbo.Pro_Inv_Sub.CURRENCY " _
              & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID INNER JOIN dbo.new_Item_Name ON " _
              & " dbo.Pro_Inv_Sub.itemID = dbo.new_Item_Name.itemID INNER JOIN dbo.Pro_Inv_Main ON dbo.New_PI.PIID = dbo.Pro_Inv_Main.PIID INNER JOIN " _
              & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID where dbo.Pro_Inv_Sub.DeleteRow=0 And  dbo.new_Supplier.suppName='" & cmbSupplier & "' AND  New_PI.PINO='" & cmbPINO & "'  AND Pro_Inv_Sub.ItemID='" & findItemID(cmbItem) & "' AND ItDesc= '" & cmbDescription.Text & "'  AND PRNO='" & cmbprno.Text & "' "

        
        
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        cmbUnit.Text = R![QtyUnit]
        cmbCurrency.Text = R![Currency]
      If R.RecordCount > 1 Then
       Do Until R.EOF
            txtMRRQty.AddItem R![PIQTY]
            txtMRRQty.ItemData(txtMRRQty.NewIndex) = R![UnitPrice]
        R.MoveNext
        Loop
       Else
        txtMRRQty.Text = R![PIQTY]
        txtUnitPrice.Text = R![UnitPrice]
       End If
    End If
    Set R = Nothing
End Sub

Private Sub cmbSupplier_Click()
    If Len(cmbSupplier) = 0 Then Exit Sub
    Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
    Call addPINO(cmbPINO, cmbSupplier.Text)
End Sub

Private Sub cmbMrrID_Click()
    If Len(cmbMRRID) = 0 Then Exit Sub
    cmbMRRID.Tag = cmbMRRID.ItemData(cmbMRRID.ListIndex)
    Call prcAddSuppMrrno(vmrrid, cmbMRRID.Text, cmbSupplier.Text)
End Sub

Private Sub cmbSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAdd_Click()
Dim ITEM As String
Dim DES As String
Dim VarPO As String, PRNo As String
Dim inti As Integer
MRRPCID = getComVolID
MRRUSERID = bytUserID
    On Error GoTo x
    If Len(txtChallanNo.Text) = 0 Then
        MsgBox "Type Challan No", vbInformation, cnstMsgInfo
        txtChallanNo.SetFocus
        Exit Sub
    End If
    If Len(txtChallanDate.Text) = 0 Then
        MsgBox "Type Challan Date", vbInformation, cnstMsgInfo
        txtChallanDate.SetFocus
        Exit Sub
    End If
    If Len(cmbMainPF.Text) = 0 Then
        MsgBox "Type Main PF No", vbInformation, cnstMsgInfo
        cmbMainPF.SetFocus
        Exit Sub
    End If
    If Len(cmbprno.Text) = 0 Then
        MsgBox "Type PR No", vbInformation, cnstMsgInfo
        cmbprno.SetFocus
        Exit Sub
    End If
    If Len(cmbPINO.Text) = 0 Then
        MsgBox "Select PI NO", vbInformation, cnstMsgInfo
        cmbPINO.SetFocus
        Exit Sub
    End If
    If Len(txtCont.Text) = 0 Then
        MsgBox "Enter Container No", vbInformation, cnstMsgInfo
        txtCont.SetFocus
        Exit Sub
    End If

    If Len(cmbUnit.Text) = 0 Then
        MsgBox "Enter UNIT", vbInformation, cnstMsgInfo
        cmbUnit.SetFocus
        Exit Sub
    End If
    
    If Len(cmbItem.Text) = 0 Then
        MsgBox "Select Name of Item", vbInformation, cnstMsgInfo
        cmbItem.SetFocus
        Exit Sub
    End If
    If Len(txtMRRQty.Text) = 0 Then
        MsgBox "Type Challan Qty", vbInformation, cnstMsgInfo
        txtMRRQty.SetFocus
        Exit Sub
    End If

    If Len(txtUnitPrice.Text) = 0 Then
        MsgBox "Type Unit Price", vbInformation, cnstMsgInfo
        txtUnitPrice.SetFocus
        Exit Sub
    End If

    strRecord = "SELECT PFNo FROM new_PF_No WHERE PFNo='" & cmbMainPF.Text & "';"
    MainComm.CommandText = strRecord
    Set R = MainComm.Execute
    If R.EOF = True And R.BOF = True Then
        MsgBox "Actual PF not found", vbInformation, cnstMsgInfo
        cmbMainPF.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If cmdAdd.Caption = "&Add" Then
    
     With flexDetails
                
                For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 7:  ITEM = .Text
                .Col = 8:  DES = .Text
                .Col = 4:  VarPO = .Text
                .Col = 10:  PRNo = .Text
                If ITEM1 = ITEM And DES1 = DES And Varpo1 = VarPO And PRNo1 = PRNo Then
                MsgBox "Sorry, already add list"
                Exit Sub
                End If
                Next inti
    End With
    End If
    strMRRID = vmrrid.Text
    With flexDetails
        If cmdAdd.Caption = "&Add" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            If txtMode.Text = "NEW" Then
                addslno = .Rows - 1
            Else
                Call prCMAX(vmrrid)
                addslno = intSLNO + 1
                bytRowNo = .Row
            End If
        Else
             .Row = bytRowNo
        End If
    
        .Col = 0:  .Text = txtChallanNo.Text
        .Col = 1:  .Text = txtChallanDate.Text
        .Col = 2:  .Text = cmbMainPF.Text
        .Col = 3:  .Text = cmbPINO.Tag
        .Col = 4:  .Text = cmbPINO.Text
        .Col = 5:  .Text = txtCont.Text
        .Col = 6:  .Text = cmbItemType.Text
        .Col = 7:  .Text = cmbItem.Text
        
        .Col = 8:  .Text = cmbDescription.Text & ""
        
        .Col = 9:  .Text = txtmatcode.Text
        .Col = 10:  .Text = cmbprno.Text
        .Col = 11:  .Text = txtMRRQty.Text
        .Col = 12:  .Text = txtReceivedQty.Text
        .Col = 13:  .Text = txtShortExcess.Text
        .Col = 14:  .Text = cmbUnit.Text
        .Col = 15:  .Text = txtUnitPrice.Text
        .Col = 16:  .Text = txtAmount.Text
        .Col = 17:  .Text = txtRemarks.Text
        .Col = 18:  .Text = addslno
        .Col = 19:  .Text = MRRPCID
        .Col = 20:  .Text = MRRUSERID
        .Col = 21:  .Text = txtDesc.Text & ""
            
            If cmdAdd.Caption <> "&Add" Then
                Call prCMAX(vmrrid)
                MainComm.CommandText = "UPDATE MRR_Challan_Sub SET DELETEROW = 2, Pcid=" & getComVolID & ",UserId=" & bytUserID & ",SLNO= " & intSLNO + 1 & ",EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE SLNO=" & addslno & " AND MRRID=" & vmrrid & " AND Deleterow=0;"
                MainComm.Execute
            End If
        
    End With
    If txtMode.Text <> "NEW" Then
        Call prcSaveGriddate(strMRRID)
    End If
    
    boEdit = False
    txtMRRValue.Text = Format(TotalMRRValue, "##,##0.00")
 
    cmbprno.Text = ""
    txtMRRQty.Text = "0"
    txtReceivedQty.Text = "0"
    txtShortExcess.Text = "0"
    cmbUnit.ListIndex = -1
    txtUnitPrice.Text = "0.00"
    txtAmount.Text = "0.00"
    txtRemarks.Text = ""
    
    cmbItem.SetFocus

    Exit Sub
x:
    MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub

Private Sub cmdAddNewItem_Click()
    strMark = "ITEM"
    frmNewItemName.Show 1
End Sub

Private Sub cmdClose_Click()
    If MsgBox("Are you sure to Close", vbInformation + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo x
    chkTextNull
    If boEmptyText = True Then Exit Sub
    If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn

        MainConn.BeginTrans
         
            MainComm.CommandText = "UPDATE New_MRR SET DeleteRow =1,EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE DeleteRow =0 AND MRRID='" & vmrrid & "'And ComId=" & bytCompID & ";"
            MainComm.Execute
            MainComm.CommandText = "UPDATE MRR_Challan_Main SET DeleteRow = 1,Pcid=" & getComVolID & ",Userid=" & bytUserID & ",EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE DeleteRow =0 AND MRRID='" & vmrrid & "'And CompId=" & bytCompID & ";"
            MainComm.Execute
            MainComm.CommandText = "UPDATE MRR_Challan_Sub SET DeleteRow = 1,Pcid=" & getComVolID & ",Userid=" & bytUserID & ",EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE DeleteRow = 0 AND MRRID='" & vmrrid & "';"
            MainComm.Execute
            MainComm.CommandText = "UPDATE Signature SET DeleteRow = 1,Pcid=" & getComVolID & ",Userid=" & bytUserID & ",EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE DeleteRow = 0 AND MRRID='" & vmrrid & "';"
            MainComm.Execute
    
        MainConn.CommitTrans
        
    Screen.MousePointer = vbDefault
    
    cmdRefresh_Click
    cmbMRRID.Text = ""
    txtMrrDate.Text = ""
    txtMRRValue.Text = ""
    cmbMainPF.ListIndex = -1
    cmbPINO.ListIndex = -1
    
    Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
        cmbMRRID.SetFocus
    Exit Sub
x:
    Screen.MousePointer = vbDefault
    MainConn.RollbackTrans
    
    If err.Number = -2147467259 Then
        MsgBox "This record cannot be deleted", vbCritical, cnstMsgErDB
    Else
        MsgBox err.Description, vbCritical, cnstMsgErDB
    End If
End Sub

Private Sub cmdEdit_Click()
    If txtMode.Text = "NEW" Then
        cmdSave.Caption = "&Update"
        cmdSave.Enabled = True
        cmdDelete.Enabled = True
        txtMRRID.Visible = False
        cmbMRRID.Top = txtMRRID.Top
        cmbMRRID.Visible = True
        cmdShow.Top = cmbMRRID.Top
        cmdShow.Visible = True
        cmdEdit.Caption = "&New"
        txtMode.Text = "EDIT"
        Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
        Call DisableSignature
    Else
        cmdSave.Caption = "&Save"
        cmdSave.Enabled = True
        cmdDelete.Enabled = False
        cmbMRRID.Visible = False
        cmdShow.Visible = False
        txtMRRID.Visible = True
        cmdEdit.Caption = "&Edit"
        txtMode.Text = "NEW"
        Call EnableSignature
    End If
    
    If txtMode.Text = "NEW" Then txtMRRID.Text = "" Else cmbMRRID.Text = ""
    txtMRRValue.Text = ""
    txtMrrDate.Text = ""
    
    txtChallanNo.Text = ""
    txtChallanDate.Text = ""
    cmbMainPF.ListIndex = -1
    cmbPINO.ListIndex = -1
    txtCont.Text = ""
    cmbItem.ListIndex = -1
    txtDesc.Text = ""
    txtmatcode.Text = ""
    cmbprno.Text = ""
    txtMRRQty.Text = ""
    cmbUnit.ListIndex = -1
    txtUnitPrice.Text = "0.00"
    txtCharges.Text = "0.00"
    txtAmount.Text = "0.00"
    flexDetails.Rows = 1
    cmbCurrency.Text = findDCurrency
    cmbSupplier.SetFocus
End Sub

Private Sub cmdNewMainPF_Click()
    strMark = "PF"
    frmNewPFNo.Show 1
End Sub

Private Sub cmdNewSupplier_Click()
    strMark = "SUPP"
    frmNewSupplier.Show 1
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo x

Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
        MainConn.BeginTrans
        MainComm.CommandText = "DELETE FROM TEMPSUM "
        MainComm.Execute
        
        Set R = New ADODB.Recordset
        Set RS = New ADODB.Recordset
        strRecord = "SELECT * FROM MRR_Challan_Sub WHERE (dbo.MRR_Challan_Sub.DeleteRow = 0) AND (dbo.MRR_Challan_Sub.MRRID = '" & vmrrid.Text & "')"

   R.Open strRecord, MainConn, adOpenStatic
    Do Until R.EOF
    
   strRecord = "SELECT SUM(dbo.MRR_Challan_Sub.RcvdQty) AS Sum,dbo.MRR_Challan_Sub.PIId, dbo.MRR_Challan_Sub.ItDesc, dbo.MRR_Challan_Sub.ItemId, dbo.MRR_Challan_Sub.PRNO FROM dbo.MRR_Challan_Sub LEFT OUTER JOIN" _
                & " dbo.MRR_Challan_Main INNER JOIN  dbo.New_MRR ON dbo.MRR_Challan_Main.MRRID = dbo.New_MRR.MRRID" _
                & " AND dbo.MRR_Challan_Main.compID = dbo.New_MRR.ComID ON dbo.MRR_Challan_Sub.MRRID = dbo.MRR_Challan_Main.MRRID" _
                & " WHERE (dbo.MRR_Challan_Sub.DeleteRow = 0) AND (dbo.MRR_Challan_Sub.MRRID <= '" & vmrrid.Text & "') AND (dbo.MRR_Challan_Sub.PIId = '" & R![PIId] & "') AND dbo.MRR_Challan_Sub.ItDesc='" & R![ItDesc] & "'AND dbo.MRR_Challan_Sub.ItemId='" & R![ITEMID] & "' AND dbo.MRR_Challan_Sub.PRNO='" & R![PRNo] & "'" _
                & " GROUP BY dbo.MRR_Challan_Sub.PIId,dbo.MRR_Challan_Sub.ItDesc, dbo.MRR_Challan_Sub.ItemId, dbo.MRR_Challan_Sub.PRNO"
                
                RS.Open strRecord, MainConn, adOpenStatic
    
        MainComm.CommandText = "INSERT INTO TEMPSUM  (MRRID,PIID,ITDESC,SUM,ITEMID,PRNO) VALUES ('" & vmrrid.Text & "','" & RS![PIId] & "','" & RS![ItDesc] & "','" & RS![Sum] & "','" & RS![ITEMID] & "','" & RS![PRNo] & "')"
        MainComm.Execute
    RS.Close
    R.MoveNext
    Loop
    MainConn.CommitTrans
    
    Set R = Nothing
    Set RS = Nothing
    
    Set R1 = New ADODB.Recordset
    Set REPORT = Appl.OpenReport(findReportPath & "\rptMrrChallan_MrrNo11.rpt")
    strRecord = "Select * from VIEW5 Where DELETEROW=0  AND compid = " & bytCompID & " and MRRID = '" & vmrrid.Text & "' order by PRNO,ItemType, ItemName, MatCode;"
    
    R1.Open strRecord, MainConn, 3, 3
    
    REPORT.DiscardSavedData
    REPORT.Database.SetDataSource R1
    frmMainReport.CRVIEWER.Refresh
    frmMainReport.CRVIEWER.ReportSource = REPORT
    frmMainReport.CRVIEWER.ViewReport
    frmMainReport.CRVIEWER.Zoom (82)
    frmMainReport.Show 1
    R1.Close
    Set R1 = Nothing

    Exit Sub
x:
    MsgBox err.Description, vbCritical, App.Title
End Sub

Private Sub cmdRefresh_Click()
    cmbItemType.ListIndex = -1
    cmbItem.ListIndex = -1
    txtDesc.Text = ""
    txtmatcode.Text = ""
    cmbprno.Text = ""
    txtMRRQty.Text = "0.00"
    cmbUnit.ListIndex = -1
    txtUnitPrice.Text = "0.00"
    txtAmount.Text = "0.00"
    txtMrrDate.Text = ""
    txtChallanNo.Text = ""
    txtChallanDate.Text = ""
    cmbMainPF.ListIndex = -1
    cmbPINO.ListIndex = -1
    txtCont.Text = ""
    txtCharges.Text = "0.00"
    txtDiscount.Text = "0.00"
    txtMRRValue.Text = "0.00"
    flexDetails.Rows = 1
    cmbDescription.Clear
    Empcode1.Text = ""
    Empcode2.Text = ""
    Empcode3.Text = ""
    Empcode4.Text = ""
    Empcode5.Text = ""
    Name1.Text = ""
    Name2.Text = ""
    Name3.Text = ""
    Name4.Text = ""
    Name5.Text = ""
    
End Sub

Private Sub cmdReload_Click()

If Len(cmbPINO.Text) = 0 Then
  MsgBox "Select PO No.", vbInformation, cnstMsgInfo
  cmbPINO.SetFocus
  Exit Sub
End If

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

MainComm.CommandText = "UPDATE A SET A.itemID = B.itemID FROM Pro_Inv_Sub A INNER JOIN (SELECT DISTINCT A.PIID,A.PINo,A.GroupName,A.itemID AS PIItemID,B.itemID,B.itemName,B.itemType" _
      & " FROM (SELECT A.PIID,A.GroupName,A.PINo,B.itemID,C.itemName,C.itemType,B.ItDesc FROM New_PI A INNER JOIN Pro_Inv_Sub B ON A.PIID = B.PIID INNER JOIN new_Item_Name C ON B.itemID = C.itemID" _
      & " WHERE A.GroupName <> C.itemType AND A.EntryDate > '31-MAR-2017') A INNER JOIN new_Item_Name B ON A.itemName = B.itemName AND A.itemID <> B.itemID WHERE A.PINo = '" & cmbPINO.Text & "') B ON A.PIID = B.PIID;"
MainComm.Execute

MainConn.CommitTrans

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  MainConn.RollbackTrans
  Screen.MousePointer = vbDefault
  err.Number = 0
  Exit Sub
End If

Screen.MousePointer = vbDefault
MsgBox "PI Item Reloded Successfully", vbInformation, cnstMsgInfo

End Sub

Private Sub cmdSave_Click()
    On Error GoTo x
    
    
       Screen.MousePointer = vbHourglass
        Set MainComm = New ADODB.Command
        MainComm.ActiveConnection = MainConn
    
    Call chkTextNull
    
    
     MainConn.BeginTrans
    
    
     If flexDetails.Row < 1 Then
     
        MsgBox "No Data found !", vbInformation, cnstMsgInfo
        cmbItemType.SetFocus
        Exit Sub
        
    End If
    
    

    If boEmptyText = True Then Exit Sub
    
    
    
    If txtMode.Text = "NEW" Then
        
        
        If MsgBox("Are you sure to Save?", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
          
          
          If txtMode.Text = "NEW" Then
          
          
        
                intYear = Year(txtMrrDate.Text)
                    
                    Set R = New ADODB.Recordset
                    
                    
                            strRecord = "SELECT itemShortName FROM new_Item_Type Where itemType = '" & cmbItemType & "';"
                            
                            R.Open strRecord, MainConn, adOpenStatic
                            
                            
                            
                                If R.EOF = False And R.BOF = False Then
                                    strGroupNameShort = R.Fields(0)
                                Else
                                    strGroupNameShort = "NF"
                                
                                End If
                        
                        R.Close
                
            Set R = Nothing
            
            On Error GoTo er1
er1:
            If err.Number = 94 Then
            
                Set R = Nothing
                
                intMRR = 1
                
            Else
            
                'AND ItemID='" & findItemID(cmbItem) & "'
                intmrrid = fncNewID("MRR")
                
                
                
                Set R = New ADODB.Recordset
                
                
                    strRecord = "SELECT Max(MRR_No) As MaxMRRNO,max(MRRID) as MRRID FROM New_MRR Where ComID = " & bytCompID & " And GroupName = '" & cmbItemType & "' And MRR_Year = " & intYear & ";"
                    R.Open strRecord, MainConn, adOpenStatic
                    
                    
                        If R.EOF = False And R.BOF = False Then
                            intMRR = R.Fields(0) + 1
                            'intmrrid = R.Fields(1) + 1
                        Else
                            intMRR = 1
                        End If
                        
                        
                        R.Close
                    
                Set R = Nothing
                
            End If
            
            
            sl = fncMRR(intMRR)
            
            
            
            txtMRRID.Text = strGroupNameShort & bytCompID & fncMRR(intMRR) & "-" & intYear
        
        End If
        
        strMRRID = intmrrid
        
'        Call chkTextNull
    Else
    
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        strMRRID = vmrrid.Text
    
    
    End If
'         Screen.MousePointer = vbHourglass
'        Set MainComm = New ADODB.Command
'        MainComm.ActiveConnection = MainConn
'MainConn.BeginTrans
   
        If txtMode.Text = "NEW" Then
         
            MainComm.CommandText = "INSERT INTO Mrr_Challan_Main(compID, suppID, MRRID, MRRDate, MRRCurr, MRRValue, MRRCharge, MRRDiscount,PCID,USERID,EntryDate,EntryTime)" _
                  & " VALUES (" & bytCompID & "," & findSupplierID(cmbSupplier.Text) & ",'" & strMRRID & "','" & Format(txtMrrDate.Text, cnstDtFrmtQ) & "'," _
                  & " '" & cmbCurrency.Text & "'," & CCur(txtMRRValue.Text) & "," & CCur(txtCharges.Text) & "," & CCur(txtDiscount.Text) & "," & getComVolID & "," & bytUserID & ",'" & Date & "','" & Time & "');"
            MainComm.Execute
            
            MainComm.CommandText = "INSERT INTO New_MRR(mrrid,comID, Groupname, MRR_No, MRR_year,MRRNO,EntryDate,EntryTime)" _
                  & " VALUES (" & strMRRID & "," & bytCompID & ",'" & cmbItemType & "','" & sl & "','" & intYear & "','" & txtMRRID & "','" & Date & "','" & Time & "');"
            MainComm.Execute

            MainComm.CommandText = "INSERT INTO Signature(MRRID,QualityID,PreparedID,QuantityID,ManagerID,SRManagerID,PCID,USERID,EntryDate,EntryTime)" _
                  & " VALUES (" & strMRRID & "," & findempID(CID, Empcode1, Name1) & "," & findempID(CID, Empcode2, Name2) & "," & findempID(CID, Empcode3, Name3) & "," & findempID(CID, Empcode4, Name4) & "," & findempID(CID, Empcode5, Name5) & "," & getComVolID & "," & bytUserID & ",'" & Date & "','" & Time & "');"
            MainComm.Execute
            
            Call prcSaveGriddate(strMRRID)
            
        End If
        
    MainConn.CommitTrans
    
    MsgBox "Transaction update successfully completed. MRR No = " + txtMRRID.Text, vbInformation + vbOKOnly, Me.Caption
    cmdRefresh_Click
    
    If txtMode.Text = "NEW" Then
        txtMRRID.Text = ""
        'txtMRRID.SetFocus
    Else
        cmbMRRID.Text = ""
        cmbMainPF.ListIndex = -1
        Call prcAddSuppMrrID(cmbMRRID, cmbSupplier.Text)
        cmbMRRID.SetFocus
        
    End If
    
    'Screen.MousePointer = vbDefault
    Set R = Nothing
   
    Screen.MousePointer = vbDefault
    Exit Sub
x:
    Screen.MousePointer = vbDefault
    MainConn.RollbackTrans
    MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub

Private Sub cmdShow_Click()
    cmdRefresh_Click
    strMRRID = vmrrid
    
    On Error GoTo x
    If Len(cmbMRRID.Text) = 0 Then
        MsgBox "Select MRR No", vbInformation, cnstMsgInfo
        cmbMRRID.SetFocus
        Exit Sub
    End If
    flexDetails.Rows = 1
    Set RS = New ADODB.Recordset
        strRecord = " SELECT  New_MRR.MRRID, QS.Empname AS EmpNameQua, QS.Empcode AS EmpcodeQua, PS.Empname AS EmpNamePre, PS.Empcode AS EmpCodePre, QNS.Empname AS EmpnameQN, QNS.Empcode AS EmpcodeQN, MS.Empname AS EmpNameM,MS.Empcode AS EmpcodeM, SRMS.Empname AS EmpNameSRM, SRMS.Empcode AS EmpCodeSRM FROM Signature INNER JOIN InfoEmp QS ON Signature.QualityID = QS.EMPID INNER JOIN InfoEmp PS ON Signature.PreparedID = PS.EMPID INNER JOIN InfoEmp QNS ON Signature.QuantityID = QNS.EMPID INNER JOIN InfoEmp SRMS ON Signature.SRManagerID = SRMS.EMPID INNER JOIN InfoEmp MS ON Signature.ManagerID = MS.EMPID INNER JOIN New_MRR ON Signature.MRRID = New_MRR.MRRID INNER JOIN Company_Information ON New_MRR.ComID = Company_Information.CompID " _
        & " WHERE (dbo.New_MRR.ComID = " & bytCompID & ") AND (dbo.Signature.MRRID = '" & vmrrid.Text & "') GROUP BY New_MRR.MRRID, QS.Empname, PS.Empname, QNS.Empname, MS.Empname, SRMS.Empname, QS.Empcode , PS.Empcode, QNS.Empcode, MS.Empcode, SRMS.Empcode, QS.Empname, QS.Empcode, PS.Empname, PS.Empcode, QNS.Empname, QNS.Empcode, PS.Empname, PS.Empcode, QNS.Empname, QNS.Empcode, MS.Empname, MS.Empcode, SRMS.Empname, SRMS.Empcode;"
    
    MainComm.CommandText = strRecord
    Set RS = MainComm.Execute
    
    If RS.EOF = False And RS.BOF = False Then
        Empcode1.Text = RS!EmpCodeQua: Name1.Text = RS!EmpNameQua
        Empcode2.Text = RS!EmpCodePre: Name2.Text = RS!EmpNamePre
        Empcode3.Text = RS!EmpCodeQN: Name3.Text = RS!EmpNameQN
        Empcode4.Text = RS!EmpCodeM: Name4.Text = RS!EmpNameM
        Empcode5.Text = RS!EmpCodeSRM: Name5.Text = RS!EmpNameSRM
    End If
    
    Set R = New ADODB.Recordset
    
    strRecord = " SELECT MRR_Challan_Main.compID, Company_Information.compName, Company_Information.compAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
        " FROM ((Company_Information INNER JOIN MRR_Challan_Main ON Company_Information.compID = MRR_Challan_Main.compID) INNER JOIN new_Supplier ON MRR_Challan_Main.suppID = new_Supplier.supplierID) INNER JOIN (new_Item_Type INNER JOIN ((MRR_Challan_Sub INNER JOIN new_Item_Name ON MRR_Challan_Sub.itemID = new_Item_Name.itemID) INNER JOIN New_PI ON MRR_Challan_Sub.PIID = New_PI.PIID) ON new_Item_Type.itemType = new_Item_Name.itemType) ON MRR_Challan_Main.MRRID = MRR_Challan_Sub.MRRID " & _
        " WHERE MRR_Challan_SUB.DELETEROW = 0 GROUP BY MRR_Challan_Main.compID, Company_Information.compName, Company_Information.compAdd, MRR_Challan_Main.suppID, new_Supplier.suppName, MRR_Challan_Main.MRRID, MRR_Challan_Main.MRRDate, MRR_Challan_Main.MRRCurr, MRR_Challan_Main.MRRValue, MRR_Challan_Sub.MainPF, New_PI.PIID, New_PI.PINo, MRR_Challan_Main.MRRCharge, MRR_Challan_Main.MRRDiscount, MRR_Challan_Sub.ChallanNo, MRR_Challan_Sub.ChallanDate, MRR_Challan_Sub.Containerno, MRR_Challan_Sub.itemID, new_Item_Name.itemName, new_Item_Name.itemType, new_Item_Type.itemGroup, MRR_Challan_Sub.ItDesc, MRR_Challan_Sub.MatCode, MRR_Challan_Sub.PRNO, MRR_Challan_Sub.MRRQty, MRR_Challan_Sub.RcvdQty, MRR_Challan_Sub.QtyUnit, MRR_Challan_Sub.UnitPrice, MRR_Challan_Sub.ItTotal, MRR_Challan_Sub.Remarks,MRR_Challan_Sub.SLNO,MRR_Challan_Sub.PCID,MRR_Challan_Sub.USERID,MRR_Challan_Sub.Remarks2 " & _
        " HAVING (((MRR_Challan_Main.compID)=" & bytCompID & ") AND ((MRR_Challan_Main.MRRID)='" & vmrrid.Text & "'));"

    MainComm.CommandText = strRecord
    Set R = MainComm.Execute
    
    If R.EOF = False And R.BOF = False Then
        
        txtEditMRRid.Text = R![MRRID]
        txtMrrDate.Text = Format(R![mrrdate], cnstDtFrmtI)
        txtDiscount.Text = Format(R![MRRDiscount], "##,##0.00")
        txtCharges.Text = Format(R![MRRCharge], "##,##0.00")
        cmbCurrency.Text = R![MRRCurr]

    With flexDetails
        Do Until R.EOF
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0: .Text = R![ChallanNo]
            .Col = 1: .Text = Format(R![ChallanDate], cnstDtFrmtI)
            .Col = 2: .Text = R![MainPF]
            .Col = 3: .Text = R![PIId]
            .Col = 4: .Text = R![PINo]
            .Col = 5: .Text = R![Containerno]
            .Col = 6: .Text = findItemType(R![itemName])
            .Col = 7: .Text = R![itemName]
            .Col = 9: If Not IsNull(R![MatCode]) Then .Text = R![MatCode]
            .Col = 10: If Not IsNull(R![PRNo]) Then .Text = R![PRNo]
            .Col = 11: .Text = R![MRRQty]
            .Col = 12: .Text = R![RcvdQty]
            .Col = 13: .Text = FormatNumber(R![MRRQty] - R![RcvdQty], 2)
            .Col = 14: .Text = R![QtyUnit]
            .Col = 15: .Text = Format(R![UnitPrice], "##,##0.00")
            .Col = 16: .Text = Format(R![ItTotal], "##,##0.00")
            .Col = 17:  If Not IsNull(R![Remarks]) Then .Text = R![Remarks]
            .Col = 18:  .Text = R![SLNo]
            .Col = 19:  .Text = R![PCID]
            .Col = 20:  .Text = R![Userid]
            .Col = 21:  If Not IsNull(R![Remarks2]) Then .Text = R![Remarks2]
            .Col = 8: If Not IsNull(R![ItDesc]) Then .Text = R![ItDesc]
        R.MoveNext
        Loop
    End With
        txtMRRValue.Text = Format(TotalMRRValue, "##,##0.00")
    Else
        MsgBox "Record Not Found", vbInformation, cnstMsgInfo
        cmdRefresh_Click
        cmbMainPF.ListIndex = -1
        cmbMRRID.SetFocus
        Exit Sub
    End If
    Set R = Nothing
    Set RS = Nothing
    Exit Sub
x:
    MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub

Private Sub Empcode1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Empcode1_LostFocus()
 Set R = New ADODB.Recordset
    strRecord = "SELECT empid,empName FROM infoemp where empcode= '" & Empcode1 & "' and compid=" & bytCompID & " and mleft=0 and  (wstatus='S'OR JSign=1);"
    R.Open strRecord, MainConn, adOpenStatic
    Name1 = ""
    If R.RecordCount > 0 Then
        Name1 = R![Empname]
    Else
    MsgBox "Sorry,Check Your Employee Code"

    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub Empcode2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Empcode2_LostFocus()
    Set R = New ADODB.Recordset
        strRecord = "SELECT empid,empName FROM infoemp where empcode= '" & Empcode2 & "' and compid=" & bytCompID & " and mleft=0 ;"
        R.Open strRecord, MainConn, adOpenStatic
        Name2 = ""
        If R.RecordCount > 0 Then
            Name2 = R![Empname]
        Else
            MsgBox "Sorry,Check Your Employee Code"
        End If
        R.Close
    Set R = Nothing
End Sub

Private Sub Empcode3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Empcode3_LostFocus()
    Set R = New ADODB.Recordset
        strRecord = "SELECT empid,empName FROM infoemp where empcode= '" & Empcode3 & "' and compid=" & bytCompID & " and mleft=0 ;"
        R.Open strRecord, MainConn, adOpenStatic
        Name3 = ""
        If R.RecordCount > 0 Then
            Name3 = R![Empname]
        Else
            MsgBox "Sorry,Check Your Employee Code"
        End If
        R.Close
    Set R = Nothing
End Sub

Private Sub Empcode4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Empcode4_LostFocus()
    Set R = New ADODB.Recordset
        strRecord = "SELECT empid,empName FROM infoemp where empcode= '" & Empcode4 & "' and compid=" & bytCompID & " and mleft=0 and  (wstatus='S'OR JSign=1);"
        R.Open strRecord, MainConn, adOpenStatic
        Name4 = ""
        If R.RecordCount > 0 Then
            Name4 = R![Empname]
        Else
            MsgBox "Sorry,Check Your Employee Code"
        End If
        R.Close
    Set R = Nothing
End Sub

Private Sub Empcode5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub Empcode5_LostFocus()
    Set R = New ADODB.Recordset
        strRecord = "SELECT empid,empName FROM infoemp where empcode= '" & Empcode5 & "' and compid=" & bytCompID & " and mleft=0 and  (wstatus='S'OR JSign=1);"
        R.Open strRecord, MainConn, adOpenStatic
        Name5 = ""
        If R.RecordCount > 0 Then
            Name5 = R![Empname]
        Else
            MsgBox "Sorry,Check Your Employee Code"
        End If
        R.Close
    Set R = Nothing
End Sub

Private Sub flexDetails_DblClick()
    Call flexDetails_KeyDown(13, 0)
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    ''On Error GoTo x
    On Error Resume Next
    cmdAdd.Caption = "&Add"
    If KeyCode = 13 Then
        cmdAdd.Caption = "&Replace"
        With flexDetails
            If .RowSel = 0 Then Exit Sub
            .Col = 0: txtChallanNo.Text = .Text
            .Col = 1: txtChallanDate.Text = .Text
            .Col = 2: cmbMainPF.Text = .Text
            .Col = 3: cmbPINO.Tag = .Text
            .Col = 4: cmbPINO.Text = .Text
            .Col = 5: txtCont.Text = .Text
            .Col = 6: cmbItemType.Text = .Text
            .Col = 7: cmbItem.Text = .Text
            .Col = 8: cmbDescription.Text = .Text
            .Col = 9: txtmatcode.Text = .Text
            .Col = 10: cmbprno.Text = .Text
            .Col = 11: txtMRRQty.Text = .Text
            .Col = 12: txtReceivedQty.Text = .Text
            .Col = 13: txtShortExcess.Text = .Text
            .Col = 14: cmbUnit.Text = .Text
            .Col = 15: txtUnitPrice.Text = .Text
            .Col = 16: txtAmount.Text = .Text
            .Col = 17: txtRemarks.Text = .Text
            .Col = 18: addslno = .Text
            .Col = 19: MRRPCID = .Text
            .Col = 20: MRRUSERID = .Text
            .Col = 21: txtDesc.Text = .Text
            bytRowNo = .Row
        End With
        tempRecQtyEdit = txtReceivedQty.Text
        boEdit = True
        cmbItem.SetFocus
    End If
    If KeyCode = vbKeyDelete Then
        With flexDetails
            If txtMode.Text = "NEW" Then
                If .RowSel = 0 Then
                    Beep
                Else
                    If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
                    If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
                    txtMRRValue.Text = Format(TotalMRRValue, "##0.00")
                End If
            Else
                If .RowSel = 0 Then
                        Beep
                Else
                    If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
                    'If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
                    .Col = 18: addslno = .Text
                    bytRowNo = .Row
                    'Call prCMAX(INTSLNO, vmrrid.Text)
                    'Call prCMAX(vmrrid)
                    MainComm.CommandText = "UPDATE MRR_Challan_Sub SET DELETEROW = 1 WHERE SLNO=" & addslno & " AND MRRID=" & vmrrid & ";"
                    MainComm.Execute
                    If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
                    txtMRRValue.Text = Format(TotalMRRValue, "##0.00")
                End If
            End If
        End With
    End If
    Exit Sub
x:
    MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub

Private Sub flexItem_DblClick()

If flexItem.Rows = 1 Then Exit Sub

With flexItem
  .Col = 0:  cmbItem.Tag = .Text
  .Col = 1:  cmbItemType.Text = .Text
  .Col = 2:  cmbItem.Text = .Text
  .Col = 3:  cmbDescription.Text = .Text
  .Col = 4:  cmbprno.Text = .Text
  .Col = 5:  txtMRRQty.Text = .Text
  .Col = 6:  cmbUnit.Text = .Text
  .Col = 7:  txtUnitPrice.Text = .Text
  .Col = 9:  cmbCurrency.Text = .Text
End With

End Sub

Private Sub flexItem_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then flexItem_DblClick
End Sub

Private Sub Form_Activate()
'    Call prcConReport(CrystalReport1)
    Me.Top = 1240
    If strMark = "SUPP" Then
        Call addSupplierName(cmbSupplier)
    ElseIf strMark = "PF" Then
        Call addPFNo(cmbMainPF)
    ElseIf strMark = "ITEM" Then
        Call addItemType(cmbItemType)
        Call addItemName(cmbItem, cmbItemType.Text)
    End If
End Sub

Private Sub Form_Load()
    Call prcMakeCenter(Me)
    Call addSupplierName(cmbSupplier)
    Call addPFNo(cmbMainPF)
    Call addCurrency(cmbCurrency)
    cmbCurrency.Text = findDCurrency
    'Call addUnit(cmbUnit)
    'Call addItemType(cmbItemType)
    CID = bytCompID
    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    strRecord = "Select DelStatus,UpStatus from InventoryPJL.dbo.Sys_User_Name where UserName='" & strUser & "'"
    R.Open strRecord, MainConn, adOpenStatic
    If R!DelStatus = 1 Then
       cmdDelete.Enabled = True
    End If
    If R!UpStatus = 1 Then
        cmdEdit.Enabled = True
    End If
    
    txtMode.Text = "NEW"
    
    With flexDetails
        .Rows = 1
        .Cols = 22
        .Row = 0
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterLeft
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignCenterCenter
        .ColAlignment(15) = flexAlignCenterCenter
        .ColAlignment(16) = flexAlignCenterCenter
        .ColAlignment(17) = flexAlignCenterCenter
        .ColAlignment(18) = flexAlignCenterLeft
        .ColAlignment(19) = flexAlignCenterLeft
        .ColAlignment(20) = flexAlignCenterLeft
        .ColAlignment(21) = flexAlignCenterCenter
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 0
        .ColWidth(4) = 2000
        .ColWidth(5) = 1600
        .ColWidth(6) = 2000
        .ColWidth(7) = 2000
        .ColWidth(8) = 3000
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1200
        .ColWidth(13) = 1500
        .ColWidth(14) = 900
        .ColWidth(15) = 1500
        .ColWidth(16) = 1500
        .ColWidth(17) = 3500
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(20) = 1500
        
        .Col = 0:      .Text = "CHALLAN NO"       'Challan No
        .Col = 1:      .Text = "CHALLAN DATE"     'Challan Date
        .Col = 2:      .Text = "ACTUAL PF"        'Actual PF
        .Col = 3:      .Text = "PI ID"            'Buyer ID
        .Col = 4:      .Text = "P.O NO."          'PI No.
        .Col = 5:      .Text = "CONTAINER NO"     'Container No
        .Col = 6:      .Text = "ITEM GROUP"       'Item Group
        .Col = 7:      .Text = "ITEM"             'Item Name
        .Col = 8:      .Text = "DESCRIPTION"      'Description
        .Col = 9:      .Text = "MATERIALS CODE"   'Materials Code
        .Col = 10:     .Text = "PR No"            'po no
        .Col = 11:     .Text = "CHALLAN QTY"      'Challan Qty
        .Col = 12:     .Text = "RECEIVED QTY"     'Received Qty
        .Col = 13:     .Text = "SHORT/EXCESS QTY" 'Short/Excess
        .Col = 14:     .Text = "QTY UNIT"         'Qty Unit
        .Col = 15:     .Text = "UNIT PRICE"       'Unit Price
        .Col = 16:     .Text = "TOTAL"            'Total Amount
        .Col = 17:     .Text = "REMARKS"          'Remarks
        .Col = 18:     .Text = "SL NO"            'SL No
        .Col = 19:     .Text = "PC ID"            'PC ID
        .Col = 20:     .Text = "USER ID"          'USER ID
        .Col = 21:     .Text = "DESCRIPTION 2"    'Description 2
    End With
    
    With flexItem
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        
        .ColWidth(0) = 0
        .ColWidth(1) = 1600
        .ColWidth(2) = 2000
        .ColWidth(3) = 1900
        .ColWidth(4) = 1000
        .ColWidth(5) = 800
        .ColWidth(6) = 700
        .ColWidth(7) = 800
        .ColWidth(8) = 1000
        .ColWidth(9) = 500
        
        .Row = 0
        .Col = 0:    .Text = "Item ID"
        .Col = 1:    .Text = "Item Group"
        .Col = 2:    .Text = "Item Name"
        .Col = 3:    .Text = "Description"
        .Col = 4:    .Text = "PR No"
        .Col = 5:    .Text = "PO Qty"
        .Col = 6:    .Text = "Unit"
        .Col = 7:    .Text = "Unit Price"
        .Col = 8:    .Text = "Total"
        .Col = 9:    .Text = ""
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strMark = ""
    Set frmMrrReceiveChallan = Nothing
End Sub

Private Sub txtAmount_GotFocus()
    Call txtGotFocus(txtAmount)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtChallanDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtChallanDate_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtChallanDate_LostFocus()
    Call FormatDate(txtChallanDate)
End Sub

Private Sub txtChallanNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtChallanNo_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtCharges_GotFocus()
    Call txtGotFocus(txtCharges)
End Sub

Private Sub txtCharges_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtCharges_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtCharges_LostFocus()
    If Len(txtCharges) = 0 Then txtCharges.Text = "0.00"
    txtCharges.Text = Format(txtCharges.Text, "##0.00")
    txtMRRValue.Text = Format(TotalMRRValue, "##,##0.00")
End Sub

Private Sub txtmatcode_GotFocus()
    Call txtGotFocus(txtmatcode)
End Sub

Private Sub txtmatcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtmatcode_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub txtmatcode_LostFocus()
    txtmatcode.Text = Trim(UCase(txtmatcode.Text))
End Sub

Private Sub txtcont_GotFocus()
    Call txtGotFocus(txtCont)
End Sub

Private Sub txtcont_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtcont_LostFocus()
    txtCont.Text = UCase(Trim(txtCont.Text))
End Sub

Private Sub txtDesc_GotFocus()
'    SendKeys "^{Home}+^{End}"
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Text = Trim(UCase(txtDesc.Text))
End Sub

Private Sub txtDiscount_GotFocus()
    Call txtGotFocus(txtDiscount)
End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtDiscount_LostFocus()
    If Len(txtDiscount) = 0 Then txtDiscount.Text = "0.00"
    txtDiscount.Text = Format(txtDiscount.Text, "##0.00")
    txtMRRValue.Text = Format(TotalMRRValue, "##,##0.00")
End Sub

Private Sub txtMRRDate_GotFocus()
    Call txtGotFocus(txtMrrDate)
End Sub

Private Sub txtMRRDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtMRRDate_KeyPress(KeyAscii As Integer)
    Call DateFunc(KeyAscii)
End Sub

Private Sub txtMRRDate_LostFocus()
    txtMrrDate.Text = FormatDate(txtMrrDate)
End Sub

Private Sub txtMRRID_GotFocus()
    Call txtGotFocus(txtMRRID)
End Sub

Private Sub txtMRRID_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtMRRID_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub


Private Sub txtMRRQty_Click()
txtUnitPrice.Text = txtMRRQty.ItemData(txtMRRQty.ListIndex)
    Set R = New ADODB.Recordset
        strRecord = "SELECT UNITPRICE,CURRENCY FROM dbo.Pro_Inv_SubOC where piID=" & findPIID(cmbPINO.Text) & " AND PRNO='" & cmbprno.Text & "' AND ItemID='" & findItemID(cmbItem) & "' AND ItDesc= '" & cmbDescription.Text & "' And DeleteRow=0 AND PIQTY=" & txtMRRQty.Text & " ;"
        R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        txtUnitPrice.Text = R![UnitPrice]
    End If
    Set R = Nothing
End Sub

Private Sub txtMRRQty_GotFocus()
'    Call txtGotFocus(txtMRRQty)
End Sub

Private Sub txtMRRQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtMRRQty_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtMRRQty_LostFocus()
'    If Len(txtUnitPrice) <> 0 And Len(txtMRRQty) <> 0 Then
'        txtAmount.Text = Format(txtMRRQty.Text, "##0.00")
'    Else
'        txtAmount.Text = "0.00"
'    End If
 If Len(txtUnitPrice.Text) <> 0 And Len(txtReceivedQty.Text) <> 0 Then
        txtAmount.Text = Format(txtReceivedQty.Text * txtUnitPrice.Text, "##0.00")
    Else
        txtAmount.Text = "0.00"
    End If
    Call prcShortExcess
End Sub

Private Sub txtMRRValue_GotFocus()
    Call txtGotFocus(txtMRRValue)
End Sub

Private Sub txtReceivedQty_GotFocus()
    Call txtGotFocus(txtReceivedQty)
End Sub

Private Sub txtReceivedQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtReceivedQty_KeyPress(KeyAscii As Integer)
    Call CurrFuncForMrrRecv(KeyAscii)
End Sub

Private Sub txtReceivedQty_LostFocus()
If Val(txtReceivedQty) <= 0 Then
MsgBox "Enter quantity properly", vbInformation, "Wrong quantity"
Exit Sub
End If

 On Error Resume Next
    If txtMode.Text = "NEW" Then
        TempMRRVal = Format(findTotal, "##,##0.00") + CCur(txtReceivedQty.Text)
    ElseIf txtMode.Text = "EDIT" And cmdAdd.Caption = "&Replace" Then
        TempMRRVal = Format(findTotal, "##,##0.00") - tempRecQtyEdit
        TempMRRVal = TempMRRVal + CCur(txtReceivedQty.Text)
    ElseIf txtMode.Text = "EDIT" And cmdAdd.Caption = "&Add" Then
        TempMRRVal = Format(findTotal, "##,##0.00") + CCur(txtReceivedQty.Text)
    End If

    If TempMRRVal <= CCur(txtMRRQty.Text) Then
        Else
        txtReceivedQty.Text = ""
        txtReceivedQty.SetFocus
        Exit Sub
    End If
    If Len(txtUnitPrice.Text) <> 0 And Len(txtReceivedQty.Text) <> 0 Then
        txtAmount.Text = Format(txtReceivedQty.Text * txtUnitPrice.Text, "##0.00")
    Else
        txtAmount.Text = "0.00"
    End If
    Call prcShortExcess

End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtShortExcess_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub


Private Sub cmbprno_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmbprno_KeyPress(KeyAscii As Integer)
    Call SingleCodeFunc(KeyAscii)
End Sub

Private Sub cmbprno_LostFocus()
    cmbprno.Text = Trim(UCase(cmbprno.Text))
End Sub

Private Sub txtUnitPrice_GotFocus()
    Call txtGotFocus(txtUnitPrice)
End Sub

Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    Call CurrFunc(KeyAscii)
End Sub

Private Sub txtUnitPrice_LostFocus()
'    If Len(txtUnitPrice) <> 0 And Len(txtMRRQty) <> 0 Then
'        txtUnitPrice.Text = Format(txtUnitPrice.Text, "##,##0.00")
'        txtAmount.Text = Format(txtReceivedQty.Text * txtUnitPrice.Text, "##,##0.00")
'    Else
'        txtAmount.Text = "0.00"
'    End If


If Len(txtUnitPrice.Text) <> 0 And Len(txtReceivedQty.Text) <> 0 Then
'        txtUnitPrice.Text = Format(txtUnitPrice.Text, "##,##0.00")
        txtAmount.Text = Format(txtReceivedQty.Text * txtUnitPrice.Text, "##0.00")
    Else
        txtAmount.Text = "0.00"
    End If
End Sub

Private Sub chkTextNull()
    boEmptyText = True
    If Len(cmbSupplier) = 0 Then
        MsgBox "Select Supplier Name", vbInformation, cnstMsgInfo
        cmbSupplier.SetFocus
        Exit Sub
    End If
    If txtMode.Text = "NEW" Then
'        If Len(txtMRRID) = 0 Then
'            MsgBox "Type MRR No", vbInformation, cnstMsgInfo
'            txtMRRID.SetFocus
'        Exit Sub
'        End If
    Else
        If Len(cmbMRRID) = 0 Then
            MsgBox "Select MRR ID", vbInformation, cnstMsgInfo
            cmbMRRID.SetFocus
            Exit Sub
        End If
    End If
    If Len(txtMrrDate) = 0 Then
        MsgBox "Type MRR Date", vbInformation, cnstMsgInfo
        txtMrrDate.SetFocus
        Exit Sub
    End If

    If Len(cmbCurrency) = 0 Then
        MsgBox "Select MRR Currency", vbInformation, cnstMsgInfo
        cmbCurrency.SetFocus
        Exit Sub
    End If
    If Len(Name1) = 0 Then
        MsgBox "Select Quality Certified Name", vbInformation, cnstMsgInfo
        Empcode1.SetFocus
        Exit Sub
    End If
    
    If Len(Name2) = 0 Then
        MsgBox "Select Prepared Name", vbInformation, cnstMsgInfo
        Empcode2.SetFocus
        Exit Sub
    End If
    
    If Len(Name3) = 0 Then
        MsgBox "Select Quantity Check Name", vbInformation, cnstMsgInfo
        Empcode3.SetFocus
        Exit Sub
    End If
    
    If Len(Name4) = 0 Then
        MsgBox "Select Manager Name", vbInformation, cnstMsgInfo
        Empcode4.SetFocus
        Exit Sub
    End If
    
    If Len(Name5) = 0 Then
        MsgBox "Select Quality Senior Manager Name", vbInformation, cnstMsgInfo
        Empcode5.SetFocus
        Exit Sub
    End If
    
    If flexDetails.Rows = 1 Then
        MsgBox "Record not found to Save/Update", vbInformation, cnstMsgInfo
        boEmptyText = True
    End If
    boEmptyText = False
End Sub

Function TotalMRRValue() As Currency
    Dim curTotalValue As Currency
    curTotalValue = 0
    With flexDetails
        If .Rows > 1 Then
            For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 16:  curTotalValue = curTotalValue + CCur(.Text)
            Next inti
        End If
    End With
    TotalMRRValue = curTotalValue + CCur(txtCharges.Text) - CCur(txtDiscount.Text)
End Function

Private Sub prcSaveGriddate(strMRRID As String)
Dim strChallNo As String, strChallDate As String, strPFNo As String, intPIID As Long, strCont As String
Dim intItemID As Long, strItDesc As String, strItDesc2 As String, strMatCode As String, strPRNo As String
Dim dblMRRQty As Double, dblRcvdQty As Double, dblSE As Double, strUnit As String
Dim curUnitPrice As Double, curTotalAmount As Currency, strRemarks As String, strSLNO As Integer, INTPCID As Integer, INTUSERID As Integer
    
If txtMode.Text = "NEW" Then
    With flexDetails
        For inti = 1 To .Rows - 1
            .Row = inti

            .Col = 0:   strChallNo = .Text
            .Col = 1:   strChallDate = Format(.Text, cnstDtFrmtQ)
            .Col = 2:   strPFNo = .Text
            .Col = 3:   intPIID = .Text
            .Col = 5:   strCont = .Text
            .Col = 7:   intItemID = findItemID(.Text)
            .Col = 8:   strItDesc = .Text
            .Col = 9:   strMatCode = .Text
            .Col = 10:  strPRNo = .Text
            .Col = 11:  dblMRRQty = CDbl(.Text)
            .Col = 12:  dblRcvdQty = CDbl(.Text)
            .Col = 14:  strUnit = .Text
            .Col = 15:  curUnitPrice = CDbl(.Text)
            .Col = 16:  curTotalAmount = CCur(.Text)
            .Col = 17:  strRemarks = .Text
            .Col = 18:  strSLNO = .Text
            .Col = 19:  INTPCID = .Text
            .Col = 20:  INTUSERID = .Text
            .Col = 21:  strItDesc2 = .Text
            
            MainComm.CommandText = "INSERT INTO MRR_Challan_Sub(MRRID, ItemID, ItDesc, MatCode, PRNO, MRRQty, RcvdQty, QtyUnit, UnitPrice, ItTotal, Remarks, ChallanNo, ChallanDate, MainPF, PIID, Containerno,slno,PCID,USERID,EntryDate,EntryTime,Remarks2) " & _
                " VALUES ('" & strMRRID & "', " & intItemID & ", '" & strItDesc & "', '" & strMatCode & "', '" & strPRNo & "', " & dblMRRQty & ", " & dblRcvdQty & ", '" & strUnit & "', " & curUnitPrice & ", " & curTotalAmount & ", '" & strRemarks & "', '" & strChallNo & "', '" & strChallDate & "', '" & strPFNo & "', " & intPIID & ", '" & strCont & "'," & strSLNO & "," & INTPCID & "," & INTUSERID & ",'" & Date & "','" & Time & "', '" & strItDesc2 & "');"
            MainComm.Execute
        Next inti
    End With
Else
    If cmdAdd.Caption = "&Add" Then
        With flexDetails
            .Row = bytRowNo
            
            .Col = 0:    strChallNo = .Text
            .Col = 1:    strChallDate = Format(.Text, cnstDtFrmtQ)
            .Col = 2:    strPFNo = .Text
            .Col = 3:    intPIID = .Text
            .Col = 5:    strCont = .Text
            .Col = 7:    intItemID = findItemID(.Text)
            .Col = 8:    strItDesc = .Text
            .Col = 9:    strMatCode = .Text
            .Col = 10:   strPRNo = .Text
            .Col = 11:   dblMRRQty = CDbl(.Text)
            .Col = 12:   dblRcvdQty = CDbl(.Text)
            .Col = 14:   strUnit = .Text
            .Col = 15:   curUnitPrice = CDbl(.Text)
            .Col = 16:   curTotalAmount = CCur(.Text)
            .Col = 17:   strRemarks = .Text
            .Col = 18:   strSLNO = addslno
            .Col = 19:   INTPCID = .Text
            .Col = 20:   INTUSERID = .Text
            .Col = 21:   strItDesc2 = .Text
            
            Call prCMAX(vmrrid)
            MainComm.CommandText = "INSERT INTO MRR_Challan_Sub(MRRID, ItemID, ItDesc, MatCode, PRNO, MRRQty, RcvdQty, QtyUnit, UnitPrice, ItTotal, Remarks, ChallanNo, ChallanDate, MainPF, PIID, Containerno,slno,PCID,USERID,Remarks2,Entrydate,Entrytime) " & _
                " VALUES ('" & strMRRID & "', " & intItemID & ", '" & strItDesc & "', '" & strMatCode & "', '" & strPRNo & "', " & dblMRRQty & ", " & dblRcvdQty & ", '" & strUnit & "', " & curUnitPrice & ", " & curTotalAmount & ", '" & strRemarks & "', '" & strChallNo & "', '" & strChallDate & "', '" & strPFNo & "', " & intPIID & ", '" & strCont & "'," & strSLNO & "," & INTPCID & "," & INTUSERID & ", '" & strItDesc2 & "','" & Date & "','" & Time & "');"
            MainComm.Execute
            
            MainComm.CommandText = "UPDATE Mrr_Challan_Main SET suppID=" & findSupplierID(cmbSupplier.Text) & ", MRRID='" & strMRRID & "', MRRDate='" & Format(txtMrrDate.Text, cnstDtFrmtQ) & "'," _
                  & " MRRCurr='" & cmbCurrency.Text & "', MRRValue=" & CCur(txtMRRValue.Text) & ", MRRCharge=" & CCur(txtCharges.Text) & ", MRRDiscount=" & CCur(txtDiscount.Text) & ",PCID=" & getComVolID & ",USERID=" & bytUserID & ",Entrydate='" & Date & "',Entrytime='" & Time & "'" _
                  & " WHERE Compid = " & bytCompID & " and  MRRID='" & txtEditMRRid.Text & "';"
            MainComm.Execute
        End With
    Else
        With flexDetails
            .Row = bytRowNo
            
            .Col = 0:   strChallNo = .Text
            .Col = 1:   strChallDate = Format(.Text, cnstDtFrmtQ)
            .Col = 2:   strPFNo = .Text
            .Col = 3:   intPIID = .Text
            .Col = 5:   strCont = .Text
            .Col = 7:   intItemID = findItemID(.Text)
            .Col = 8:   strItDesc = .Text
            .Col = 9:   strMatCode = .Text
            .Col = 10:  strPRNo = .Text
            .Col = 11:  dblMRRQty = CDbl(.Text)
            .Col = 12:  dblRcvdQty = CDbl(.Text)
            .Col = 14:  strUnit = .Text
            .Col = 15:  curUnitPrice = CDbl(.Text)
            .Col = 16:  curTotalAmount = CCur(.Text)
            .Col = 17:  strRemarks = .Text
            .Col = 18:  strSLNO = .Text
            .Col = 19:  INTPCID = .Text
            .Col = 20:  INTUSERID = .Text
            .Col = 21:  strItDesc2 = .Text
            
            'MsgBox strSLNO
            MainComm.CommandText = "INSERT INTO MRR_Challan_Sub(MRRID, ItemID, ItDesc, MatCode, PRNO, MRRQty, RcvdQty, QtyUnit, UnitPrice, ItTotal, Remarks, ChallanNo, ChallanDate, MainPF, PIID, Containerno,slno,PCID,USERID,Remarks2,Entrydate,Entrytime) " & _
                " VALUES ('" & strMRRID & "', " & intItemID & ", '" & strItDesc & "', '" & strMatCode & "', '" & strPRNo & "', " & dblMRRQty & ", " & dblRcvdQty & ", '" & strUnit & "', " & curUnitPrice & ", " & curTotalAmount & ", '" & strRemarks & "', '" & strChallNo & "', '" & strChallDate & "', '" & strPFNo & "', " & intPIID & ", '" & strCont & "'," & strSLNO & "," & INTPCID & "," & INTUSERID & ", '" & strItDesc2 & "','" & Date & "','" & Time & "');"
            MainComm.Execute
        
            MainComm.CommandText = "UPDATE Mrr_Challan_Main SET suppID=" & findSupplierID(cmbSupplier.Text) & ", MRRID='" & strMRRID & "', MRRDate='" & Format(txtMrrDate.Text, cnstDtFrmtQ) & "'," _
                  & " MRRCurr='" & cmbCurrency.Text & "', MRRValue=" & CCur(txtMRRValue.Text) & ", MRRCharge=" & CCur(txtCharges.Text) & ", MRRDiscount=" & CCur(txtDiscount.Text) & ",PCID=" & getComVolID & ",USERID=" & bytUserID & ",Entrydate='" & Date & "',Entrytime='" & Time & "'" _
                  & " WHERE Compid = " & bytCompID & " and  MRRID='" & txtEditMRRid.Text & "';"
            MainComm.Execute
            
            cmdAdd.Caption = "&Add"
        End With
    End If
End If
End Sub

Private Sub prcShortExcess()
    If Len(txtMRRQty.Text) = 0 Then txtShortExcess.Text = "0.00": Exit Sub
    If Len(txtReceivedQty.Text) = 0 Then txtShortExcess.Text = "0.00": Exit Sub
    
    txtShortExcess.Text = FormatNumber(CDbl(txtMRRQty.Text) - CDbl(txtReceivedQty.Text), 2)
End Sub

Function findTotal() As Currency
 On Error Resume Next
    findTotal = 0
    Set RS = New ADODB.Recordset
    strRecord = "Select Sum(RcvdQty) As TValue From  MRR_Challan_Sub Where PIId=" & cmbPINO.Tag & " And PRNO='" & cmbprno.Text & "' And MRRQty='" & txtMRRQty.Text & "' And DeleteRow=0 AND ItemID='" & findItemID(cmbItem) & "' AND ItDesc= '" & cmbDescription.Text & "';"
    RS.Open strRecord, MainConn, adOpenStatic
    If RS.RecordCount > 0 Then
        If Not IsNull(RS![TValue]) Then findTotal = RS![TValue]
    End If
    RS.Close
    Set RS = Nothing
'ErrHand:
'    MsgBox Err.Description, vbCritical
End Function

Private Sub DisableSignature()
    Empcode1.Enabled = False
    Empcode2.Enabled = False
    Empcode3.Enabled = False
    Empcode4.Enabled = False
    Empcode5.Enabled = False
End Sub
Private Sub EnableSignature()
    Empcode1.Enabled = True
    Empcode2.Enabled = True
    Empcode3.Enabled = True
    Empcode4.Enabled = True
    Empcode5.Enabled = True
End Sub


