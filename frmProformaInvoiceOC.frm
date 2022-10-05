VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProformaInvoiceOC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9690
   ClientLeft      =   2580
   ClientTop       =   1575
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmProformaInvoiceOC.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSubMenu 
      BackColor       =   &H00C0C000&
      Caption         =   "&Deduction"
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
      Left            =   2805
      TabIndex        =   51
      ToolTipText     =   "Preview PI"
      Top             =   9000
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
      Left            =   4110
      TabIndex        =   22
      ToolTipText     =   "Preview PI"
      Top             =   9000
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
      Left            =   10590
      TabIndex        =   21
      ToolTipText     =   "Close Form"
      Top             =   8985
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
      Left            =   6690
      TabIndex        =   19
      ToolTipText     =   "Refresh"
      Top             =   9000
      UseMaskColor    =   -1  'True
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
      Left            =   9300
      TabIndex        =   16
      ToolTipText     =   "Save/Update Record"
      Top             =   8985
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
      Left            =   5400
      TabIndex        =   20
      ToolTipText     =   "Delete Record"
      Top             =   9000
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
      Left            =   7995
      TabIndex        =   17
      ToolTipText     =   "New/Edit Record"
      Top             =   9000
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
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   43
      ToolTipText     =   "Mode of Data Entry"
      Top             =   9015
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtEditPINo 
      Alignment       =   2  'Center
      BackColor       =   &H00E26101&
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
      Left            =   1350
      MaxLength       =   30
      TabIndex        =   23
      Top             =   9060
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P.O.  Details"
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
      Height          =   5895
      Left            =   180
      TabIndex        =   24
      Top             =   2760
      Width           =   11595
      Begin VB.TextBox txtCondition 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   57
         Top             =   1560
         Width           =   9495
      End
      Begin VB.TextBox txtCondition0 
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
         Height          =   315
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1560
         Visible         =   0   'False
         Width           =   555
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
         ItemData        =   "frmProformaInvoiceOC.frx":0442
         Left            =   6375
         List            =   "frmProformaInvoiceOC.frx":0444
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1170
         Width           =   1305
      End
      Begin VB.ComboBox cboItType 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   510
         Width           =   2340
      End
      Begin VB.ComboBox cboItem 
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
         Height          =   286
         Left            =   2580
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   510
         Width           =   4121
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
         Left            =   6750
         Picture         =   "frmProformaInvoiceOC.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Add New Item"
         Top             =   510
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtPrNo 
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
         Left            =   180
         MaxLength       =   200
         TabIndex        =   9
         Top             =   1170
         Width           =   2340
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
         Height          =   315
         Left            =   10845
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Add to List"
         Top             =   1170
         Width           =   600
      End
      Begin VB.TextBox txtPIQty 
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
         Height          =   315
         Left            =   2580
         TabIndex        =   10
         Top             =   1170
         Width           =   1500
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
         Height          =   315
         Left            =   7740
         TabIndex        =   13
         Top             =   1170
         Width           =   1455
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
         Height          =   315
         Left            =   9345
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1170
         Width           =   1395
      End
      Begin VB.ComboBox cboUnit 
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
         Left            =   4140
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1170
         Width           =   2190
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
         Left            =   7125
         MaxLength       =   255
         TabIndex        =   8
         ToolTipText     =   "Item Description"
         Top             =   510
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   3810
         Left            =   105
         TabIndex        =   25
         Top             =   2040
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   6720
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   0
         RowHeightMin    =   285
         BackColor       =   16777215
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Conditions : "
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
         Left            =   240
         TabIndex        =   58
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unit : "
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
         Left            =   4140
         TabIndex        =   47
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Currency:"
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
         Left            =   6405
         TabIndex        =   46
         Top             =   915
         Width           =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   11640
         Y1              =   1920
         Y2              =   1920
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
         TabIndex        =   39
         Top             =   270
         Width           =   1125
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
         Left            =   9345
         TabIndex        =   31
         Top             =   900
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
         Left            =   7740
         TabIndex        =   30
         Top             =   900
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
         Left            =   2580
         TabIndex        =   29
         Top             =   900
         Width           =   705
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
         Left            =   7125
         TabIndex        =   28
         Top             =   240
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
         TabIndex        =   27
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PR No : "
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
         TabIndex        =   26
         Top             =   900
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   180
      TabIndex        =   32
      Top             =   960
      Width           =   11595
      Begin VB.CheckBox ChkShow 
         Caption         =   "Show"
         Height          =   255
         Left            =   6480
         TabIndex        =   56
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtTDate 
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
         Left            =   7878
         TabIndex        =   5
         Top             =   1200
         Width           =   1830
      End
      Begin VB.TextBox txtFDate 
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
         Left            =   5040
         TabIndex        =   4
         Top             =   1200
         Width           =   1830
      End
      Begin VB.CheckBox chkpad 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pad Report"
         Height          =   255
         Left            =   8040
         TabIndex        =   50
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboPINo 
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
         Height          =   315
         Left            =   1530
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   780
         Visible         =   0   'False
         Width           =   4035
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
         Left            =   6960
         Picture         =   "frmProformaInvoiceOC.frx":0538
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Add New Supplier"
         Top             =   240
         Visible         =   0   'False
         Width           =   285
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1530
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
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
         Left            =   5610
         TabIndex        =   3
         ToolTipText     =   "Show Record"
         Top             =   780
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPINo 
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
         TabIndex        =   18
         Top             =   780
         Width           =   4035
      End
      Begin VB.TextBox txtPIDate 
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
         TabIndex        =   2
         Top             =   1215
         Width           =   1830
      End
      Begin VB.Label Label13 
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
         Left            =   7680
         TabIndex        =   55
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label Label12 
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
         Left            =   4920
         TabIndex        =   54
         Top             =   1200
         Width           =   60
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
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
         Index           =   5
         Left            =   6960
         TabIndex        =   53
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   3840
         TabIndex        =   52
         Top             =   1200
         Width           =   960
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
         ForeColor       =   &H00E26101&
         Height          =   255
         Left            =   1380
         TabIndex        =   49
         Top             =   1200
         Width           =   180
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
         Height          =   255
         Left            =   1380
         TabIndex        =   48
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supplier Name"
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
         TabIndex        =   37
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P.O.  No."
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
         TabIndex        =   36
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P.O.  Date"
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
         Index           =   0
         Left            =   180
         TabIndex        =   35
         Top             =   1230
         Width           =   1200
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
         Left            =   1380
         TabIndex        =   34
         Top             =   810
         Width           =   180
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   15
      TabIndex        =   44
      Top             =   8760
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   15
      TabIndex        =   42
      Top             =   8790
      Width           =   12240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Currency Purchase Order"
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
      TabIndex        =   41
      Top             =   210
      Width           =   5040
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
      TabIndex        =   40
      Top             =   855
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmProformaInvoiceOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim boEdit As Boolean, bytRowNo As Byte
Dim intYear As Integer
Dim intPI As Integer
Dim strPI As String
Dim strGroupNameShort As String
Dim intGSLNO As Integer
Dim intGMaxSLNO As Long
Dim ITEM1 As String
Dim DES1 As String
Private Sub cboItem_Click()
  cboItem.Tag = ""
  If Len(cboItem) = 0 Then Exit Sub
  cboItem.Tag = cboItem.ItemData(cboItem.ListIndex)
  cboItem.ToolTipText = cboItem.Tag
  ITEM1 = cboItem.Text
End Sub
Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub

Private Sub cboItType_Click()
    If Len(cboItType) = 0 Then Exit Sub
    addItemName cboItem, cboItType.Text
End Sub

Public Function fncMaxSL(ByVal Param As Long)
    On Error GoTo x
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset
        Mysql = "Select  Max(SLNO) AS NewID From Pro_Inv_SubOC Where PIID = " & Param & ";"
    
        MainComm.CommandText = Mysql
        Set R = MainComm.Execute
        
        If IsNull(R(0).Value) Then
            fncMaxSL = 1
        Else
            fncMaxSL = R(0).Value + 1
        End If
    Set R = Nothing
    Exit Function
x:
        MsgBox err.Description, vbCritical, AppTitle
    Set R = Nothing
End Function

Private Sub cboItType_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboPINo_Click()
    cboPINo.Tag = cboPINo.ItemData(cboPINo.ListIndex)
    intGMaxSLNO = fncMaxSL(cboPINo.Tag)
End Sub

Private Sub cboPINo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cboPINo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub

Private Sub cboPINo_LostFocus()
    cboPINo.Text = Trim(UCase(cboPINo.Text))
End Sub

Private Sub cboSupplier_Click()
    If Len(cboSupplier) = 0 Then Exit Sub
    addSuppPINoOC cboPINo, cboSupplier.Text
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub



Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub ChkShow_Click()
If ChkShow.Value = 1 Then
cboPINo.Visible = True
txtPINo.Visible = False
If Len(cboSupplier) = 0 Then Exit Sub
    addSuppPINoOC cboPINo, cboSupplier.Text

End If

If ChkShow.Value = 0 Then
cboPINo.Visible = False
txtPINo.Visible = True
cboPINo.Text = ""
End If
End Sub

Private Sub cmbCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub cmdAdd_Click()

Dim ITEM As String
Dim DES As String
Dim PR1 As String
Dim PR2 As String

    If Len(cboItem) = 0 Then
        MsgBox "Select Name of Item", vbInformation, cnstMsgInfo
        cboItem.SetFocus
        Exit Sub
    End If
    
    If Len(txtPrNo) = 0 Then
        MsgBox "Type PRNO", vbInformation, cnstMsgInfo
        txtPrNo.SetFocus
        Exit Sub
    End If
    
    If Len(txtPIQty) = 0 Then
        MsgBox "Type P.O. Qty", vbInformation, cnstMsgInfo
        txtPIQty.SetFocus
        Exit Sub
    End If
    
    If Len(txtUnitPrice) = 0 Then
        MsgBox "Type Unit Price", vbInformation, cnstMsgInfo
        txtUnitPrice.SetFocus
        Exit Sub
    End If
    
    If CCur(txtAmount) = 0 Then
        MsgBox "Amount must be grater than '0'", vbInformation, cnstMsgInfo
        txtUnitPrice.SetFocus
        Exit Sub
    End If
    If cmbCurrency.Text = "" Then
        MsgBox "Select Currency, vbInformation, cnstMsgInfo"
        cmbCurrency.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    '=======Edit Row ====='
'    Set MainComm = New ADODB.Command
'    MainComm.ActiveConnection = MainConn
'    MainComm.CommandText = "UPDATE Pro_Inv_SubOC SET DeleteRow = 2,EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' WHERE PIID='" & cboPINo.Tag & "' AND SLNO = " & intGSLNO & ";"
'    MainComm.Execute
    intGSLNO = 0
    '====================='
       If cmdAdd.Caption = "&Add" Then
    
    With flexDetails
                
                For inti = 1 To .Rows - 1
                .Row = inti
                .Col = 0:  ITEM = .Text
                .Col = 1:  DES = .Text
                .Col = 1:  PR1 = .Text
                
                DES1 = txtDesc.Text
                PR2 = txtPrNo.Text
                
                If ITEM1 = ITEM And DES1 = DES And PR1 = PR2 Then
                MsgBox "Sorry, already add list"
                Exit Sub
                End If
                Next inti
        
    End With
    
    End If
    
    With flexDetails
        If boEdit = False Then
          .Rows = .Rows + 1
          .Row = .Rows - 1
        Else
          .Row = bytRowNo
        End If
        .Col = 0:  .Text = cboItem.Text
        .Col = 1:  .Text = txtDesc.Text
        .Col = 2:  .Text = txtPrNo.Text
        .Col = 3:  .Text = txtPIQty.Text
        .Col = 4:  .Text = cboUnit.Text
        .Col = 5:  .Text = txtUnitPrice.Text
        .Col = 6:  .Text = txtAmount.Text
        If cmdEdit.Caption = "&Edit" Then
            .Col = 7:  .Text = .Row
        ElseIf cmdEdit.Caption = "&New" Then
            .Col = 7:  .Text = intGMaxSLNO
            intGMaxSLNO = intGMaxSLNO + 1
        End If
        .Col = 8:  .Text = cmbCurrency.Text
        '.Col = 9:  .Text = txtmatcode.Text
        .Col = 9:  .Text = cboItem.Tag
    End With
    boEdit = False
    'txtPIValue.Text = Format(TotalPIValue, "##,##0.00")
    cboItem.ListIndex = -1
    txtDesc.Text = ""
    txtPrNo.Text = ""
    txtPIQty.Text = "0"
    cboUnit.ListIndex = -1
    txtUnitPrice.Text = "0.000"
    txtAmount.Text = "0.00"
    'txtmatcode.Text = ""
    cboItType.Enabled = False
    cboItem.SetFocus
ErrorHandler:
    If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB
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
    chkTextNull
    If boEmptyText = True Then Exit Sub
    
    If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
    
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans
    
    MainComm.CommandText = "UPDATE New_PI SET DeleteRow =1,UserId=" & bytUserID & ",EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' WHERE PIID='" & cboPINo.Tag & "' And ComId=" & bytCompID & " AND DeleteRow = 0;"
    MainComm.Execute
    MainComm.CommandText = "UPDATE Pro_Inv_MainOC SET DeleteRow = 1,UserId=" & bytUserID & ",EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' WHERE PIID='" & cboPINo.Tag & "'And CompId=" & bytCompID & " AND DeleteRow = 0;"
    MainComm.Execute
    MainComm.CommandText = "UPDATE Pro_Inv_SubOC SET DeleteRow = 1,UserId=" & bytUserID & ",EntryDate = '" & Date & "' ,EntryTime = '" & Time & "' WHERE PIID='" & cboPINo.Tag & "' AND DeleteRow = 0;"
    MainComm.Execute
    
    MainConn.CommitTrans
ErrorHandler:
    If err.Number <> 0 Then
        If err.Number = -2147467259 Then
            MsgBox "This record cannot be deleted", vbCritical, cnstMsgErDB
        Else
            MsgBox err.Description, vbCritical, cnstMsgErDB
        End If
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        Exit Sub
    End If
    cmdRefresh_Click
    cboPINo.Text = ""
    txtPIDate.Text = ""
'    addSuppPINo cboPINo, cboSupplier.Text  'R1
     addSuppPINoOC cboPINo, cboSupplier.Text
    cboPINo.SetFocus
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdEdit_Click()
    If txtMode.Text = "NEW" Then
        cmdSave.Caption = "&Update"
        cmdDelete.Enabled = True
        txtPINo.Visible = False
        cboPINo.Top = txtPINo.Top
        cboPINo.Visible = True
        cmdShow.Top = cboPINo.Top
        cmdShow.Visible = True
        cmdEdit.Caption = "&New"
        cboItType.Enabled = False
        txtMode.Text = "EDIT"
'        addSuppPINo cboPINo, cboSupplier.Text 'R1
        addSuppPINoOC cboPINo, cboSupplier.Text
    Else
        cmdSave.Caption = "&Save"
        cmdDelete.Enabled = False
        cboPINo.Visible = False
        cmdShow.Visible = False
        txtPINo.Visible = True
        cboItType.Enabled = True
        cmdEdit.Caption = "&Edit"
        txtMode.Text = "NEW"
    End If
    If txtMode.Text = "NEW" Then txtPINo.Text = "" Else cboPINo.Text = ""
    
    txtPIDate.Text = ""
    
    cboItem.ListIndex = -1
    txtDesc.Text = ""
    txtPrNo.Text = ""
    
    txtPIQty.Text = ""
    cboUnit.ListIndex = -1
    txtUnitPrice.Text = "0.000"
    
    txtAmount.Text = "0.00"
    flexDetails.Rows = 1
    
    cboSupplier.SetFocus
End Sub
Private Sub cmdNewMainPF_Click()
    strMark = "PF"
    frmNewPFNo.Show 1
End Sub
Private Sub cmdNewSupplier_Click()
    strMark = "SUPP"
    frmNewSupplier.Show 1
End Sub
Private Sub cmdNewUsesPF_Click()
    strMark = "USESPF"
    frmNewUsesPF.Show 1
End Sub

Private Sub cmdPreview_Click()
    If cboPINo.Text = "" Then
        MsgBox "Please, Select P.O. NO.", vbCritical + vbOKOnly, Me.Caption
        If cboPINo.Visible = True Then
        cboPINo.SetFocus
        End If
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans
    
    MainComm.CommandText = "DELETE FROM Temp_Import_Pro_Invoice;"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate, " _
        & " Fdate,Tdate,PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItNAME)" _
        & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
        & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_MainOC.PIDate,Pro_Inv_MainOC.FDate,Pro_Inv_MainOC.TDate,Pro_Inv_SubOC.PIQty, Pro_Inv_SubOC.QtyUnit,Pro_Inv_SubOC.UnitPrice, " _
        & " Pro_Inv_SubOC.ItTotal,Pro_Inv_SubOC.ItDesc, Pro_Inv_SubOC.PRNo,Pro_Inv_SubOC.Currency,Pro_Inv_SubOC.Condition,New_Item_Name.ItemName FROM Pro_Inv_MainOC " _
        & " INNER JOIN Company_Information ON (Pro_Inv_MainOC.compID = Company_Information.compID) " _
        & " INNER JOIN new_Supplier ON (Pro_Inv_MainOC.SuppID = new_Supplier.SupplierID) " _
        & " Inner Join new_PI ON (Pro_Inv_MainOC.PIID = new_PI.PIID) " _
        & " INNER JOIN Pro_Inv_SubOC ON (Pro_Inv_MainOC.PIID = Pro_Inv_SubOC.PIID)" _
        & " Inner Join new_Item_Name On (Pro_Inv_SubOC.ItemID = new_Item_Name.ItemID) " _
        & " WHERE Pro_Inv_MainOC.PIID ='" & cboPINo.Tag & "' AND Pro_Inv_MainOC.compID =" & bytCompID & " AND Pro_Inv_SubOC.DeleteRow = 0;"
    MainComm.Execute
        
    MainConn.CommitTrans
ErrorHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        err.Number = 0
        Exit Sub
    End If
    
       
    If chkpad.Value = 1 Then
    Set REPORT = Appl.OpenReport(findReportPath & "\rptPOpadOC.rpt")
    Else
    Set REPORT = Appl.OpenReport(findReportPath & "\rptPOOC.rpt")
    End If
        
    REPORT.DiscardSavedData
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = REPORT
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1
   
    Set R = Nothing
End Sub

Private Sub cmdRefresh_Click()
    cboItType.ListIndex = -1
    cboItem.ListIndex = -1
    txtDesc.Text = ""
    txtPIDate.Text = ""
    txtUnitPrice.Text = "0.000"
    cboUnit.ListIndex = -1
    'txtPIValue.Text = "0.00"
    flexDetails.Rows = 1
    If txtMode.Text = "NEW" Then
        cboItType.Enabled = True
    End If
End Sub
Private Sub cmdSave_Click()
Dim strPINo As String, strItDesc As String, strPRNo As String, strUnit As String
Dim dblPIQty As Double, curUnitPrice As Double, curTotalAmount As Currency, intItemID As Integer
Dim intPIID As Long, intSLNO As Integer
Dim strCurrency As String, strMatCode As String
    If boEmptyText = True Then Exit Sub
    
     If flexDetails.Row < 1 Then
     
        MsgBox "No Data found !", vbInformation, cnstMsgInfo
        cboItType.SetFocus
        Exit Sub
    End If
    
    If txtMode.Text = "NEW" Then
        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        'strPI = txtPINo.Text
    Else
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        intPIID = cboPINo.Tag
        strPI = cboPINo.Text
    End If

    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans

    If txtMode.Text = "NEW" Then
        '====================
        intYear = Year(txtPIDate.Text)
            strRecord = "SELECT itemShortName FROM new_Item_Type Where itemType = '" & cboItType.Text & "';"
            MainComm.CommandText = strRecord
            Set R = MainComm.Execute
            If R.EOF = False And R.BOF = False Then
                strGroupNameShort = R.Fields(0)
            Else
                strGroupNameShort = "NF"
            End If
        Set R = Nothing
        On Error GoTo er1
er1:
        If err.Number = 94 Then
            Set R = Nothing
            intPI = 1
        Else
            Set R = New ADODB.Recordset
                strRecord = "SELECT Max(PI_No) As MaxPINo FROM New_PI Where ComID = " & bytCompID & " And GroupName = '" & cboItType.Text & "' And PI_Year = " & intYear & ";"
                MainComm.CommandText = strRecord
                Set R = MainComm.Execute
                If R.EOF = False And R.BOF = False Then
                    intPI = R.Fields(0) + 1
                Else
                    intPI = 1
                End If
            Set R = Nothing
        End If
        intPIID = fncNewID("PI")
        txtPINo.Text = strGroupNameShort & bytCompID & fncMRR(intPI) & "-" & intYear
        strPI = txtPINo.Text
        MainComm.CommandText = "INSERT INTO NEW_PI (PIID,comID,GroupName,PI_No,PI_Year,PINo,UserID,PCName,EntryDate,EntryTime)" _
            & " VALUES (" & intPIID & "," & bytCompID & ",'" & cboItType.Text & "','" & fncMRR(intPI) & "','" & intYear & "','" & txtPINo.Text & "'," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "');"
        MainComm.Execute
        '========================='
        MainComm.CommandText = "INSERT INTO Pro_Inv_MainOC(compID,suppID,PIID,PIDate,FDate,TDate,UserID,IDNo,PCName,EntryDate,EntryTime)" _
            & " VALUES (" & bytCompID & "," & findSupplierID(cboSupplier.Text) & "," & intPIID & ",'" & Format(txtPIDate.Text, cnstDtFrmtQ) & "','" & Format(txtFDate.Text, cnstDtFrmtQ) & "','" & Format(txtTDate.Text, cnstDtFrmtQ) & "'," & bytUserID & "," & getComVolID & ",'" & getComName & "','" & Date & "','" & Time & "');"
        MainComm.Execute
    Else
'        MainComm.CommandText = "DELETE FROM Pro_Inv_SubOC WHERE PIID=" & cboPINo.Tag & " And DeleteRow = 0;"
        MainComm.CommandText = "Update Pro_Inv_SubOC set Deleterow=2,UserId=" & bytUserID & ",Pcname='" & getComName & "',Entrydate='" & Date & "',Entrytime='" & Time & "' WHERE PIID=" & cboPINo.Tag & " And DeleteRow = 0;"
        MainComm.Execute
      
        MainComm.CommandText = "UPDATE Pro_Inv_MainOC SET suppID=" & findSupplierID(cboSupplier.Text) & ",PIDate='" & Format(txtPIDate.Text, cnstDtFrmtQ) & "',FDate='" & Format(txtFDate.Text, cnstDtFrmtQ) & "',Tdate='" & Format(txtTDate.Text, cnstDtFrmtQ) & "',UserId=" & bytUserID & ",Pcname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE PIID= " & cboPINo.Tag & " and Deleterow=0;"
        MainComm.Execute
    End If
    With flexDetails
        For inti = 1 To .Rows - 1
            .Row = inti
            ''.Col = 0:    intItemID = findItemID(.Text)
            .Col = 1:    strItDesc = .Text
            .Col = 2:    strPRNo = .Text
            .Col = 3:    dblPIQty = CDbl(.Text)
            .Col = 4:    strUnit = .Text
            .Col = 5:    curUnitPrice = CDbl(.Text)
            .Col = 6:    curTotalAmount = CCur(.Text)
            .Col = 7:    intSLNO = .Text
            .Col = 8:    strCurrency = .Text
            '.Col = 9:    strMatCode = .Text
'        If txtMode.Text = "NEW" Then
            .Col = 9:    intItemID = CInt(.Text)
'        End If
            
            MainComm.CommandText = "INSERT INTO Pro_Inv_SubOC(PIID,ItemID,ItDesc,PRNo,PIQty,QtyUnit,UnitPrice,ItTotal,SLNO,Currency,Condition,UserID,PCName,EntryDate,EntryTime) VALUES (" & intPIID & "," & intItemID & "," _
                  & " '" & strItDesc & "','" & strPRNo & "'," & dblPIQty & ",'" & strUnit & "'," & curUnitPrice & "," & curTotalAmount & "," & intSLNO & ",'" & strCurrency & "','" & txtCondition.Text & "'," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "');"
            MainComm.Execute
        Next inti
    End With
    
    MainConn.CommitTrans
    MsgBox "Transaction update successfully completed. P.O. No = " + strPI, vbInformation + vbOKOnly, Me.Caption
ErrorHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, cnstMsgErDB
        Screen.MousePointer = vbDefault
        MainConn.RollbackTrans
        Exit Sub
    End If
    cmdRefresh_Click
    If txtMode.Text = "NEW" Then
        txtPINo.Text = ""
        cboItType.Enabled = True
        txtPINo.SetFocus
    Else
        cboPINo.Text = ""
'        addSuppPINo cboPINo, cboSupplier.Text 'R1
         addSuppPINoOC cboPINo, cboSupplier.Text
        
        cboPINo.SetFocus
    End If
    Screen.MousePointer = vbDefault
    Set R = Nothing
End Sub
Private Sub cmdShow_Click()
    If Len(cboPINo) = 0 Then
        MsgBox "Select P.O. No", vbInformation, cnstMsgInfo
        cboPINo.SetFocus
        Exit Sub
    End If
    flexDetails.Rows = 1
    
    On Error GoTo ErrorHandler
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    Set R = New ADODB.Recordset
    
    strRecord = "SELECT Pro_Inv_MainOC.COMPID, Pro_Inv_MainOC.PIID,Pro_Inv_MainOC.PIDate," _
        & " new_Item_Name.itemName,Pro_Inv_SubOC.ItDesc,Pro_Inv_SubOC.PRNo,Pro_Inv_SubOC.PIQty," _
        & " Pro_Inv_SubOC.QtyUnit,Pro_Inv_SubOC.UnitPrice,Pro_Inv_SubOC.ItTotal,Pro_Inv_SubOC.SLNO, " _
        & " Pro_Inv_SubOC.Currency,Pro_Inv_SubOC.Condition,Pro_Inv_SubOC.ItemID FROM new_Item_Name INNER JOIN (Pro_Inv_MainOC INNER JOIN Pro_Inv_SubOC" _
        & " ON Pro_Inv_MainOC.PIID = Pro_Inv_SubOC.PIID) ON new_Item_Name.itemID = Pro_Inv_SubOC.itemID " _
        & " Where Pro_Inv_MainOC.PIID =" & cboPINo.Tag & " AND Pro_Inv_MainOC.COMPID =" & bytCompID & " AND Pro_Inv_MainOC.suppID =" & findSupplierID(cboSupplier.Text) & " AND Pro_Inv_SubOC.DeleteRow = 0 Order By Pro_Inv_SubOC.SLNO;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        'txtEditPINo.Text = R![PINo]
        txtPIDate.Text = Format(R![PIDate], cnstDtFrmtI)
        If Not IsNull(R![Condition]) Then txtCondition.Text = R![Condition] Else txtCondition.Text = ""
        With flexDetails
            Do Until R.EOF
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                If Not IsNull(R![itemName]) Then .Text = R![itemName]
                .Col = 1
                If Not IsNull(R![ItDesc]) Then .Text = R![ItDesc]
                .Col = 2
                If Not IsNull(R![PRNo]) Then .Text = R![PRNo]
                .Col = 3
                .Text = R![PIQty]
                .Col = 4
                .Text = R![QtyUnit]
                .Col = 5
                .Text = R![UnitPrice]
                .Col = 6
                .Text = Format(R![ItTotal], "##,##0.00")
                .Col = 7
                .Text = R![SLNo]
                .Col = 8
                .Text = R![Currency]
                .Col = 9
                .Text = R![ITEMID]
                R.MoveNext
            Loop
        End With
        Set R = Nothing
        Set R = New ADODB.Recordset
            strRecord = "SELECT new_Item_Name.itemName FROM new_Item_Name " _
                & " INNER JOIN (Pro_Inv_MainOC INNER JOIN Pro_Inv_SubOC ON Pro_Inv_MainOC.PIID = Pro_Inv_SubOC.PIID) ON new_Item_Name.itemID = Pro_Inv_SubOC.itemID " _
                & " Where Pro_Inv_MainOC.PIID =" & cboPINo.Tag & " AND Pro_Inv_MainOC.COMPID =" & bytCompID & " AND Pro_Inv_MainOC.suppID =" & findSupplierID(cboSupplier.Text) & " AND Pro_Inv_SubOC.DeleteRow = 0 Order By Pro_Inv_SubOC.SLNO ;"
                '& " Group By new_Item_Name.itemName ;"
            R.Open strRecord, MainConn, adOpenStatic
            
            cboItType.Text = findItemType(R![itemName])
        Set R = Nothing
        'txtPIValue.Text = Format(TotalPIValue, "##,##0.00")
    Else
        MsgBox "Record Not Found", vbInformation, cnstMsgInfo
        cmdRefresh_Click
        cboMainPF.ListIndex = -1
        cboUsesPF.ListIndex = -1
        cboPINo.SetFocus
        Exit Sub
    End If
'    R.Close
    'Set R = Nothing
ErrorHandler:
    If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub

Private Sub cmdSubMenu_Click()
frmPOSubMenu.Show 1
End Sub

Private Sub flexDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrorHandler
    flexDetails.Col = 7
    intGSLNO = flexDetails.Text
    'MsgBox intGSLNO
    If KeyCode = 13 Then
        With flexDetails
            If .RowSel = 0 Then Exit Sub
            .Col = 0
            cboItType.Text = findItemType(.Text)
            cboItem.Text = .Text
            .Col = 1
            txtDesc.Text = .Text
            .Col = 2
            txtPrNo.Text = .Text
            .Col = 3
            txtPIQty.Text = .Text
            .Col = 4
            cboUnit.Text = .Text
            .Col = 5
            txtUnitPrice.Text = .Text
            .Col = 6
            txtAmount.Text = .Text
            .Col = 8
            cmbCurrency.Text = .Text
'            .Col = 9
'            txtmatcode.Text = .Text
            bytRowNo = .Row
        End With
        boEdit = True
        cboItem.SetFocus
    End If
    If KeyCode = vbKeyDelete Then
        With flexDetails
            If .RowSel = 0 Then
                Beep
            Else
                If MsgBox("Are you sure to Delete?", vbYesNo + vbQuestion + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
                If (.RowSel = 1) And (.Rows = 2) Then .Rows = 1 Else .RemoveItem .RowSel
                '==== Delete ===='
                If cmdEdit.Caption = "&Edit" Then
                    'With flexDetails
                        For inti = .Row To .Rows - 1
                            .Row = inti
                            .Col = 7:  .Text = .Row
                        Next inti
                    'End With
                End If
                If cmdEdit.Caption = "&New" Then
                    Set MainComm = New ADODB.Command
                    MainComm.ActiveConnection = MainConn
                    MainComm.CommandText = "UPDATE Pro_Inv_SubOC SET DeleteRow = 1 WHERE PIID='" & cboPINo.Tag & "' AND SLNO = " & intGSLNO & ";"
                    MainComm.Execute
                    intGSLNO = 0
                End If
                '==============='
            End If
        End With
    End If
ErrorHandler:
    If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB
End Sub
Private Sub Form_Activate()
'    Call prcConReport(CrystalReport1)
    If strMark = "SUPP" Then
        addSupplierName cboSupplier
    ElseIf strMark = "ITEM" Then
        addItemType cboItType
        addItemName cboItem, cboItType.Text
    End If
End Sub
Private Sub Form_Load()
    Call prcMakeCenter(Me)
    Call addCurrency(cmbCurrency)
    addSupplierName cboSupplier
    addUnit cboUnit
    addItemType cboItType
    
    Set R = New ADODB.Recordset
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
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignCenterCenter
    
      .ColAlignment(3) = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignCenterCenter
      .ColAlignment(5) = flexAlignCenterCenter
      .ColAlignment(6) = flexAlignCenterCenter
      .ColAlignment(7) = flexAlignLeftCenter
      .ColAlignment(8) = flexAlignCenterCenter
      .ColAlignment(9) = flexAlignRightCenter
      .ColWidth(0) = 2000
      .ColWidth(1) = 2000
      .ColWidth(2) = 1500
      .ColWidth(3) = 900
      .ColWidth(4) = 700
      .ColWidth(5) = 1500
      .ColWidth(6) = 1300
      .ColWidth(7) = 0
      .ColWidth(8) = 1000
      .ColWidth(9) = 0 '1000
      .Row = 0
      .Col = 0
      .Text = "Name of Item"
      .Col = 1
      .Text = "Description"
      .Col = 2
      .Text = "PR No"
      .Col = 3
      .Text = "PO Qty"
      .Col = 5
      .Text = "Unit Price"
      .Col = 6
      .Text = "Total"
      .Col = 7
      .Text = "SLNo"
      .Col = 8
      .Text = "Currency"
      .Col = 9
'      .Text = "M Code"
      .Text = "Item ID"
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    strMark = ""
    Set frmProformaInvoiceOC = Nothing
End Sub

Private Sub lblActPF_Click()

End Sub



Private Sub txtAmount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtCharges_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtCharges_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtCharges_LostFocus()
    If Len(txtCharges) = 0 Then txtCharges.Text = "0.00"
    txtCharges.Text = Format(txtCharges.Text, "##0.00")
    txtpivalue.Text = Format(TotalPIValue, "##,##0.00")
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtmatcode_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtmatcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFDate_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtFDate_KeyDown(KeyCode As Integer, Shift As Integer)
   Call prcMoveTab(KeyCode)
End Sub

Private Sub txtFDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub

Private Sub txtFDate_LostFocus()
txtFDate.Text = FormatDate(txtFDate)
End Sub

Private Sub txtPIDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPIQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPrNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPrNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtPrNo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtPrNo_LostFocus()
    txtPrNo.Text = Trim(UCase(txtPrNo.Text))
End Sub
Private Sub txtDesc_GotFocus()
    SendKeys "^{Home}+^{End}"
End Sub
Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txtDesc_LostFocus()
    txtDesc.Text = Trim(UCase(txtDesc.Text))
End Sub
Private Sub txtDiscount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtDiscount_LostFocus()
    If Len(txtDiscount) = 0 Then txtDiscount.Text = "0.00"
    txtDiscount.Text = Format(txtDiscount.Text, "##0.00")
    txtpivalue.Text = Format(TotalPIValue, "##,##0.00")
End Sub
Private Sub txtPIDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtPIDate_KeyPress(KeyAscii As Integer)
    DateFunc KeyAscii
End Sub
Private Sub txtPIDate_LostFocus()
    txtPIDate.Text = FormatDate(txtPIDate)
End Sub
Private Sub txtPINo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtPINo_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtPINo_LostFocus()
    If Len(txtPINo) = 0 Then Exit Sub
    txtPINo.Text = Trim(UCase(txtPINo.Text))
    Set R = New ADODB.Recordset
        strRecord = "SELECT PINo FROM Pro_Inv_MainOC WHERE PINo='" & txtPINo.Text & "';"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            MsgBox "Duplicate P.O. No", vbInformation, cnstMsgInfo
            txtPINo.SetFocus
        End If
        R.Close
    Set R = Nothing
End Sub
Private Sub txtPIQty_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtPIQty_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtPIQty_LostFocus()
    If Len(txtUnitPrice) <> 0 And Len(txtPIQty) <> 0 Then
        txtAmount.Text = Format(txtPIQty.Text * txtUnitPrice.Text, "##0.000")
    Else
        txtAmount.Text = "0.00"
    End If
End Sub
Private Sub txtPIValue_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtSize_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtSize_LostFocus()
    txtSize.Text = Trim(UCase(txtSize.Text))
End Sub

Private Sub txtTDate_GotFocus()
 SendKeys "{Home}+{End}"
End Sub

Private Sub txtTDate_KeyDown(KeyCode As Integer, Shift As Integer)
   Call prcMoveTab(KeyCode)
End Sub

Private Sub txtTDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub

Private Sub txtTDate_LostFocus()
txtTDate.Text = FormatDate(txtTDate)
End Sub

Private Sub txtUnitPrice_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    CurrFunc KeyAscii
End Sub
Private Sub txtUnitPrice_LostFocus()
    If Len(txtUnitPrice) <> 0 And Len(txtPIQty) <> 0 Then
        txtUnitPrice.Text = Format(txtUnitPrice.Text, "##,##0.00000")
        txtAmount.Text = Format(txtPIQty.Text * txtUnitPrice.Text, "##,##0.00000")
    Else
        txtAmount.Text = "0.00"
    End If
End Sub
Sub chkTextNull()
    boEmptyText = False
    If Len(cboSupplier) = 0 Then
        MsgBox "Select Supplier Name", vbInformation, cnstMsgInfo
        cboSupplier.SetFocus
        boEmptyText = True
        Exit Sub
    End If
    If txtMode.Text = "NEW" Then
        If Len(txtPINo) = 0 Then
            MsgBox "Type P.O. No", vbInformation, cnstMsgInfo
            txtPINo.SetFocus
            boEmptyText = True
            Exit Sub
        End If
    Else
        If Len(cboPINo) = 0 Then
            MsgBox "Select P.O. No", vbInformation, cnstMsgInfo
            cboPINo.SetFocus
            boEmptyText = True
            Exit Sub
        End If
    End If
    If Len(txtPIDate) = 0 Then
        MsgBox "Type P.O. Date", vbInformation, cnstMsgInfo
        txtPIDate.SetFocus
        boEmptyText = True
        Exit Sub
    End If
    If flexDetails.Rows = 1 Then
        MsgBox "Record not found to Save/Update", vbInformation, cnstMsgInfo
        boEmptyText = True
        Exit Sub
    End If
End Sub


