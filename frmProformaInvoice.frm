VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmProformaInvoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9420
   ClientLeft      =   2160
   ClientTop       =   1545
   ClientWidth     =   11985
   ControlBox      =   0   'False
   Icon            =   "frmProformaInvoice.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Terms && Conditions "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1698
      Left            =   7436
      TabIndex        =   55
      Top             =   950
      Width           =   4329
      Begin VB.CheckBox chkTax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "4. Govt. Tax and Vat applicable."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   130
         TabIndex        =   53
         Top             =   1300
         Width           =   2795
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1768
         MaxLength       =   50
         TabIndex        =   52
         Top             =   494
         Width           =   2197
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmProformaInvoice.frx":0442
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   572
         Left            =   130
         TabIndex        =   58
         Top             =   728
         Width           =   4095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2. Place of Delivery :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   130
         TabIndex        =   57
         Top             =   494
         Width           =   1625
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1. Specify the purchase order Ref : on your bill/challan."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   130
         TabIndex        =   56
         Top             =   260
         Width           =   4095
      End
   End
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
      TabIndex        =   49
      ToolTipText     =   "Preview PI"
      Top             =   8760
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
      TabIndex        =   20
      ToolTipText     =   "Preview PI"
      Top             =   8760
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
      TabIndex        =   19
      ToolTipText     =   "Close Form"
      Top             =   8745
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
      TabIndex        =   17
      ToolTipText     =   "Refresh"
      Top             =   8760
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
      TabIndex        =   14
      ToolTipText     =   "Save/Update Record"
      Top             =   8745
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
      TabIndex        =   18
      ToolTipText     =   "Delete Record"
      Top             =   8760
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
      TabIndex        =   15
      ToolTipText     =   "New/Edit Record"
      Top             =   8760
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
      TabIndex        =   41
      ToolTipText     =   "Mode of Data Entry"
      Top             =   8775
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
      TabIndex        =   21
      Top             =   8820
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
      Height          =   5655
      Left            =   180
      TabIndex        =   22
      Top             =   2760
      Width           =   11595
      Begin VB.TextBox txtCondition 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1976
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   1560
         Width           =   8762
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
         ItemData        =   "frmProformaInvoice.frx":04D0
         Left            =   6375
         List            =   "frmProformaInvoice.frx":04D2
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
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
         TabIndex        =   4
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
         Height          =   315
         Left            =   2580
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   510
         Width           =   4170
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
         Picture         =   "frmProformaInvoice.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   36
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
         TabIndex        =   7
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
         TabIndex        =   13
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
         TabIndex        =   8
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   9
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
         TabIndex        =   6
         ToolTipText     =   "Item Description"
         Top             =   510
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   3570
         Left            =   105
         TabIndex        =   23
         Top             =   2040
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   6297
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Others Condition : "
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
         Left            =   234
         TabIndex        =   54
         Top             =   1560
         Width           =   1677
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   37
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   900
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   180
      TabIndex        =   30
      Top             =   960
      Width           =   7280
      Begin VB.CheckBox ChkShow 
         Caption         =   "Show"
         Height          =   255
         Left            =   4511
         TabIndex        =   50
         Top             =   1235
         Width           =   728
      End
      Begin VB.CheckBox chkpad 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pad Report"
         Height          =   255
         Left            =   5694
         TabIndex        =   48
         Top             =   1248
         Width           =   1339
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
         Height          =   286
         Left            =   1534
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   793
         Visible         =   0   'False
         Width           =   3978
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
         Left            =   6773
         Picture         =   "frmProformaInvoice.frx":05C6
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Height          =   286
         Left            =   1530
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   5187
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
         Left            =   5694
         TabIndex        =   3
         ToolTipText     =   "Show Record"
         Top             =   780
         Visible         =   0   'False
         Width           =   715
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
         Height          =   286
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         Top             =   780
         Width           =   4004
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
         Height          =   286
         Left            =   1530
         TabIndex        =   2
         Top             =   1215
         Width           =   1830
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
         Height          =   260
         Left            =   1326
         TabIndex        =   47
         Top             =   1196
         Width           =   182
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
         Left            =   1352
         TabIndex        =   46
         Top             =   234
         Width           =   130
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
         Height          =   260
         Left            =   130
         TabIndex        =   35
         Top             =   234
         Width           =   1300
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         Height          =   260
         Left            =   1326
         TabIndex        =   32
         Top             =   806
         Width           =   182
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Left            =   15
      TabIndex        =   42
      Top             =   8520
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   15
      TabIndex        =   40
      Top             =   8550
      Width           =   12240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
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
      Height          =   507
      Index           =   2
      Left            =   195
      TabIndex        =   39
      Top             =   208
      Width           =   5044
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
      TabIndex        =   38
      Top             =   855
      Width           =   12240
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   12120
   End
End
Attribute VB_Name = "frmProformaInvoice"
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
  fncMaxSL = 1
  Set R = New ADODB.Recordset
  Mysql = "SELECT MAX(SLNO) AS NewID FROM Pro_Inv_Sub WHERE PIID = " & Param & ";"
  R.Open Mysql, MainConn, adOpenStatic
  If R.RecordCount > 0 Then fncMaxSL = fncMaxSL + R![NewID]
  R.Close
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
  addSuppPINo cboPINo, cboSupplier.Text
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
  txtPINo.Text = ""
  If Len(cboSupplier) = 0 Then Exit Sub
  addSuppPINo cboPINo, cboSupplier.Text

Else
  cboPINo.Visible = False
  cboPINo.Text = ""
  txtPINo.Visible = True
End If
End Sub
Private Sub cmbCurrency_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub cmdAdd_Click()

Dim ITEM As String, DES As String, PR1 As String, PR2 As String

If Len(cboItem) = 0 Then
  MsgBox "Select Name of Item", vbInformation, cnstMsgInfo
  cboItem.SetFocus
  Exit Sub
End If
If Len(txtPIQty) = 0 Then
  MsgBox "Enter P.O. Qty", vbInformation, cnstMsgInfo
  txtPIQty.SetFocus
  Exit Sub
End If
If Len(txtPrNo) = 0 Then
  MsgBox "Enter PR No.", vbInformation, cnstMsgInfo
  txtPrNo.SetFocus
  Exit Sub
End If
If Len(txtUnitPrice) = 0 Then
  MsgBox "Enter Item Unit Price", vbInformation, cnstMsgInfo
  txtUnitPrice.SetFocus
  Exit Sub
End If
If CCur(txtAmount) = 0 Then
  MsgBox "Amount must be grater than '0'", vbInformation, cnstMsgInfo
  txtUnitPrice.SetFocus
  Exit Sub
End If
If cmbCurrency.Text = "" Then
  MsgBox "Select Currency", vbInformation, cnstMsgInfo
  cmbCurrency.SetFocus
  Exit Sub
End If
    
On Error GoTo ErrorX

intGSLNO = 0
    
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
        MsgBox "Item already Add.", vbInformation, cnstMsgInfo
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
''txtPIValue.Text = Format(TotalPIValue, "##,##0.00")
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

ErrorX:
If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgInfo
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

If Len(cboPINo) = 0 Then
  MsgBox "Select P.O. No", vbInformation, cnstMsgInfo
  cboPINo.SetFocus
  Exit Sub
End If
If flexDetails.Rows = 1 Then
  MsgBox "PO Details not found to Delete", vbInformation, cnstMsgInfo
  Exit Sub
End If

If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub

Set R = New ADODB.Recordset
strRecord = "SELECT DISTINCT A.MRRID,A.MRRNo FROM New_MRR A INNER JOIN MRR_Challan_Sub B ON A.MRRID = B.MRRID WHERE B.PIID = '" & cboPINo.Tag & "';"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  MsgBox "This PO already uses in MRR No - (" & R![MRRNo] & ")", vbCritical, cnstMsgInfo
  cboPINo.SetFocus
  Exit Sub
End If
R.Close

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

MainComm.CommandText = "DELETE FROM New_PI WHERE PIID = '" & cboPINo.Tag & "' AND ComId = " & bytCompID & ";"
MainComm.Execute
MainComm.CommandText = "DELETE FROM Pro_Inv_Main WHERE PIID = '" & cboPINo.Tag & "' AND CompId = " & bytCompID & ";"
MainComm.Execute
MainComm.CommandText = "DELETE FROM Pro_Inv_Sub WHERE PIID = '" & cboPINo.Tag & "';"
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
  err.Number = 0
  Exit Sub
End If

cmdRefresh_Click
cboPINo.Text = ""
txtPIDate.Text = ""
addSuppPINo cboPINo, cboSupplier.Text
cboPINo.SetFocus
Screen.MousePointer = vbDefault
Set R = Nothing
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
  addSuppPINo cboPINo, cboSupplier.Text
  txtPINo.Text = ""
Else
  cmdSave.Caption = "&Save"
  cmdDelete.Enabled = False
  cboPINo.Visible = False
  cmdShow.Visible = False
  txtPINo.Visible = True
  cboItType.Enabled = True
  cmdEdit.Caption = "&Edit"
  txtMode.Text = "NEW"
  cboPINo.Text = ""
End If

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
    Screen.MousePointer = vbHourglass
    
    Set R = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
    MainConn.BeginTrans
    
    MainComm.CommandText = "DELETE FROM Temp_Import_Pro_Invoice;"
    MainComm.Execute
    
    MainComm.CommandText = "INSERT INTO Temp_Import_Pro_Invoice(CompID,CompName,CompAdd,SuppName,SuppAddress,PINo,PIDate,PIQty,Unit,UnitPrice,Total,Remark,PRNo,PICurr,Condition,ItNAME)" _
          & " SELECT Company_Information.CompID,Company_Information.CompName,Company_Information.CompAdd,new_Supplier.SuppName, " _
          & " new_Supplier.suppAddress,new_PI.PINo,Pro_Inv_Main.PIDate,Pro_Inv_Sub.PIQty, Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice, " _
          & " Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.ItDesc, Pro_Inv_Sub.PRNo,Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition,New_Item_Name.ItemName FROM Pro_Inv_Main " _
          & " INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
          & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID) Inner Join new_PI ON (Pro_Inv_Main.PIID = new_PI.PIID) " _
          & " INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID)" _
          & " Inner Join new_Item_Name On (Pro_Inv_Sub.ItemID = new_Item_Name.ItemID) " _
          & " WHERE Pro_Inv_Main.PIID ='" & cboPINo.Tag & "' AND Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0 ORDER BY Pro_Inv_Sub.PRNo,New_Item_Name.ItemName;"
    MainComm.Execute
    
    MainComm.CommandText = "UPDATE Temp_Import_Pro_Invoice SET Condition = '" & txtCondition.Text & "';"
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
      Set REPORT = Appl.OpenReport(findReportPath & "\rptPOpad.rpt")
    Else
      Set REPORT = Appl.OpenReport(findReportPath & "\rptPO.rpt")
    End If
    
    REPORT.FormulaFields.GetItemByName("DelPlace").Text = "'" & txtPlace.Text & "'"
    If chkTax.Value = 1 Then REPORT.FormulaFields.GetItemByName("GovtTax").Text = "'" & chkTax.Caption & "'"

    REPORT.DiscardSavedData
    frmMainReport.CRViewer.Refresh
    frmMainReport.CRViewer.ReportSource = REPORT
    frmMainReport.CRViewer.ViewReport
    frmMainReport.CRViewer.Zoom (82)
    frmMainReport.Show 1
   
    Screen.MousePointer = vbDefault
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
  If txtMode.Text = "NEW" Then cboItType.Enabled = True
End Sub
Private Sub cmdSave_Click()

Dim strPINo As String, strItDesc As String, strPRNo As String, strUnit As String
Dim dblPIQty As Double, curUnitPrice As Double, curTotalAmount As Currency, intItemID As Integer
Dim intPIID As Long, intSLNO As Integer
Dim strCurrency As String, strMatCode As String
    
If boEmptyText = True Then Exit Sub

If flexDetails.Rows = 1 Then
  MsgBox "PO Details not found", vbInformation, cnstMsgInfo
  cboItType.SetFocus
  Exit Sub
End If
    
Set R = New ADODB.Recordset
    
If txtMode.Text = "NEW" Then
  If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  'strPI = txtPINo.Text
  
''  strRecord = "SELECT PINo FROM Pro_Inv_Main WHERE PINo = '" & txtPINo.Text & "';"
''  R.Open strRecord, MainConn, adOpenStatic
''  If R.RecordCount > 0 Then
''    MsgBox "Duplicate PO No.", vbInformation, cnstMsgInfo
''    txtPINo.SetFocus
''    Exit Sub
''  End If
''  R.Close

  strGroupNameShort = "NF"
  strRecord = "SELECT itemShortName FROM new_Item_Type WHERE itemType = '" & cboItType.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  If R.EOF = False And R.BOF = False Then strGroupNameShort = R![itemShortName]
  R.Close

  intPI = 1
  intYear = Year(txtPIDate.Text)
  strRecord = "SELECT Max(PI_No) As MaxPINo FROM New_PI WHERE ComID = " & bytCompID & " AND GroupName = '" & cboItType.Text & "' AND PI_Year = " & intYear & ";"
  R.Open strRecord, MainConn, adOpenStatic
  If R.EOF = False And R.BOF = False Then intPI = IIf(IsNull(R![MaxPINo]), 0, R![MaxPINo]) + 1
  R.Close
  
  intPIID = fncNewID("PI")
  txtPINo.Text = strGroupNameShort & bytCompID & fncMRR(intPI) & "-" & intYear
  strPI = txtPINo.Text
        
Else
  If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
  intPIID = cboPINo.Tag
  strPI = cboPINo.Text
End If

On Error GoTo ErrorHandler
Screen.MousePointer = vbHourglass

Set MainComm = New ADODB.Command
MainComm.ActiveConnection = MainConn

MainConn.BeginTrans

If txtMode.Text = "NEW" Then
  MainComm.CommandText = "INSERT INTO NEW_PI (PIID,comID,GroupName,PI_No,PI_Year,PINo,UserID,PCName,EntryDate,EntryTime)" _
        & " VALUES (" & intPIID & "," & bytCompID & ",'" & cboItType.Text & "','" & fncMRR(intPI) & "','" & intYear & "','" & txtPINo.Text & "'," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "');"
  MainComm.Execute

  MainComm.CommandText = "INSERT INTO Pro_Inv_Main(compID,suppID,PIID,PIDate,UserID,IDNo,PCName,EntryDate,EntryTime)" _
        & " VALUES (" & bytCompID & "," & findSupplierID(cboSupplier.Text) & "," & intPIID & ",'" & Format(txtPIDate.Text, cnstDtFrmtQ) & "'," & bytUserID & "," & getComVolID & ",'" & getComName & "','" & Date & "','" & Time & "');"
  MainComm.Execute

Else
  MainComm.CommandText = "UPDATE Pro_Inv_Main SET suppID = " & findSupplierID(cboSupplier.Text) & ",PIDate = '" & Format(txtPIDate.Text, cnstDtFrmtQ) & "',UserId=" & bytUserID & ",Pcname='" & getComName & "',EntryDate = '" & Date & "',EntryTime = '" & Time & "' WHERE PIID = " & cboPINo.Tag & ";"
  MainComm.Execute
  
  MainComm.CommandText = "DELETE FROM Pro_Inv_Sub WHERE PIID = " & cboPINo.Tag & ";"
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
    .Col = 9:    intItemID = CInt(.Text)

    MainComm.CommandText = "INSERT INTO Pro_Inv_Sub(PIID,ItemID,ItDesc,PRNo,PIQty,QtyUnit,UnitPrice,ItTotal,SLNO,Currency,Condition,UserID,PCName,EntryDate,EntryTime) VALUES (" & intPIID & "," & intItemID & ",'" & strItDesc & "'," _
          & " '" & strPRNo & "'," & dblPIQty & ",'" & strUnit & "'," & curUnitPrice & "," & curTotalAmount & "," & intSLNO & ",'" & strCurrency & "','" & txtCondition.Text & "'," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "');"
    MainComm.Execute
  Next inti
End With

MainConn.CommitTrans

ErrorHandler:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  MainConn.RollbackTrans
  err.Number = 0
  Exit Sub
End If

MsgBox "Transaction Update successfully completed. P.O. No = " + strPI, vbInformation + vbOKOnly, Me.Caption

On Error GoTo ErrorX

cmdRefresh_Click
If txtMode.Text = "NEW" Then
  txtPINo.Text = ""
  cboItType.Enabled = True
  txtPINo.SetFocus
Else
  cboPINo.Text = ""
  addSuppPINo cboPINo, cboSupplier.Text
  cboPINo.SetFocus
End If

txtPlace.Text = "At our Factory"

ErrorX:
If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgInfo
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
    
    strRecord = "SELECT Pro_Inv_Main.COMPID, Pro_Inv_Main.PIID,Pro_Inv_Main.PIDate,Pro_Inv_Sub.itemID," _
        & " new_Item_Name.itemName,Pro_Inv_Sub.ItDesc,Pro_Inv_Sub.PRNo,Pro_Inv_Sub.PIQty," _
        & " Pro_Inv_Sub.QtyUnit,Pro_Inv_Sub.UnitPrice,Pro_Inv_Sub.ItTotal,Pro_Inv_Sub.SLNO, " _
        & " Pro_Inv_Sub.Currency,Pro_Inv_Sub.Condition FROM new_Item_Name INNER JOIN (Pro_Inv_Main INNER JOIN Pro_Inv_Sub" _
        & " ON Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID) ON new_Item_Name.itemID = Pro_Inv_Sub.itemID " _
        & " Where Pro_Inv_Main.PIID = " & cboPINo.Tag & " AND Pro_Inv_Main.COMPID = " & bytCompID & " AND Pro_Inv_Main.suppID = " & findSupplierID(cboSupplier.Text) & " AND Pro_Inv_Sub.DeleteRow = 0;"
    
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
                .Text = R![PIQTY]
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
                ''.Text = R![MatCode] & ""
                .Text = R![ITEMID]
                R.MoveNext
            Loop
        End With
        Set R = Nothing
        Set R = New ADODB.Recordset
            strRecord = "SELECT new_Item_Name.itemName FROM new_Item_Name " _
                & " INNER JOIN (Pro_Inv_Main INNER JOIN Pro_Inv_Sub ON Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID) ON new_Item_Name.itemID = Pro_Inv_Sub.itemID " _
                & " Where Pro_Inv_Main.PIID =" & cboPINo.Tag & " AND Pro_Inv_Main.COMPID =" & bytCompID & " AND Pro_Inv_Main.suppID =" & findSupplierID(cboSupplier.Text) & " AND Pro_Inv_Sub.DeleteRow = 0 Order By Pro_Inv_Sub.SLNO ;"
                '& " Group By new_Item_Name.itemName ;"
            R.Open strRecord, MainConn, adOpenStatic
            If R.RecordCount > 0 Then
              cboItType.Text = findItemType(R![itemName])
            End If
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
    If err.Number <> 0 Then
      MsgBox err.Description, vbCritical, cnstMsgErDB
      err.Number = 0
    End If
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
                    MainComm.CommandText = "UPDATE Pro_Inv_Sub SET DeleteRow = 1 WHERE PIID='" & cboPINo.Tag & "' AND SLNO = " & intGSLNO & ";"
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
      .ColWidth(9) = 0 ''1000
      .Row = 0
      .Col = 0
      .Text = "Name of Item"
      .Col = 1
      .Text = "Description"
      .Col = 2
      .Text = "PR No"
      .Col = 3
      .Text = "PO Qty"
      .Col = 4
      .Text = "Unit"
      .Col = 5
      .Text = "Unit Price"
      .Col = 6
      .Text = "Total"
      .Col = 7
      .Text = "SLNo"
      .Col = 8
      .Text = "Currency"
      .Col = 9
      .Text = "Item ID"
    End With
    txtPlace.Text = "At our Factory"
If strUserName = "PACIFIC" Then cmdDelete.Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  strMark = ""
  Set frmProformaInvoice = Nothing
End Sub
Private Sub txtAmount_GotFocus()
  txtGotFocus txtAmount
End Sub
Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtAmount_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtPIDate_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtPIQty_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtPrNo_GotFocus()
  txtGotFocus txtPrNo
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
  txtGotFocus txtDesc
End Sub
Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtDesc_LostFocus()
  txtDesc.Text = Trim(UCase(txtDesc.Text))
End Sub
Private Sub txtPIDate_GotFocus()
  txtGotFocus txtPIDate
End Sub
Private Sub txtPIDate_KeyPress(KeyAscii As Integer)
  DateFunc KeyAscii
End Sub
Private Sub txtPIDate_LostFocus()
  txtPIDate.Text = FormatDate(txtPIDate)
End Sub
Private Sub txtPINo_GotFocus()
  txtGotFocus txtPINo
End Sub
Private Sub txtPINo_KeyPress(KeyAscii As Integer)
  SingleCodeFunc KeyAscii
End Sub
Private Sub txtPINo_LostFocus()
  If Len(txtPINo) = 0 Then Exit Sub
  txtPINo.Text = Trim(UCase(txtPINo.Text))
End Sub
Private Sub txtPIQty_GotFocus()
  txtGotFocus txtPIQty
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
Private Sub txtUnitPrice_GotFocus()
  txtGotFocus txtUnitPrice
End Sub
Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
  Call prcMoveTab(KeyCode)
End Sub
Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
  CurrFunc KeyAscii
End Sub
Private Sub txtUnitPrice_LostFocus()
  If Len(txtUnitPrice) <> 0 And Len(txtPIQty) <> 0 Then
    txtUnitPrice.Text = Format(txtUnitPrice.Text, "##,##0.0000")
    txtAmount.Text = Format(txtPIQty.Text * txtUnitPrice.Text, "##,##0.0000")
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
