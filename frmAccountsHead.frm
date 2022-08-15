VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAccountsHead 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5395
   ClientLeft      =   2210
   ClientTop       =   1794
   ClientWidth     =   6695
   ControlBox      =   0   'False
   Icon            =   "frmAccountsHead.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5395
   ScaleWidth      =   6695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
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
      Height          =   525
      Left            =   4680
      TabIndex        =   5
      Top             =   4680
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
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
      Height          =   525
      Left            =   3360
      TabIndex        =   4
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAF5F5&
      Height          =   3645
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   6675
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E26101&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   3270
         Width           =   555
      End
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   2880
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   6585
         _ExtentX        =   11622
         _ExtentY        =   5080
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   255
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   -2147483632
         ForeColorFixed  =   16776960
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.1509
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E26101&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Entries  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   3240
         Width           =   6645
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0C000&
      Caption         =   "&Add"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   4680
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
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
      Height          =   525
      Left            =   720
      TabIndex        =   0
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   4440
      Width           =   6795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name Of The Accounts Head"
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
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   5010
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6675
   End
End
Attribute VB_Name = "frmAccountsHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  strFormName = "Acc Head For BillDetails"
  strMark = "HEAD"
  flagMark = "INSERT"
  strValue = ""
  
  frmNewAccountsHead.Show 1
  Call prcLoadData
End Sub
Private Sub cmdDelete_Click()
On Error GoTo x
    If MsgBox("Are you sure to delete?" & vbCrLf & "district '" & flexDetails.Text & "' ?", vbQuestion + vbYesNo, AppTitle) = vbNo Then Exit Sub
        con.BeginTrans
            com.CommandText = "DELETE FROM AccountsHead WHERE HeadName='" & flexDetails.Text & "';"
            com.Execute
        con.CommitTrans
    Call prcLoadData
    Call prcGridSetting(Me, flexDetails)
    'Call prcGridFormatting(Me, flexDetails)
Exit Sub
x:
    con.RollbackTrans
    MsgBox err.Description, vbCritical, cnstMsgErDB
Exit Sub
End Sub
Private Sub cmdEdit_Click()
    strFormName = "Acc Head For BillDetails"
    strMark = "HEAD"
    flagMark = "EDIT"
    With frmNewAccountsHead
        flexDetails.Col = 0:    .Tag = flexDetails.Text
        flexDetails.Col = 1:
            .txtAccountsHead.Text = flexDetails.Text
            .txtAccountsHead.Tag = flexDetails.Text
        .Caption = "Edit Accounts Head"
        .Show 1
    End With
    Call prcLoadData
    Call prcGridSetting(Me, flexDetails)
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo x
    Call prcMakeCenter(Me)
    Call prcGridInitialize
    Call prcLoadData
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAccountsHead = Nothing
End Sub

Private Sub prcLoadData()
Dim inti As Integer
On Error GoTo x
Set RS = New ADODB.Recordset
    Mysql = "SELECT HeadID, HeadName FROM AccountsHead ORDER BY HeadID;"
'    com.CommandText = Mysql
'    Set RS = com.Execute
    RS.Open Mysql, MainConn, adOpenStatic
        
    inti = 0
    flexDetails.Rows = 0
    If RS.EOF = False And RS.BOF = False Then
        With flexDetails
            Do While Not RS.EOF
                inti = inti + 1
                
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .Col = 0:  .Text = RS(0).Value
                .Col = 1:  .Text = RS(1).Value
            RS.MoveNext
            Loop
        End With
        txtTotal.Text = inti
    End If
    Call prcGridSetting(Me, flexDetails)
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Private Sub prcGridInitialize()
    With flexDetails
        .Row = 0
        .Cols = 2
        
        .ColWidth(0) = 0: .ColWidth(1) = 5985
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
    End With
End Sub

