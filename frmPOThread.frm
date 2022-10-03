VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPOThread 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9255
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
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
      Height          =   442
      Left            =   6840
      TabIndex        =   30
      ToolTipText     =   "Save/Update Record"
      Top             =   8640
      Width           =   1185
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0C000&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   442
      Left            =   5280
      TabIndex        =   29
      ToolTipText     =   "Edit/Delete Record"
      Top             =   8655
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
      Height          =   442
      Left            =   8520
      TabIndex        =   28
      ToolTipText     =   "Edit/Delete Record"
      Top             =   8640
      Width           =   1185
   End
   Begin VB.ComboBox cboPINo 
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
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.CommandButton cmdimageSelect 
      BackColor       =   &H8000000D&
      Caption         =   "UPLOAD IMAGE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   403
      Left            =   7560
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   26
      Top             =   7826
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   3960
      ScaleHeight     =   6705
      ScaleWidth      =   7665
      TabIndex        =   25
      Top             =   960
      Width           =   7695
      Begin VB.Image Image1 
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00C0C000&
      Caption         =   "Pre&view"
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
      Height          =   442
      Left            =   2400
      TabIndex        =   24
      ToolTipText     =   "Preview Report"
      Top             =   8640
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "NEW"
      ToolTipText     =   "Mode of Data Entry"
      Top             =   8760
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "ZOOM OUT/IN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   546
      Left            =   4560
      TabIndex        =   19
      Top             =   7800
      Width           =   2295
      Begin VB.CommandButton cmdZoomMinus 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   169
         Width           =   495
      End
      Begin VB.CommandButton cmdZoomEqual 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   169
         Width           =   495
      End
      Begin VB.CommandButton CmdZoomPlus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   169
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUPPLIER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   3615
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
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   570
         Width           =   3150
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
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
         Left            =   360
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1200
         Width           =   3030
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SUPPLIER NAME"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADDRESS"
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
         TabIndex        =   17
         Top             =   960
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   3495
      Begin VB.ComboBox cboeditPO 
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
         Left            =   240
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtPINo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         Top             =   480
         Width           =   2955
      End
      Begin VB.TextBox txtpodate 
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
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "P.O.  No : "
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
         TabIndex        =   13
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "P.O. DATE : "
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
         TabIndex        =   12
         Top             =   960
         Width           =   930
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   3495
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
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2940
      End
      Begin VB.TextBox txtBuyer 
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
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ITEM GROUP "
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BUYER REF."
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1605
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
      Height          =   442
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "Delete Record"
      Top             =   8640
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   7080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   240
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO (THREAD) "
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
      Index           =   1
      Left            =   4560
      TabIndex        =   35
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11715
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PO (THREAD) "
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
      Left            =   3720
      TabIndex        =   33
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E26101&
      Height          =   855
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   8460
      Width           =   11715
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
      TabIndex        =   31
      Top             =   840
      Width           =   5850
   End
End
Attribute VB_Name = "frmPOThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim opt As String
      Dim strPath As String
      Dim cx As Integer
      Dim serialId As Integer
      Dim RS As Recordset
      Dim ChunkSize As Integer
      Dim conChunkSize As Integer
      Dim i As Long
      Dim ZoomSizeOrgWid As String
      Dim ZoomSizeOrgHgh As String
      Dim l As Double, a As Double, x As Double
      Dim Select_Image As Boolean
      Dim strPI As String
      Dim intPI As Integer
      Dim b As String
       Dim min As String
    Dim c As String
    Dim y As String
    Dim s As String
    Dim h As String
    Dim mon As String
    Dim d As String
   
      Dim bmphold As StdPicture
      Private Const BLOCK_SIZE = 10000




Dim FormOrgWid As String
Private Sub cboeditPO_Click()

cboeditPO.Tag = cboeditPO.ItemData(cboeditPO.ListIndex)






' cboItType.Text = ""
 txtpodate.Text = ""
 txtBuyer.Text = ""

 Set R = New ADODB.Recordset
 
 
 
 
        strRecord = "SELECT ItemGr,PODATE,BuyerRef FROM  [ThreadInfo] WHERE COMPID=" & bytCompID & "  AND suppid=" & cboSupplier.Tag & " and POID=" & cboeditPO.Tag & "  AND DELETEROW=0 "


      R.Open strRecord, MainConn, adOpenStatic
      
      
    If R.RecordCount > 0 Then
         cboItType.Text = R![ItemGr]
         txtpodate.Text = R![PODATE]
         txtBuyer.Text = R![BuyerRef]
    R.Close
    End If

End Sub

Private Sub cboItType_Click()
cboItType.Tag = cboItType.ItemData(cboItType.ListIndex)
End Sub

Private Sub cboItType_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub cboSupplier_Click()

If Len(cboSupplier) = 0 Then Exit Sub
   cboSupplier.Text = UCase(Trim(cboSupplier.Text))
  
    cboSupplier.Tag = cboSupplier.ItemData(cboSupplier.ListIndex)
  
  
  Set R = New ADODB.Recordset
  
  strRecord = "SELECT suppName,suppAddress,suppCountry FROM new_Supplier WHERE suppName='" & cboSupplier.Text & "';"
  R.Open strRecord, MainConn, adOpenStatic
  
  If R.RecordCount > 0 Then
  
   ' txtEditSupplier.Text = R![suppName]
    txtAddress.Text = R![suppAddress] & ""
   'IIf IsNull(R![suppCountry]), cboCountry.ListIndex = -1, cboCountry.Text = R![suppCountry]
  End If
  
  R.Close
  
  
  Set R = Nothing
  

    If txtMode.Text = "EDIT" Then
       
       EditSuppPINo cboeditPO, cboSupplier.Tag
       
       Else
       
     ' addSuppPINo cboPINo, cboSupplier.Text
      
      End If

End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
 Call prcMoveTab(KeyCode)
End Sub

Private Sub cboSupplier_LostFocus()

 cboSupplier.Text = Trim(UCase(cboSupplier.Text))
  If Len(cboSupplier) > 50 Then
    MsgBox "Maximum size is '50'", vbInformation, cnstMsgInfo
    cboSupplier.SetFocus
  End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
 On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    
     
    
     If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo + vbDefaultButton2, cnstMsgQ) = vbNo Then Exit Sub
 
    
       MainConn.BeginTrans
    
    
                 MainComm.CommandText = "UPDATE [ThreadInfo] SET USERID='" & bytUserID & "',PCNAME='" & getComName & "',ENTRYDATE='" & Date & "',ENTRYTIME='" & Time & "' ,DELETEROW=1 WHERE COMPID=" & bytCompID & "  AND SUPPID=" & cboSupplier.Tag & " AND POID=" & cboeditPO.Tag & " AND DELETEROW=0 "
                 MainComm.Execute
    
    
    
        MainConn.CommitTrans
        
           Image1.Picture = Nothing
            Select_Image = False
        
        
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
End Sub

Private Sub cmdEdit_Click()
 If txtMode.Text = "NEW" Then
        cmdSave.Caption = "&Update"
         cmdDelete.Enabled = True
         txtPINo.Visible = False
         cboeditPO.Top = txtPINo.Top
         cboeditPO.Visible = True
'        cmdShow.Top = cboPINo.Top
'        cmdShow.Visible = True
        cmdEdit.Caption = "&New"
        cboItType.Enabled = False
        txtMode.Text = "EDIT"
        cmdReport.Enabled = True
       ' addSupplier cboSupplier
'        addSuppPINo cboPINo, cboSupplier.Text
    Else
         cmdSave.Caption = "&Save"
          cmdDelete.Enabled = False
          cboeditPO.Visible = False
'        cmdShow.Visible = False
         txtPINo.Visible = True
         cboItType.Enabled = True
         cmdEdit.Caption = "&Edit"
         txtMode.Text = "NEW"
         cmdReport.Enabled = False
    End If
End Sub

Private Sub cmdimageSelect_Click()
Dim str1 As String
Dim file_num As String
Dim file_length As String
Dim bytes() As Byte
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long




    Dlg.Flags = _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNExplorer
    Dlg.CancelError = True
    
    '*.TIF;
    Dlg.Filter = "Graphics Files|*.bmp;*.ico;*.jpg;*.gif"



    On Error Resume Next
    Dlg.ShowOpen
    
    
    
    
    
    If err.Number = cdlCancel Then
        Exit Sub
    ElseIf err.Number <> 0 Then
        MsgBox "Error " & Format$(err.Number) & _
            " selecting file." & vbCrLf & err.Description
        Exit Sub
    End If

    
  

'BLOCK_SIZE = 10000


    ' Open the picture file.
    
    
    
    file_num = FreeFile
    
    
    
    Open Dlg.FileName For Binary Access Read As #file_num

    file_length = LOF(file_num)
    
        
    Image1.Picture = LoadPicture(Dlg.FileName)
    
    If file_length > 0 Then
        
        
        num_blocks = file_length / BLOCK_SIZE
        left_over = file_length Mod BLOCK_SIZE



    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500


MainComm.CommandText = "delete  from spPhoto;"
MainComm.Execute

    
        Set RS = New Recordset
        
        str1 = "select* from SpPhoto"
        RS.Open str1, MainConn, 3, 3
        
        
        
        RS.AddNew
        
        RS.Fields("SpLength").Value = file_length

        ReDim bytes(BLOCK_SIZE)
        
        For block_num = 1 To num_blocks
            Get #file_num, , bytes()
            RS.Fields("SpPhoto").AppendChunk bytes()
        Next block_num

        If left_over > 0 Then
            ReDim bytes(left_over)
            Get #file_num, , bytes()
            RS.Fields("SpPhoto").AppendChunk bytes()
        End If

          RS.Update
        Close #file_num
        RS.Close
        End If
   

Select_Image = True

Image1.Stretch = False

        ZoomSizeOrgWid = Image1.Width
        ZoomSizeOrgHgh = Image1.Height
        
 Image1.Stretch = True
 
 Image1.Width = Picture2.Width
 Image1.Height = Picture2.Height
End Sub

Private Sub cmdReport_Click()

 '---------
Dim str1 As String
Dim file_num As String
Dim file_length As String
Dim bytes() As Byte
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long
Dim aa As String


'------------




If Len(cboSupplier.Text) = 0 Then

        MsgBox "Please Select SUPPLIER NAME", vbInformation, cnstMsgInfo
        cboSupplier.SetFocus
        Exit Sub
    
  End If
  
  
  
  
  If Len(cboeditPO.Text) = 0 Then
  
        MsgBox "SELECT  PO NO.", vbInformation, cnstMsgInfo
        cboeditPO.SetFocus
        Exit Sub
    
  End If
  
  

 
  Set R = New ADODB.Recordset
  Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn
    MainComm.CommandTimeout = 500
 
 
 
 MainComm.CommandText = "delete  from spPhoto;"
 MainComm.Execute

 MainComm.CommandText = "delete  from Temp_ThreadInfo;"
 MainComm.Execute

 
 
 
 '------------------FIND OUT PATH
aa = ""

  strRecord = "SELECT  Path  FROM [ThreadInfo] WHERE POID=" & cboeditPO.Tag & " AND PODATE='" & txtpodate.Text & "'  AND DELETEROW=0  ;"
  
  R.Open strRecord, MainConn, adOpenStatic
  
    If R.RecordCount > 0 Then
    ' --IIf(IsNull(R![Invvalue]), "0.00", R![Invvalue])
     aa = IIf(IsNull(R![Path]), "", R![Path])
     
    End If
  
  R.Close
  Set R = Nothing




'----------------------------
 
 
  file_num = FreeFile
    
    
 
  
 
  
  If aa <> "" Then
  
  
 Open aa For Binary Access Read As #file_num
 
 
 

    file_length = LOF(file_num)
    
    Set StdPicture = LoadPicture(aa)
    
    

    
  
  
   If file_length > 0 Then
        
        
        num_blocks = file_length / BLOCK_SIZE
        left_over = file_length Mod BLOCK_SIZE



   


'MainComm.CommandText = "delete  from spPhoto;"
'MainComm.Execute
'
'MainComm.CommandText = "delete  from Temp_ThreadInfo;"
'MainComm.Execute

    
        Set RS = New Recordset
        
        str1 = "select* from SpPhoto"
        RS.Open str1, MainConn, 3, 3
        
        
        
        RS.AddNew
        
        RS.Fields("SpLength").Value = file_length

        ReDim bytes(BLOCK_SIZE)
        
        For block_num = 1 To num_blocks
            Get #file_num, , bytes()
            RS.Fields("SpPhoto").AppendChunk bytes()
        Next block_num

        If left_over > 0 Then
            ReDim bytes(left_over)
            Get #file_num, , bytes()
            RS.Fields("SpPhoto").AppendChunk bytes()
        End If

          RS.Update
        Close #file_num
        RS.Close
        End If
   
  
  
'-------------------


  Set REPORT = Appl.OpenReport(findReportPath & "\rptThread.rpt")
 
        MainComm.CommandText = " INSERT INTO Temp_ThreadInfo Select * From ThreadInfo Where CompID=" & bytCompID & "  AND POID=" & cboeditPO.Tag & "  AND DELETEROW=0 "
        MainComm.Execute
        
        MainComm.CommandText = "Update Temp_ThreadInfo set Temp_ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM Temp_ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & cboeditPO.Text & "' AND PODATE='" & txtpodate.Text & "' AND Temp_ThreadInfo.DELETEROW=0 "
        MainComm.Execute
  
  Else
  
  
   Set REPORT = Appl.OpenReport(findReportPath & "\rptThread.rpt")
 
        MainComm.CommandText = " INSERT INTO Temp_ThreadInfo Select * From ThreadInfo Where CompID=" & bytCompID & "  AND POID=" & cboeditPO.Tag & "  AND DELETEROW=0 "
        MainComm.Execute
        
        MainComm.CommandText = "Update Temp_ThreadInfo set Temp_ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM Temp_ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & cboeditPO.Text & "' AND PODATE='" & txtpodate.Text & "' AND Temp_ThreadInfo.DELETEROW=0 "
        MainComm.Execute
  
  
 End If
  
'  MainComm.CommandText = "Update ThreadInfo set ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & cboeditPO.Text & "' AND DELETEROW=0 "
'                  MainComm.Execute
                 
                 
      frmMainReport.CRVIEWER.Refresh
      frmMainReport.CRVIEWER.ReportSource = REPORT
      frmMainReport.CRVIEWER.ViewReport
      frmMainReport.CRVIEWER.Zoom (82)
      frmMainReport.Show 1
      
'      Else
'      MsgBox "RECORD NOT FOUND", vbInformation, cnstMsgInfo
      Exit Sub
      
'      End If
'      R.Close

End Sub

Private Sub cmdSave_Click()

Dim lngLastID As Integer
Dim ComADD As String
Dim strPINo As String
Dim pt As String

   


  On Error GoTo ErrorHandler
  '  Screen.MousePointer = vbHourglass
    
    Set R = New ADODB.Recordset
    Set R1 = New ADODB.Recordset
    Set MainComm = New ADODB.Command
    MainComm.ActiveConnection = MainConn


 
 
   min = Minute(Time)
   s = Second(Time)
   h = Hour(Time)
   y = Year(Date)
   mon = Month(Date)
   d = Day(Date)



 

If Len(cboSupplier.Text) = 0 Then

        MsgBox "Please Select SUPPLIER NAME", vbInformation, cnstMsgInfo
        cboSupplier.SetFocus
        Exit Sub
    
  End If
  
  If Len(txtpodate.Text) = 0 Then

        MsgBox "Please PUT THE PO DATE", vbInformation, cnstMsgInfo
        txtpodate.SetFocus
        Exit Sub
    
  End If
  
  
  
  
  
  
  
  
If Select_Image = False Then

        MsgBox "Please Select Image", vbInformation, cnstMsgInfo
        cboSupplier.SetFocus
        Exit Sub
    
  End If
  
  '------SAVE PATH IN FOLDER--
  
  Dim PicPath As String, strPicPath As String, strSerPicPath As String

strPicPath = ""
strSerPicPath = ""
strSerPicPath = "\\172.16.220.30\export$\Inventory Image\InventoryThread"
strPicPath = "\\172.16.220.30\export$\Inventory Image\Picture\"


    
    
   

  '------END
  
  
  '-----------
  
   ComADD = ""
        
    strRecord = "SELECT CompAdd FROM Company_Information WHERE CompID= " & bytCompID & " "
         
            
            
            R1.Open strRecord, MainConn, adOpenStatic
            
             If R1.RecordCount > 0 Then
             
                 ComADD = R1![CompAdd]
               
             End If
             
             Set R1 = Nothing
  
  '-----------
  
  
  
   If txtMode.Text = "NEW" Then
   
   
   
        If MsgBox("Are you sure to Save", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
       
        
        
 
        
        
        
    Else
    
    
    
    
    
        If MsgBox("Are you sure to Update", vbQuestion + vbYesNo + vbDefaultButton1, cnstMsgQ) = vbNo Then Exit Sub
        
        intPIID = cboeditPO.Tag
        strPI = cboeditPO.Text
        
        
        
    End If
    
    
  
  
   
    
    
    
     MainConn.BeginTrans
   '---------------
   
   
   
   
   
   
   
   If txtMode.Text = "NEW" Then
   
        '====================
        
        
        intYear = Year(txtpodate.Text)
        
        
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
               
            
                           strRecord = "SELECT CASE WHEN  Max(PI_No) IS NULL THEN 0 ELSE Max(PI_No) END   MaxPINo FROM New_PI Where ComID = " & bytCompID & " And GroupName = '" & cboItType.Text & "' And PI_Year = " & intYear & ";"
                            MainComm.CommandText = strRecord
                  
                
                
                       Set R = MainComm.Execute
                     
                
                
                                If R.EOF = False And R.BOF = False Then
                                
                                    intPI = R.Fields(0) + 1
                                    
                                Else
                                
                                    intPI = 1
                                    
                                End If
                                
                                R.Close
                                
                                
                
                       Set R = Nothing
            
            
        End If
        
        

        
     
        
               intPIID = fncNewID("PI")
        
        
                txtPINo.Text = strGroupNameShort & bytCompID & fncMRR(intPI) & "-" & intYear
                strPI = txtPINo.Text
                
                
                SavePicture Image1.Picture, strSerPicPath & "\" & "C" & bytCompID & "Y" & y & "M" & mon & "D" & d & "H" & h & "M" & min & "S" & s & "P" & intPIID & ".jpg"
'                SavePicture Image1.Picture, strPicPath & "\Picture\" & "C_" & bytCompID & "_D_" & Date & "_PO_" & txtPINo.Text & "_P_" & intPIID & "_I_" & Image1 & ".jpg"
'                PicPath = strPicPath & "\Picture\" & "C_" & bytCompID & "_D_" & Date & "_PO_" & txtPINo.Text & "_P_" & intPIID & "_I_" & Image1 & ".jpg"
                SavePicture Image1.Picture, strPicPath & "\" & "C" & bytCompID & "Y" & y & "M" & mon & "D" & d & "H" & h & "M" & min & "S" & s & "P" & intPIID & ".jpg"
                PicPath = strPicPath & "\" & "C" & bytCompID & "Y" & y & "M" & mon & "D" & d & "H" & h & "M" & min & "S" & s & "P" & intPIID & ".jpg"
        
        
        
        
                MainComm.CommandText = "INSERT INTO NEW_PI (PIID,comID,GroupName,PI_No,PI_Year,PINo,UserID,PCName,EntryDate,EntryTime)" _
                                     & "VALUES (" & intPIID & "," & bytCompID & ",'" & cboItType.Text & "','" & fncMRR(intPI) & "','" & intYear & "','" & txtPINo.Text & "'," & bytUserID & ",'" & getComName & "','" & Date & "','" & Time & "');"
                MainComm.Execute



                 MainComm.CommandText = "INSERT INTO [ThreadInfo]([CompID],[ComName],[ComAdd],[SuppID],SuppName,[Suppadd],[POID],[PONO],[PODATE],[ItemGr],[Path],[BuyerRef],Remarks,USERID,PCNAME,ENTRYDATE,ENTRYTIME) Values( " & bytCompID & ",'" & strComNameTxt & "','" & ComADD & "'," & cboSupplier.Tag & ",'" & cboSupplier.Text & "','" & txtAddress.Text & "'," & intPIID & ",'" & txtPINo.Text & "','" & txtpodate.Text & "','" & cboItType.Text & "','" & PicPath & "','" & txtBuyer.Text & "','" & txtRemarks.Text & "','" & bytUserID & "','" & getComName & "','" & Date & "','" & Time & "');"
                 MainComm.Execute


'                 MainComm.CommandText = "INSERT INTO [ThreadInfo]([CompID],[ComName],[ComAdd],[SuppID],SuppName,[Suppadd],[POID],[PONO],[PODATE],[ItemGr],[BuyerRef],Remarks,USERID,PCNAME,ENTRYDATE,ENTRYTIME) Values( " & bytCompID & ",'" & strComNameTxt & "','" & ComADD & "'," & cboSupplier.Tag & ",'" & cboSupplier.Text & "','" & PicPath & "'," & intPIID & ",'" & txtPINo.Text & "','" & txtpodate.Text & "','" & cboItType.Text & "','" & txtBuyer.Text & "','" & txtRemarks.Text & "','" & bytUserID & "','" & getComName & "','" & Date & "','" & Time & "');"
'                 MainComm.Execute
                
'                 MainComm.CommandText = "Update ThreadInfo set ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & txtPINo.Text & "' AND DELETEROW=0 "
'                 MainComm.Execute
              
              
         MsgBox "Data SAVED Successfully ."
        
    Else
    
    
    
'              SavePicture Image1.Picture, strPicPath & "\Picture\" & "C_" & bytCompID & "_D_" & Date & "_P_" & intPIID & "_I_" & Image1 & ".jpg"
'              PicPath = strPicPath & "\Picture\" & "C_" & bytCompID & "_D_" & Date & "_P_" & intPIID & "_I_" & Image1 & ".jpg"


                SavePicture Image1.Picture, strSerPicPath & "\" & "C" & bytCompID & "Y" & y & "M" & mon & "D" & d & "H" & h & "M" & min & "S" & s & "P" & intPIID & ".jpg"

                SavePicture Image1.Picture, strPicPath & "\" & "C" & bytCompID & "Y" & y & "M" & mon & "D" & d & "H" & h & "M" & min & "S" & s & "P" & intPIID & ".jpg"
                PicPath = strPicPath & "\" & "C" & bytCompID & "Y" & y & "M" & mon & "D" & d & "H" & h & "M" & min & "S" & s & "P" & intPIID & ".jpg"


    
                 MainComm.CommandText = "UPDATE [ThreadInfo] SET USERID='" & bytUserID & "',PCNAME='" & getComName & "',ENTRYDATE='" & Date & "',ENTRYTIME='" & Time & "' ,DELETEROW=2 WHERE COMPID=" & bytCompID & "  AND SUPPID=" & cboSupplier.Tag & " AND POID=" & intPIID & " AND DELETEROW=0 "
                 MainComm.Execute
                 
                 
                 
                  MainComm.CommandText = "INSERT INTO [ThreadInfo]([CompID],[ComName],[ComAdd],[SuppID],SuppName,[Suppadd],[POID],[PONO],[PODATE],[ItemGr],[Path],[BuyerRef],Remarks,USERID,PCNAME,ENTRYDATE,ENTRYTIME) Values( " & bytCompID & ",'" & strComNameTxt & "','" & ComADD & "'," & cboSupplier.Tag & ",'" & cboSupplier.Text & "','" & txtAddress.Text & "'," & intPIID & ",'" & strPI & "','" & txtpodate.Text & "','" & cboItType.Text & "','" & PicPath & "','" & txtBuyer.Text & "','" & txtRemarks.Text & "','" & bytUserID & "','" & getComName & "','" & Date & "','" & Time & "');"
                  MainComm.Execute
                
'                  MainComm.CommandText = "Update ThreadInfo set ThreadInfo.Spphoto=dbo.spPhoto.SpPhoto FROM ThreadInfo CROSS JOIN dbo.spPhoto where PONO='" & cboeditPO.Text & "' AND DELETEROW=0 "
'                  MainComm.Execute
                 
                 
 MsgBox "Data UPDATE Successfully ."

    End If
   
   
   ''--------------
    
    
    
     MainConn.CommitTrans
     
        

           
         
            Image1.Picture = Nothing
            Select_Image = False
            
'            cboSupplier.Text = ""
            txtAddress.Text = ""
            txtPINo.Text = ""
            txtpodate.Text = ""
'            cboItType.Text = ""
            txtBuyer.Text = ""
            txtRemarks.Text = ""
            

  
Screen.MousePointer = vbDefault
Set R = Nothing

'addSupplier cboSupplier

ErrorHandler:
    If err.Number <> 0 Then MsgBox err.Description, vbCritical, cnstMsgErDB

End Sub

Private Sub cmdZoomEqual_Click()
Image1.Width = ZoomSizeOrgWid
  Image1.Left = 0
  Image1.Top = 0
  Image1.Height = ZoomSizeOrgHgh

End Sub

Private Sub cmdZoomMinus_Click()
Image1.Stretch = True

If Image1.Width < 500 Then
Exit Sub
End If

If Image1.Height < 500 Then
Exit Sub
End If
 
Image1.Width = Image1.Width - (Image1.Width * 0.1)
Image1.Left = 0
Image1.Top = 0
Image1.Height = Image1.Height - (Image1.Height * 0.1)
End Sub

Private Sub CmdZoomPlus_Click()
Image1.Stretch = True
 
Image1.Width = Image1.Width + (Image1.Width * 0.1)
Image1.Left = 0
Image1.Top = 0
Image1.Height = Image1.Height + (Image1.Height * 0.1)


End Sub

Private Sub Form_Load()

 ' addSupplier cboSupplier
  addSupplierName cboSupplier
   addItemType cboItType

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmThredPO = Nothing
End Sub

Private Sub txtAddress_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
 SingleCodeFunc KeyAscii
End Sub

Private Sub txtAddress_LostFocus()
 txtAddress.Text = Trim(txtAddress.Text)
End Sub

Private Sub txtBuyer_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub

Private Sub txtpodate_KeyDown(KeyCode As Integer, Shift As Integer)
   Call prcMoveTab(KeyCode)
End Sub

Private Sub txtpodate_LostFocus()
 txtpodate.Text = FormatDate(txtpodate)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
Call prcMoveTab(KeyCode)
End Sub
