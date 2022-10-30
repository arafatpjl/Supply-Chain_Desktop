VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4290
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   7590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00B37102&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4394
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      Begin VB.Label lblB3OP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BASHUNDHARA 3/O, 3/P PROJECT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   26
         MouseIcon       =   "frmSplash.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Tag             =   "BASHUNDHARA R/A - PLOT # 3/O, 3/P PROJECT"
         Top             =   2522
         Width           =   2860
      End
      Begin VB.Label lblB5MN 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BASHUNDHARA 5/M, 5/N PROJECT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   195
         MouseIcon       =   "frmSplash.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Tag             =   "BASHUNDHARA R/A - PLOT # 5/M, 5/N PROJECT"
         Top             =   2040
         Width           =   2860
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(HU)"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   247
         Left            =   7072
         TabIndex        =   19
         Top             =   2522
         Width           =   377
      End
      Begin VB.Label lblCRP 
         BackStyle       =   0  'Transparent
         Caption         =   "28 R.C. CHURCH ROAD RESIDENTIAL PROJECT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   3068
         MouseIcon       =   "frmSplash.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   3406
         Width           =   3952
      End
      Begin VB.Label lblPCL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC CASUALS LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   572
         MouseIcon       =   "frmSplash.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   3406
         Width           =   1768
      End
      Begin VB.Label lblPWP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAKKA RASTER MATHA WAREHOUSE PROJECT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   3240
         MouseIcon       =   "frmSplash.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2040
         Width           =   4050
      End
      Begin VB.Label lblHUAL 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIVERSAL JEANS LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   5239
         MouseIcon       =   "frmSplash.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   2522
         Width           =   1820
      End
      Begin VB.Label lblNZND 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIAMOND FASHION LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   2340
         MouseIcon       =   "frmSplash.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   3965
         Visible         =   0   'False
         Width           =   2444
      End
      Begin VB.Label lblNZNG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NZN GARMENTS LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   130
         MouseIcon       =   "frmSplash.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   3952
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label lblCdaAgrabad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC JEANS LTD (CDA AGRABAD)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   600
         MouseIcon       =   "frmSplash.frx":1850
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   3000
         Width           =   3060
      End
      Begin VB.Label lblFisharighat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC JEANS LTD (FISHARIGHAT)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   225
         Left            =   3960
         MouseIcon       =   "frmSplash.frx":1B5A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   3000
         Width           =   2955
      End
      Begin VB.Label lblNZNF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NZN FASHION WEARS LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   2990
         MouseIcon       =   "frmSplash.frx":1E64
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2522
         Width           =   2080
      End
      Begin VB.Label lblNHTHOLD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NHT HOLDINGS LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   5161
         MouseIcon       =   "frmSplash.frx":216E
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1560
         Width           =   1573
      End
      Begin VB.Label lblNHT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NHT FASHIONS LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   3120
         MouseIcon       =   "frmSplash.frx":2478
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1560
         Width           =   1547
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLY CHAIN MANAGEMENT"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   435
         Left            =   1200
         TabIndex        =   7
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label lblALL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALL COMPANY"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   4836
         MouseIcon       =   "frmSplash.frx":2782
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   3952
         Visible         =   0   'False
         Width           =   1274
      End
      Begin VB.Label lblNZN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NZN GROUP"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   6266
         MouseIcon       =   "frmSplash.frx":2A8C
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   3965
         Visible         =   0   'False
         Width           =   1131
      End
      Begin VB.Label lblPAL 
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC ACCESSORIES LTD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   364
         MouseIcon       =   "frmSplash.frx":2D96
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1521
         Width           =   2184
      End
      Begin VB.Label lblUJL 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIVERSAL JEANS LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   416
         MouseIcon       =   "frmSplash.frx":30A0
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1040
         Width           =   2314
      End
      Begin VB.Label lblJ2L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JEANS 2000 LTD."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   5304
         MouseIcon       =   "frmSplash.frx":33AA
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1040
         Width           =   1196
      End
      Begin VB.Label lblPJL 
         BackStyle       =   0  'Transparent
         Caption         =   "PACIFIC JEANS LIMITED"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E26101&
         Height          =   221
         Left            =   2938
         MouseIcon       =   "frmSplash.frx":36B4
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1040
         Width           =   1989
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E26101&
         BorderWidth     =   3
         X1              =   0
         X2              =   8520
         Y1              =   3822
         Y2              =   3822
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E26101&
         BorderWidth     =   3
         X1              =   -870
         X2              =   7650
         Y1              =   675
         Y2              =   675
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

Call ConnectDatabase

Dim logID As Integer

Set R1 = New ADODB.Recordset
strRecord = "SELECT ComId,UJL,J2L,PAL,NHT,NHTHOLD,CDA,FG,NZNF,NZNG,NZND,HUAL,PWP,PCL,CRP,B5MN,B3OP FROM InventoryPJL.dbo.Sys_Configure_Status WHERE IDNo = '" & getComVolID & "';"
R1.Open strRecord, MainConn, adOpenStatic
If R1.RecordCount > 0 Then
  For i = 1 To 17
    If i = 1 Then
        logID = R1!ComID
      ElseIf i = 2 Then logID = R1!UJL
      ElseIf i = 3 Then logID = R1!J2L
      ElseIf i = 4 Then logID = R1!PAL
      ElseIf i = 5 Then logID = R1!NHT
      ElseIf i = 6 Then logID = R1!NHTHOLD
      ElseIf i = 7 Then logID = R1!CDA
      ElseIf i = 8 Then logID = R1!FG
      ElseIf i = 9 Then logID = R1!NZNF
      ElseIf i = 10 Then logID = R1!NZNG
      ElseIf i = 11 Then logID = R1!NZND
      ElseIf i = 12 Then logID = R1!HUAL
      ElseIf i = 13 Then logID = R1!PWP
      ElseIf i = 14 Then logID = R1!PCL
      ElseIf i = 15 Then logID = R1!CRP
      ElseIf i = 16 Then logID = R1!B5MN
      ElseIf i = 17 Then logID = R1!B3OP
      
    End If
    
    If logID = 1 Then
        lblPJL.Enabled = True
      ElseIf logID = 3 Then lblPAL.Enabled = True
      ElseIf logID = 4 Then lblNZN.Enabled = True
      ElseIf logID = 5 Then lblJ2L.Enabled = True
      ElseIf logID = 7 Then lblUJL.Enabled = True
      ElseIf logID = 8 Then lblALL.Enabled = True
      ElseIf logID = 13 Then lblNHT.Enabled = True
      ElseIf logID = 14 Then lblNHTHOLD.Enabled = True
      ElseIf logID = 15 Then lblCdaAgrabad.Enabled = True
      ElseIf logID = 16 Then lblFisharighat.Enabled = True
      ElseIf logID = 17 Then lblNZNF.Enabled = True
      ElseIf logID = 18 Then lblNZNG.Enabled = True
      ElseIf logID = 19 Then lblNZND.Enabled = True
      ElseIf logID = 20 Then lblHUAL.Enabled = True
      ElseIf logID = 21 Then lblPWP.Enabled = True
      ElseIf logID = 22 Then lblPCL.Enabled = True
      ElseIf logID = 23 Then lblCRP.Enabled = True
      ElseIf logID = 24 Then lblB5MN.Enabled = True
      ElseIf logID = 25 Then lblB3OP.Enabled = True
    End If
  Next i

Else
  MsgBox "Please Contact Software Vendor..", vbCritical, "ERROR"
End If

R1.Close
Set R1 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmSplash = Nothing
End Sub

Private Sub Frame1_DblClick()
  Unload Me
End Sub

Private Sub lblALL_Click()
  Frame1.Tag = lblALL.Caption
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblB3OP_Click()
  Frame1.Tag = lblB3OP.Tag
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  ''Load frmUserLogin
  ''frmUserLogin.Refresh
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblB5MN_Click()
  Frame1.Tag = lblB5MN.Tag
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  ''Load frmUserLogin
  ''frmUserLogin.Refresh
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblCdaAgrabad_Click()
  Frame1.Tag = lblCdaAgrabad.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblCRP_Click()
  Frame1.Tag = lblCRP.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblFisharighat_Click()
  Frame1.Tag = lblFisharighat.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblHUAL_Click()
  Frame1.Tag = lblHUAL.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblJ2L_Click()
  Frame1.Tag = lblJ2L.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblNHT_Click()
  Frame1.Tag = lblNHT.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblNHTHOLD_Click()
  Frame1.Tag = lblNHTHOLD.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblNZN_Click()
  Frame1.Tag = lblNZN.Caption
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblNZND_Click()
  Frame1.Tag = lblNZND.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblNZNF_Click()
  Frame1.Tag = lblNZNF.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblNZNG_Click()
  Frame1.Tag = lblNZNG.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblPAL_Click()
  Frame1.Tag = lblPAL.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblPCL_Click()
  Frame1.Tag = lblPCL.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  ''Load frmUserLogin
  ''frmUserLogin.Refresh
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblPJL_Click()
  Frame1.Tag = lblPJL.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  ''Load frmUserLogin
  ''frmUserLogin.Refresh
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblPWP_Click()
  Frame1.Tag = lblPWP.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Private Sub lblUJL_Click()
  Frame1.Tag = lblUJL.Caption
  strComNameTxt = Frame1.Tag
  Call PrcCompanyLogIn
  frmUserLogin.Show
  Unload frmSplash
End Sub

Public Function PrcCompanyLogIn()

On Error GoTo ErrorX
Screen.MousePointer = vbHourglass

Set R = New ADODB.Recordset
strRecord = "SELECT * FROM Company_Information WHERE CompName = '" & Frame1.Tag & "';"
R.Open strRecord, MainConn, adOpenStatic

WSName = R("ServerName").Value
WDBName = R("DBName").Value
WDBPrefix = R("DBPrefix").Value
ReportPath = R("ReportPath").Value
bytCompID = R("CompID").Value
strRptPath = R("ReportPath").Value
isPrimaryDB = False

Call ConnectDatabase

' frmUserLogin.Show
' Unload frmSplash

ErrorX:
If err.Number <> 0 Then
  MsgBox err.Description, vbCritical, cnstMsgErDB
  Screen.MousePointer = vbDefault
  Exit Function
End If

R.Close
Set R = Nothing

End Function
