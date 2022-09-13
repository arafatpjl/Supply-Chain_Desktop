VERSION 5.00
Begin VB.Form frmNewAccountsHead 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " New Accounts Head"
   ClientHeight    =   1230
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   5160
   Icon            =   "frmNewAccountsHead.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   3585
      Begin VB.TextBox txtAccountsHead 
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
         Left            =   150
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "New Accounts Head :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C000&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   2
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C000&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   1
      ToolTipText     =   "Create New Company"
      Top             =   180
      Width           =   1065
   End
End
Attribute VB_Name = "frmNewAccountsHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
Dim NewID As Integer
Dim dtCurrentDate As Date
    dtCurrentDate = Date
    On Error GoTo x
    
    
    
    
    If Len(txtAccountsHead) = 0 Then
        If strMark = "HEAD" Then
            MsgBox "Enter name of Accounts Head", vbInformation, AppTitle
        End If
        txtAccountsHead.SetFocus
        Exit Sub
    End If
    
    
    If frmNewAccountsHead.Caption = " New Accounts Head" Then
     
     strFormName = "Acc Head For BillDetails"
     strMark = "HEAD"
     flagMark = "INSERT"
     strValue = ""
     
     ElseIf frmNewAccountsHead.Caption = "Edit Accounts Head" Then
     
     
        strFormName = "Acc Head For BillDetails"
        strMark = "HEAD"
        flagMark = "EDIT"
        
        
'    With frmNewAccountsHead
'
'        flexDetails.Col = 0:    .Tag = flexDetails.Text
'
'        flexDetails.Col = 1:
'
'            .txtAccountsHead.Text = flexDetails.Text
'            .txtAccountsHead.Tag = flexDetails.Text
'
'        .Caption = "Edit Accounts Head"
'        .Show 1
'    End With
     
     
     
     End If
     
     
    
Set MainComm = New ADODB.Command
 MainComm.ActiveConnection = MainConn

    If fncDuplicate = True Then Exit Sub
    MainConn.BeginTrans
        If flagMark = "INSERT" Then
            If strFormName = "Acc Head For BillDetails" Then
                NewID = fncNewID("HEAD")
                MainComm.CommandText = "INSERT INTO InventoryPJL.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryUJL.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryJ2L.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryPAL.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryNHT.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryNHTHOLD.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryPJLCA.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryPJLFG.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryNZNF.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryNZNG.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                MainComm.CommandText = "INSERT INTO InventoryHUAL.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                 MainComm.CommandText = "INSERT INTO InventoryPWP.DBO.AccountsHead( HeadID, HeadName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
                MainComm.Execute
                
                
            ElseIf strFormName = "Acc Head For C_And_F" Then
                NewID = fncNewID("HEAD_C_and_F")
                MainComm.CommandText = "INSERT INTO C_And_F_new_Acc_Head( AccID, AccName,CreateDate,UserName,PCName) VALUES(" & NewID & ", '" & txtAccountsHead.Text & "', '" & dtCurrentDate & "','" & strUser & "','" & getComName & "');"
             MainComm.Execute
            End If
'             MainComm.Execute
        Else
            If strFormName = "Acc Head For BillDetails" Then
                 MainComm.CommandText = "UPDATE InventoryPJL.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryUJL.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryJ2L.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryPAL.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryNHT.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryNHTHOLD.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryPJLCA.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryPJLFG.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryNZNF.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryNZNG.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 MainComm.CommandText = "UPDATE InventoryHUAL.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                  MainComm.CommandText = "UPDATE InventoryPWP.DBO.AccountsHead SET HeadName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE HeadID=" & Tag & ";"
                 MainComm.Execute
                 
            ElseIf strFormName = "Acc Head For C_And_F" Then
               MainComm.CommandText = "UPDATE C_And_F_new_Acc_Head SET AccName='" & txtAccountsHead.Text & "',CreateDate = '" & dtCurrentDate & "', UserName = '" & strUser & "', PCName = '" & getComName & "' WHERE AccID=" & Tag & ";"
              MainComm.Execute
            End If
            
        End If
    MainConn.CommitTrans
    strValue = txtAccountsHead.Text
    Unload Me
Exit Sub
x:
    MsgBox err.Description, vbCritical, cnstMsgErDB
    MainConn.RollbackTrans
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNewAccountsHead = Nothing
End Sub
Private Sub txtAccountsHead_GotFocus()
    Call txtGotFocus(txtAccountsHead)
End Sub
Private Sub txtAccountsHead_KeyDown(KeyCode As Integer, Shift As Integer)
    Call prcMoveTab(KeyCode)
End Sub
Private Sub txtAccountsHead_KeyPress(KeyAscii As Integer)
    SingleCodeFunc KeyAscii
End Sub
Private Sub txtAccountsHead_LostFocus()
    txtAccountsHead.Text = UCase(Trim(txtAccountsHead.Text))
End Sub

Private Function fncDuplicate() As Boolean
    fncDuplicate = True
    If flagMark = "INSERT" Or flagMark = "EDIT" Or (Trim(txtAccountsHead.Text) = Trim(txtAccountsHead.Tag)) Then
        If strFormName = "Acc Head For BillDetails" Then
            Mysql = "SELECT Count(*) FROM AccountsHead WHERE (HeadName='" & txtAccountsHead.Text & "');"
        ElseIf strFormName = "Acc Head For C_And_F" Then
            Mysql = "SELECT Count(*) FROM C_And_F_new_Acc_Head WHERE (AccName='" & txtAccountsHead.Text & "');"
        End If
 
        com.CommandText = Mysql
        Set RS = com.Execute
        If RS.Fields(0) > 0 Then
            MsgBox "Accounts Head '" & txtAccountsHead.Text & "' already exist." & vbCrLf & "Change the Accounts Head name.", vbInformation, AppTitle
            txtAccountsHead.SetFocus
            Exit Function
        End If
    End If
    fncDuplicate = False
End Function
