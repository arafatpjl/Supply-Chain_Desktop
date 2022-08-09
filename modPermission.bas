Attribute VB_Name = "modPermission"
Public Sub prcAddUserName(cmb As ComboBox)
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn

    Mysql = "Select UserId, UserName from InventoryPJL.dbo.Sys_User_Name;"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    
    cmb.Clear
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
'            cmb.AddItem DecryptIt(RS.Fields(1).Value, 11)
            cmb.AddItem RS.Fields(1).Value
            cmb.ItemData(cmb.NewIndex) = RS.Fields(0).Value
        RS.MoveNext
        Loop
    End If
    
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Public Sub prcAddMM(cmb As ComboBox, Userid As Integer)
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    Set R = New Recordset

    Mysql = "Select Menu_Main_Id, Menu_Main_Caption from InventoryPJL.dbo.viewUserMM Where UserId = " & Userid & " Order By Menu_Main_Id"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    
    cmb.Clear
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            cmb.AddItem RS.Fields(1).Value
            cmb.ItemData(cmb.NewIndex) = RS.Fields(0).Value
        RS.MoveNext
        Loop
    End If
    
    Set R = Nothing
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Public Sub prcAddMS(cmb As ComboBox, Userid As Integer, MMID As Integer)
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    Set R = New Recordset

    Mysql = "Select Menu_Sub_Id, Menu_Sub_Caption from InventoryPJL.dbo.viewUserMS Where UserId = " & Userid & " and menu_main_id = " & MMID & " Order By Menu_Sub_Id"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    
    cmb.Clear
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            cmb.AddItem RS.Fields(1).Value
            cmb.ItemData(cmb.NewIndex) = RS.Fields(0).Value
        RS.MoveNext
        Loop
    End If
    
    Set R = Nothing
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Public Sub prcFillListMM(lst As ListBox)
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    Set R = New Recordset
    
    Mysql = "Select Menu_Main_Id, Menu_Main_Caption from InventoryPJL.dbo.Menu_Main Order By Menu_Main_Id;"
    MainComm.CommandText = Mysql
    Set R = MainComm.Execute
    
    lst.Clear
    If R.EOF = False And R.BOF = False Then
        Do While Not R.EOF
            lst.AddItem R.Fields(1).Value
            lst.ItemData(lst.NewIndex) = R.Fields(0).Value
        R.MoveNext
        Loop
    End If
    Set R = Nothing
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Public Sub prcCheckPermMM(lst As ListBox, Userid As Integer)
    Dim i As Integer
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next
    
    Mysql = "Select Menu_Main_Id, Menu_Main_Caption from InventoryPJL.dbo.viewUserMM Where UserID = " & Userid & ";"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            lst.Tag = 0
            
            For i = 0 To lst.ListCount - 1
                lst.Tag = lst.ItemData(i)
                If RS("Menu_Main_Id") = lst.Tag Then lst.Selected(i) = True
            Next
        RS.MoveNext
        Loop
    End If
    Set MainComm = Nothing
End Sub

Public Sub prcCheckPermMS(lst As ListBox, Userid As Integer, MMID As Integer)
    Dim i As Integer
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next
    
    Mysql = "Select Menu_Sub_Id, Menu_Sub_Caption from InventoryPJL.dbo.viewUserMS Where UserID = " & Userid & " And Menu_Main_ID = " & MMID & ";"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            lst.Tag = 0
            
            For i = 0 To lst.ListCount - 1
                lst.Tag = lst.ItemData(i)
                If RS("Menu_Sub_Id") = lst.Tag Then lst.Selected(i) = True
            Next
        RS.MoveNext
        Loop
    End If
    Set MainComm = Nothing
End Sub

Public Sub prcCheckPermMSS(lst As ListBox, Userid As Integer, MMID As Integer, MSID As Integer)
    Dim i As Integer
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next
    
    Mysql = "Select Menu_Sub_Sub_Id, Menu_Sub_Sub_Caption from InventoryPJL.dbo.viewUserMSS Where UserID = " & Userid & " And Menu_Main_ID = " & MMID & "  And Menu_Sub_ID = " & MSID & ";"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute
    If RS.EOF = False And RS.BOF = False Then
        Do While Not RS.EOF
            lst.Tag = 0
            
            For i = 0 To lst.ListCount - 1
                lst.Tag = lst.ItemData(i)
                If RS("Menu_Sub_Sub_Id") = lst.Tag Then lst.Selected(i) = True
            Next
        RS.MoveNext
        Loop
    End If
    Set MainComm = Nothing
End Sub

Public Sub prcFillListMS(lst As ListBox, MMID As Integer)
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    Set R = New Recordset
    
    Mysql = "Select Menu_Sub_Id, Menu_Sub_Caption from InventoryPJL.dbo.Menu_Sub Where Menu_Main_Id = " & MMID & " Order By Menu_Sub_Id;"
    MainComm.CommandText = Mysql
    Set R = MainComm.Execute
    
    lst.Clear
    If R.EOF = False And R.BOF = False Then
        Do While Not R.EOF
            lst.AddItem R.Fields(1).Value
            lst.ItemData(lst.NewIndex) = R.Fields(0).Value
        R.MoveNext
        Loop
    End If
    Set R = Nothing
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Public Sub prcFillListMSS(lst As ListBox, MMID As Integer, MSID As Integer)
On Error GoTo x
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    Set R = New Recordset
    
    Mysql = "Select Menu_Sub_Sub_Id, Menu_Sub_Sub_Caption from InventoryPJL.dbo.Menu_Sub_Sub Where Menu_Sub_Id = " & MSID & " Order By Menu_Sub_Sub_Id;"
    MainComm.CommandText = Mysql
    Set R = MainComm.Execute
    
    lst.Clear
    If R.EOF = False And R.BOF = False Then
        Do While Not R.EOF
            lst.AddItem R.Fields(1).Value
            lst.ItemData(lst.NewIndex) = R.Fields(0).Value
        R.MoveNext
        Loop
    End If
    Set R = Nothing
    Set MainComm = Nothing
Exit Sub
x:
MsgBox err.Description, vbCritical, AppTitle
End Sub

Public Sub prcMenuInvisibleAll()
   
    With frmMDIMain
       
        .mnuHouseKeeping.Visible = False                'Menu_Main_Id = 1
            .mnuItemGroup.Visible = True
            .mnuNewItemName.Visible = False                  'Menu_Sub_Id = 1
            .mnuNewSupplier.Visible = False                 'Menu_Sub_Id = 2
            .mnuNewPFNo.Visible = False                     'Menu_Sub_Id = 3
            .mnuacchead.Visible = False
        .mnuDataEntry.Visible = False                   'Menu_Main_Id = 2
            .mnuaccount.Visible = False
                .menuchkinfo.Visible = False
                .menuBillNOCash.Visible = False
                .mnuBillCr.Visible = True
            .mnuSupplyChain.Visible = True
                .mnuPurchaseOrder.Visible = True
                .mnuthread.Visible = False
                .mnuPurchaseOrderOC.Visible = False
                .mencashcr.Visible = False
                .menuchk.Visible = False
            .mnuNext01.Visible = False
            .mnuStore.Visible = False
                .mnumrrinvchallan.Visible = True                'Menu_Sub_Id = 5
                .mnuMRRInfoOthCur.Visible = False
        .mnuReport.Visible = False                      'Menu_Main_Id = 3
            .mnurptSupplyChain.Visible = False                       'Menu_Sub_Id = 6
                .mnurptmrrchallan.Visible = True
                .mnurptPOBalanceQty.Visible = False
                .mnuRptBillwise.Visible = False
                .mnuRptPartystatement.Visible = False
                .mnuMRRwisereport.Visible = False
                .mnuRptTransactionwisereport.Visible = False
                .mnuRptChallanwisereport.Visible = False
                
                .mnuMonthlystatement.Visible = False
                .mnuitemname.Visible = False
                
            .mnunext02.Visible = False
            .mnurptStore.Visible = False
                .mnurptpurchaseinvoice.Visible = True
                .mnurptPOBalanceQty.Visible = False
                
            .mnuNext03.Visible = True
            .mnurptaudit.Visible = False                        'Menu_Sub_Id = 10
                .mnurptmrrEditDelete.Visible = True                         'Menu_Sub_Sub_Id = 3
                .mnurptchallanEditDelete.Visible = False                     'Menu_Sub_Sub_Id = 4
                .MnuNext07.Visible = False
                .mnurptPOEditDelete.Visible = False
            .menurptAccount.Visible = False
                .mnurptcheque.Visible = False
                .mnurptbilcash.Visible = False
                .mnurptbilcr.Visible = True
                
                
        .mnuOption.Visible = False                              'Menu_Main_Id = 4
            .mnuModifyComp.Visible = False                   'Menu_Sub_Id = 11
            .mnuDeleteComp.Visible = False                  'Menu_Sub_Id = 12
            .mnuNext04.Visible = False                       'Menu_Sub_Id = 13
            .mnuCreateUser.Visible = False                  'Menu_Sub_Id = 14
            .mnuNext05.Visible = False                      'Menu_Sub_Id = 15
            .mnuSetupPermission.Visible = False             'Menu_Sub_Id = 16
            .mnuNext06.Visible = False                      'Menu_Sub_Id = 17
                .mnuUserCompany.Visible = True                     'Menu_Sub_Sub_Id = 5
                .mnuUserMenu.Visible = False                        'Menu_Sub_Sub_Id = 6
            .mnuoptLockData.Visible = True
                  
    End With
End Sub

Public Sub prcMenuVisibleCompany()
    Dim DHK As Integer, DDE As Integer
    DHK = 0: DDE = 0
    
    Set MainComm = New Command
    MainComm.ActiveConnection = MainConn
    Set RS = New Recordset

    Mysql = "Select Menu_Main_Name, Menu_Sub_Name, Menu_Sub_Id, UserId from InventoryPJL.dbo.viewUserMS Where UserId = " & bytUserID & " Order By Menu_Sub_Id;"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute

    With frmMDIMain
        If RS.EOF = False And RS.BOF = False Then
            Do While Not RS.EOF
                ''  Start - House Keeping
                If RS("Menu_Main_Name").Value = "mnuHouseKeeping" Then .mnuHouseKeeping.Visible = True                              'Menu_Main_Id = 1
                    If RS("Menu_Sub_Name").Value = "mnuItemGroup" Then                                            'Menu_Sub_Id = 1
                        .mnuItemGroup.Visible = True
                        DHK = 1
                    End If
                    If RS("Menu_Sub_Name").Value = "mnuNewItemName" Then .mnuNewItemName.Visible = True             'Menu_Sub_Id = 2
                    If RS("Menu_Sub_Name").Value = "mnuNewSupplier" Then .mnuNewSupplier.Visible = True             'Menu_Sub_Id = 2
                    If RS("Menu_Sub_Name").Value = "mnuNewPFNo" Then .mnuNewPFNo.Visible = True                     'Menu_Sub_Id = 3
                    If RS("Menu_Sub_Name").Value = "mnuacchead" Then .mnuacchead = True                      'Menu_Sub_Id = 3
                '' End  -   House Keeping
                
                '' Start - Data Entry
                If RS("Menu_Main_Name").Value = "mnuDataEntry" Then .mnuDataEntry.Visible = True                                    'Menu_Main_Id = 2
                    If RS("Menu_Sub_Name").Value = "mnuSupplyChain" Then
                        .mnuSupplyChain.Visible = True                                                            'Menu_Sub_Id = 4
                        DDE = 1
                    End If
                    If RS("Menu_Sub_Name").Value = "mnunext01" Then .mnuNext01.Visible = True          'Menu_Sub_Id = 5
                    If RS("Menu_Sub_Name").Value = "mnuStore" Then .mnuStore.Visible = True          'Menu_Sub_Id = 5
                    If RS("Menu_Sub_Name").Value = "mnuaccount" Then .mnuaccount.Visible = True          'Menu_Sub_Id = 5
                    
                ''End - Data Entry
                    
                ''Start - Reports Menu
                If RS("Menu_Main_Name").Value = "mnuReport" Then .mnuReport.Visible = True                                          'Menu_Main_Id = 3
                    If RS("Menu_Sub_Name").Value = "mnurptSupplyChain" Then .mnurptSupplyChain.Visible = True                            'Menu_Sub_Id = 6
                    If RS("Menu_Sub_Name").Value = "mnunext02" Then .mnunext02.Visible = True                          'Menu_Sub_Id = 7
                    If RS("Menu_Sub_Name").Value = "mnurptstore" Then .mnurptStore.Visible = True                            'Menu_Sub_Id = 8
                    If RS("Menu_Sub_Name").Value = "mnunext03" Then .mnuNext03.Visible = True                                  'Menu_Sub_Id = 9
                    If RS("Menu_Sub_Name").Value = "mnurptaudit" Then .mnurptaudit.Visible = True                                     'Menu_Sub_Id = 10
                    If RS("Menu_Sub_Name").Value = "menurptAccount" Then .menurptAccount.Visible = True                                     'Menu_Sub_Id = 10
                '' End - Report Menu
                 
                 ''Start - Option Menu
                 If RS("Menu_Main_Name").Value = "mnuOption" Then .mnuOption.Visible = True                                          'Menu_Main_Id = 4
                     If RS("Menu_Sub_Name").Value = "mnuModifyComp" Then .mnuModifyComp.Visible = True               'Menu_Sub_Id = 11
                     If RS("Menu_Sub_Name").Value = "mnuDeleteComp" Then .mnuDeleteComp.Visible = True               'Menu_Sub_Id = 12
                     If RS("Menu_Sub_Name").Value = "mnuNext04" Then .mnuNext04.Visible = True                         'Menu_Sub_Id = 13
                     If RS("Menu_Sub_Name").Value = "mnuCreateUser" Then .mnuCreateUser.Visible = True               'Menu_Sub_Id = 14
                     If RS("Menu_Sub_Name").Value = "mnuNext05" Then .mnuNext05.Visible = True                       'Menu_Sub_Id = 15
                     If RS("Menu_Sub_Name").Value = "mnuSetupPermission" Then .mnuSetupPermission.Visible = True                       'Menu_Sub_Id = 15
                     If RS("Menu_Sub_Name").Value = "mnuNext06" Then .mnuNext06.Visible = True
                     If RS("Menu_Sub_Name").Value = "mnuoptLockData" Then .mnuoptLockData.Visible = True
''                     If rs("Menu_Sub_Name").Value = "mnuUserPerSetup" Then .mnuUserPerSetup.Visible = True         'Menu_Sub_Id = 65
                ''End - Option Menu
            RS.MoveNext
            Loop
        End If
    End With
    
    Mysql = "Select Menu_Sub_Name, Menu_Sub_Sub_Name, Menu_Sub_Sub_Id, UserId from InventoryPJL.dbo.viewUserMSS Where UserId = " & bytUserID & " Order By Menu_Sub_Sub_Id;"
    MainComm.CommandText = Mysql
    Set RS = MainComm.Execute

    With frmMDIMain
        If RS.EOF = False And RS.BOF = False Then
            Do While Not RS.EOF
                ''Start - Date Entry Menu
                If RS("Menu_Sub_Name").Value = "mnuSupplyChain" Then
                    .mnuSupplyChain.Visible = True                      'Menu_Main_Id = 4
                End If
                
                    If RS("Menu_Sub_Sub_Name").Value = "mnuPurchaseOrder" Then
                            .mnuPurchaseOrder.Visible = True            'Menu_Sub_Sub_Id = 1
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuPurchaseOrderOC" Then
                            .mnuPurchaseOrderOC.Visible = True            'Menu_Sub_Sub_Id = 1
                    End If
                    
                    
                     If RS("Menu_Sub_Sub_Name").Value = "mnuthread" Then
                            .mnuthread.Visible = True               'Menu_Sub_Sub_Id = 1
                    End If
                    
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mencashcr" Then
                            .mencashcr.Visible = True            'Menu_Sub_Sub_Id = 1
                    End If
                    
                     If RS("Menu_Sub_Sub_Name").Value = "menuchk" Then
                            .menuchk.Visible = True            'Menu_Sub_Sub_Id = 1
                    End If
                    
                If RS("Menu_Sub_Name").Value = "mnuStore" Then
                    .mnuStore.Visible = True                      'Menu_Main_Id = 4
                End If
                
                    If RS("Menu_Sub_Sub_Name").Value = "mnuMrrinvChallan" Then
                            .mnumrrinvchallan.Visible = True            'Menu_Sub_Sub_Id = 1
                    End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnuMRRInfoOthCur" Then
                            .mnuMRRInfoOthCur.Visible = True            'Menu_Sub_Sub_Id = 1
                    End If
                    
            '' End - Date Entry Menu
            ''Start - Report Menu
                If RS("Menu_Sub_Name").Value = "mnurptSupplyChain" Then
                    .mnurptSupplyChain.Visible = True                      'Menu_Main_Id = 4
                End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptpurchaseinvoice" Then
                        .mnurptpurchaseinvoice.Visible = True        'Menu_Sub_Sub_Id = 2
                    End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptPOBalanceQty" Then
                        .mnurptPOBalanceQty.Visible = True         'Menu_Sub_Sub_Id = 11
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuRptBillwise" Then
                        .mnuRptBillwise.Visible = True       'Menu_Sub_Sub_Id = 11
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuRptPartystatement" Then
                        .mnuRptPartystatement.Visible = True    'Menu_Sub_Sub_Id = 11
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuMRRwisereport" Then
                        .mnuMRRwisereport.Visible = True
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuRptTransactionwisereport" Then
                        .mnuRptTransactionwisereport.Visible = True
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuRptChallanwisereport" Then
                        .mnuRptChallanwisereport.Visible = True
                    End If
                 ''''''
                     If RS("Menu_Sub_Sub_Name").Value = "mnuitemname" Then
                        .mnuitemname.Visible = True
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuMonthlystatement" Then
                        .mnuMonthlystatement.Visible = True
                    End If
                    
                If RS("Menu_Sub_Name").Value = "mnurptStore" Then
                    .mnurptStore.Visible = True             'Menu_Sub_Id = 8
                End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptmrrchallan" Then
                        .mnurptmrrchallan.Visible = True            'Menu_Sub_Sub_Id = 7
                    End If
                    
                If RS("Menu_Sub_Name").Value = "mnuaudit" Then
                    .mnurptaudit.Visible = True                'Menu_Sub_Id = 10
                End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptmrrEditDelete" Then
                        .mnurptmrrEditDelete.Visible = True            'Menu_Sub_Sub_Id = 3
                    End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptchallanEditDelete" Then
                        .mnurptchallanEditDelete.Visible = True        'Menu_Sub_Sub_Id = 4
                    End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnuNext07" Then
                        .MnuNext07.Visible = True        'Menu_Sub_Sub_Id = 9
                    End If
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptPOEditDelete" Then
                        .mnurptPOEditDelete.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    If RS("Menu_Sub_Sub_Name").Value = "menuchkinfo" Then
                        .menuchkinfo.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "menuBillNOCash" Then
                        .menuBillNOCash.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnuBillCr" Then
                        .mnuBillCr.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptcheque" Then
                        .mnurptcheque.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptbilcash" Then
                        .mnurptbilcash.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    
                    If RS("Menu_Sub_Sub_Name").Value = "mnurptbilcr" Then
                        .mnurptbilcr.Visible = True        'Menu_Sub_Sub_Id = 10
                    End If
                    
                 '' End - Report Menu
                 
                 ''Start - Option Menu
                     If RS("Menu_Sub_Name").Value = "mnuSetupPermission" Then
                         .mnuSetupPermission.Visible = True                'Menu_Sub_Id = 65
                     End If
                     If RS("Menu_Sub_Sub_Name").Value = "mnuUserCompany" Then
                         .mnuUserCompany.Visible = True                 'Menu_Sub_Sub_Id = 12
                     End If
                     If RS("Menu_Sub_Sub_Name").Value = "mnuUserMenu" Then
                         .mnuUserMenu.Visible = True                    'Menu_Sub_Sub_Id = 13
                     End If
                ''End - Option Menu
            RS.MoveNext
            Loop
        End If
        
        If ((DHK = 0) And (.mnuHouseKeeping.Visible)) Then .mnuItemGroup.Visible = False
        If ((DDE = 0) And (.mnuDataEntry.Visible)) Then .mnuSupplyChain.Visible = False
    End With
    Set RS = Nothing
    Set MainComm = Nothing
End Sub

