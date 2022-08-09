Attribute VB_Name = "ModBill"
'strReportPath = App.Path & "\Reports\"

Public Function fncTotalDay(MNO As Integer) As Integer
    If MNO = 1 Or MNO = 3 Or MNO = 5 Or MNO = 7 Or MNO = 8 Or MNO = 10 Or MNO = 12 Then
        fncTotalDay = 31
    ElseIf MNO = 4 Or MNO = 6 Or MNO = 9 Or MNO = 11 Then
        fncTotalDay = 30
    ElseIf MNO = 2 Then
        fncTotalDay = 28
    End If
End Function

Function fncDateFunc(num As Integer) As String
    Dim strValid As String
    strValid = "0123456789"
    If num > 26 Then
        If InStr(strValid, Chr(num)) = 0 Then num = 0
    End If
End Function

Public Sub prcGridSetting(ByRef frm As Form, ByRef Grid As MSFlexGrid)
    Dim i As Integer, J As Integer
    With Grid
        If .Rows > 0 Then
            For i = 0 To .Rows - 1
                .Row = i
                If i Mod 2 = 0 Then
                    For J = 0 To .Cols - 1
                        .Col = J:   .CellBackColor = &HF0F1F2
                    Next
                Else
                    For J = 0 To .Cols - 1
                        .Col = J:   .CellBackColor = &H80000005
                    Next
                End If
                If .Row < (.Rows - 1) Then .Row = .Row + 1
            Next i
        End If
        If .Rows = 0 Then
            frm.cmdEdit.Enabled = False
            frm.cmdDelete.Enabled = False
        Else
            frm.cmdEdit.Enabled = True
            frm.cmdDelete.Enabled = True
        End If
    End With
End Sub

Public Function fncNewID2(ByVal Param As String)
On Error GoTo x
    If UCase(Param) = "COMPANY" Then
        Mysql = "Select  Max(CompanyID) AS NewID From Company_Information;"
    ElseIf UCase(Param) = "USER" Then
        Mysql = "Select  Max(UserID) AS NewID From User_Name_And_Password;"
    ElseIf UCase(Param) = "BUYER" Then
        Mysql = "Select  Max(BuyerID) AS NewID From Buyer_Information;"
    ElseIf UCase(Param) = "BUYERCATEGORY" Then
        Mysql = "Select  Max(BuyerCategoryID) AS NewID From Buyer_Information_Category;"
    ElseIf UCase(Param) = "PF" Then
        Mysql = "Select  Max(PFID) AS NewID From PF_Information;"
    ElseIf UCase(Param) = "STYLE" Then
        Mysql = "Select  Max(StyleID) AS NewID From Style_Information;"
    ElseIf UCase(Param) = "SIZE" Then
        Mysql = "Select  Max(SizeID) AS NewID From Size_Information_Main;"
    ElseIf UCase(Param) = "SUBSIZE" Then
        Mysql = "Select  Max(SubSizeID) AS NewID From Size_Information_Sub;"
    ElseIf UCase(Param) = "COLOR" Then
        Mysql = "Select  Max(ColorID) AS NewID From Color_Information;"
    ElseIf UCase(Param) = "INSEAM" Then
        Mysql = "Select  Max(InSeamID) AS NewID From In_Seam_Information;"
    ElseIf UCase(Param) = "ORDERNO" Then
        Mysql = "Select  Max(OrderID) AS NewID From NewOrder;"
    ElseIf UCase(Param) = "FACTORY" Then
        Mysql = "Select  Max(FactoryID) AS NewID From Factory_Information;"
    ElseIf UCase(Param) = "LINE" Then
        Mysql = "Select  Max(LineID) AS NewID From Line_Information;"
    ElseIf UCase(Param) = "SEASON" Then
        Mysql = "Select  Max(SeasonID) AS NewID From Season_Information;"
    ElseIf UCase(Param) = "SAMPLE" Then
        Mysql = "Select  Max(SampleID) AS NewID From Sample_Information;"
    ElseIf UCase(Param) = "SUPPLIER" Then
        Mysql = "Select  Max(SupplierID) AS NewID From Supplier_Information;"
    ElseIf UCase(Param) = "PARTY" Then
        Mysql = "Select  Max(PartyID) AS NewID From Party;"
    ElseIf UCase(Param) = "WASH" Then
        Mysql = "Select  Max(WashID) AS NewID From Wash_Information;"
    ElseIf UCase(Param) = "TRIMS" Then
        Mysql = "Select  Max(TrimsID) AS NewID From Trims_Information;"
    ElseIf UCase(Param) = "VEHICLES" Then
        Mysql = "Select  Max(SLNO) AS NewID From ParticularsOfVehicles;"
    ElseIf UCase(Param) = "COUNTRY" Then
        Mysql = "Select  Max(CountryID) AS NewID From Country_Information;"
    ElseIf UCase(Param) = "BANK" Then
        Mysql = "Select  Max(BankID) AS NewID From New_Bank;"
    ElseIf UCase(Param) = "HEAD" Then
        Mysql = "Select  Max(HeadID) AS NewID From AccountsHead;"
    ElseIf UCase(Param) = "HEAD_C_AND_F" Then
        Mysql = "Select  Max(AccID) AS NewID From C_And_F_new_Acc_Head;"
    ElseIf UCase(Param) = "EXPENSE" Then
        Mysql = "Select  Max(ExpenseID) AS NewID From C_And_F_new_Expense;"
    ElseIf UCase(Param) = "PARTY C AND F" Then
        Mysql = "Select  Max(PartyID) AS NewID From C_And_F_new_Party;"
    ElseIf UCase(Param) = "FUND" Then
        Mysql = "Select  Max(RefID) AS NewID From C_And_F_Indent_Fund_Main;"
    ElseIf UCase(Param) = "CFOUT" Then
        Mysql = "Select  Max(M_ID) AS NewID From C_And_F_Out_Main;"
    End If
    com.CommandText = Mysql
    Set RS = com.Execute
    
    If IsNull(RS(0).Value) Then
        fncNewID2 = 1
    Else
        fncNewID2 = RS(0).Value + 1
    End If
Exit Function
x:
MsgBox err.Description, vbCritical, AppTitle
End Function


Public Sub prcLoadCompany(cmb As ComboBox)
Dim strUserCategory As String
'    strUserCategory = frmMaster.StatusBar1.Panels(5).Text
    Call prcSetConnectionActive
    cmb.Clear
'    If strUserCategory = "<a>" Then
        Mysql = "Select * From Company_Information WHERE COMPID=" & bytCompID & ";"
'        com.CommandText = Mysql
'        Set RS = com.Execute
        RS.Open Mysql, MainConn, adOpenStatic
        While Not RS.EOF
            cmb.AddItem RS.Fields(1)
            cmb.ItemData(cmb.NewIndex) = RS.Fields(0)
            RS.MoveNext
        Wend
''''''    Else
''''''        Mysql = "Select * From viePermission_Tier_1 Where UserID=" & INTUSERID & " Order By CompanyID;"
''''''        com.CommandText = Mysql
''''''        Set RS = com.Execute
''''''        While Not RS.EOF
''''''            cmb.AddItem RS.Fields(3)
''''''            cmb.ItemData(cmb.NewIndex) = RS.Fields(2)
''''''            RS.MoveNext
''''''        Wend
'''''    End If
End Sub

Public Sub prcAddBank(cmb As ComboBox)
    Call prcSetConnectionActive
    Set R = New ADODB.Recordset
        strRecord = "Select Bankid, BankName From New_Bank Order By BankName;"
    R.Open strRecord, MainConn, adOpenStatic
        cmb.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmb.AddItem R("BankName")
                cmb.ItemData(cmb.NewIndex) = R("BankID")
                R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Sub prcLoadDeptUnit(cmb As ComboBox)
'    Call prcSetConnectionActive
    Set R = New ADODB.Recordset
        strRecord = "Select DeptUnitID, DeptUnitName From New_Department_Unit Order By DeptUnitID;"
       R.Open strRecord, MainConn, adOpenStatic
'        Set R = com.Execute
        cmb.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmb.AddItem R("DeptUnitName")
                cmb.ItemData(cmb.NewIndex) = R("DeptUnitID")
                R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Sub prcButtonPermission(strFormName As String)
'canInsert
'canUpdate
'canDelete
'canPreview
    Call prcSetConnectionActive
    Set RS = New ADODB.Recordset
    Mysql = "Select * From Permission_Tier_3 Where UserID = " & bytUserID & " and FormName = '" & strFormName & "'"
    RS.Open Mysql, MainConn, adOpenStatic
    If RS.EOF = False And RS.BOF = False Then
    While Not RS.EOF
        If RS.Fields(2) = "cmdInsert" And RS.Fields(3) = True Then
            canInsert = 1
        ElseIf RS.Fields(2) = "cmdUpdate" And RS.Fields(3) = True Then
            canUpdate = 1
        ElseIf RS.Fields(2) = "cmdDelete" And RS.Fields(3) = True Then
            canDelete = 1
        Else
            canInsert = 0
            canUpdate = 0
            canDelete = 0
            canPreview = 0
        End If
        RS.MoveNext
    Wend
    Else
        canInsert = 1
        canUpdate = 1
        canDelete = 1
        canPreview = 1
    End If
End Sub

Public Sub prcAddHead(ByRef cmbAccountsHead As ComboBox)
    Set R = New ADODB.Recordset
        If strFormName = "Acc Head For C_And_F" Then
            Mysql = "Select AccID, AccName From C_And_F_new_Acc_Head Order By AccID"
        Else
            Mysql = "Select HeadID, HeadName From AccountsHead Order By HeadID"
        End If
          R.Open Mysql, MainConn, adOpenStatic
    
'        Set R = com.Execute
        cmbAccountsHead.Clear
        If R.EOF = False And R.BOF = False Then
            R.MoveFirst
            Do While Not R.EOF
                If strFormName = "Acc Head For C_And_F" Then
                    cmbAccountsHead.AddItem R("AccName")
                    cmbAccountsHead.ItemData(cmbAccountsHead.NewIndex) = R("AccID")
                Else
                    cmbAccountsHead.AddItem R("HeadName")
                    cmbAccountsHead.ItemData(cmbAccountsHead.NewIndex) = R("HeadID")
                End If
            R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Sub prcAddSupplier(ByRef cmbSupplier As ComboBox)
    Set R = New ADODB.Recordset
        Mysql = "Select SupplierID, SupplierName From Supplier_Information Order By SupplierID"
        com.CommandText = Mysql
        Set R = com.Execute
        cmbSupplier.Clear
        If R.EOF = False And R.BOF = False Then
            R.MoveFirst
            Do While Not R.EOF
                cmbSupplier.AddItem R("SupplierID") & " " & R("SupplierName")
            R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Sub prcAddParty(ByRef cmbPartyName As ComboBox)
    Set R = New ADODB.Recordset
        If strFormName = "Party Name For C_and_F" Then
            Mysql = "Select PartyID, PartyName From C_And_F_New_Party Order By PartyName"
        Else
            Mysql = "Select PartyID, PartyName From Party Order By PartyName"
        End If
        com.CommandText = Mysql
        Set R = com.Execute
        
        cmbPartyName.Clear
        If R.EOF = False And R.BOF = False Then
            R.MoveFirst
            Do While Not R.EOF
                If strFormName = "Party Name For C_and_F" Then
                    cmbPartyName.AddItem R("PartyName")
                    cmbPartyName.ItemData(cmbPartyName.NewIndex) = R("PartyID")
                Else
                    cmbPartyName.AddItem R("PartyName")
                    cmbPartyName.ItemData(cmbPartyName.NewIndex) = R("PartyID")
                End If
            R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Sub prcAddBuyerName(ByRef cmb As ComboBox)
    Set R = New ADODB.Recordset
        strRecord = "Select * From Buyer_Information Order By BuyerName;"
        R.Open strRecord, MainConn, adOpenStatic
        cmb.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmb.AddItem R.Fields(1).Value
                cmb.ItemData(cmb.NewIndex) = R.Fields(0).Value
                R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Sub prcAddExpenses(ByRef cmb As ComboBox)
    Set R = New ADODB.Recordset
        strRecord = "Select * From C_And_F_New_Expense Order By ExpenseName;"
        R.Open strRecord, MainConn, adOpenStatic
        cmb.Clear
        If R.EOF = False And R.BOF = False Then
            Do While Not R.EOF
                cmb.AddItem R.Fields(1).Value
                cmb.ItemData(cmb.NewIndex) = R.Fields(0).Value
                R.MoveNext
            Loop
        End If
    Set R = Nothing
End Sub

Public Function MaxStatementNo(intTranType As String) As Integer
Dim CurrentDate As Date
Dim intYear As Integer
    On Error GoTo er1
    intYear = Year(Date)
    CurrentDate = "01-Jan-" & intYear
    Set RS = New ADODB.Recordset
        Mysql = "SELECT Max(StatementNo) AS MaxStatementNo From Statement_of_Cash_and_Credit " & _
            " Where Tran_Date >= '" & CurrentDate & "' And Tran_Type ='" & intTranType & "' And CompanyID= " & bytCompID & ""
     RS.Open Mysql, MainConn, adOpenStatic
        MaxStatementNo = RS("MaxStatementNo") + 1
    Set RS = Nothing
    Exit Function
er1:
'    If Err.Number = 3021 Then
        MsgBox "Well-Come to New Year.", vbInformation + vbOKOnly
        MaxStatementNo = 1
        Set RS = Nothing
'    End If
End Function

Public Sub prcSetConnectionActive()
    Set com = New Command
    com.ActiveConnection = MainConn
    Set R = New Recordset
    Set RS = New Recordset
    Set rs1 = New Recordset
    Set rs2 = New Recordset
    Set rs3 = New Recordset
End Sub


Public Sub prcNumCheck(num As Integer)
Dim strValid As String
    strValid = "0123456789.-"
    If num > 26 Then
        If InStr(strValid, Chr(num)) = 0 Then num = 0
    End If
End Sub

Public Function findslnoMAX(CboGPID As String) As Integer
     Set RS = New ADODB.Recordset
    strRecord = "SELECT MAX(SLNO) AS MAX FROM  Cheque_Child_Table WHERE IDNO=" & CboGPID & " "
    RS.Open strRecord, MainConn, adOpenStatic
    If RS.RecordCount > 0 Then
        If RS![Max] = Null Then
           findslnoMAX = 0
         Else
          findslnoMAX = RS![Max]
         End If
    End If
    RS.Close
    Set RS = Nothing
End Function
Public Function findmxslnobill(CboGPID As String) As Integer
     Set RS = New ADODB.Recordset
'    strRecord = "SELECT MAX(SLNO) AS MAX FROM  billdetails WHERE billID=" & CboGPID & " "
    strRecord = "SELECT     MAX(dbo.BillDetails.SLNO) AS MAX  FROM     dbo.BillDetails INNER JOIN      dbo.BillMain ON dbo.BillDetails.BillID = dbo.BillMain.BillID WHERE     (dbo.BillMain.IDNO = '" & CboGPID & "')"
        
    RS.Open strRecord, MainConn, adOpenStatic
    If RS.RecordCount > 0 Then
        If IsNull(RS![Max]) = True Then
           findmxslnobill = 0
         Else
          findmxslnobill = RS![Max]
         End If
    End If
    RS.Close
    Set RS = Nothing
End Function
