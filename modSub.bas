Attribute VB_Name = "modSub"
Public ComName As String    ''For Entry Computer Name

Public Sub addUserName(cboMain As ComboBox)
    cboMain.Clear
    Set R1 = New ADODB.Recordset
    strRecord = "SELECT UserName from InventoryPJL.dbo.Sys_User_Name;"
    R1.Open strRecord, MainConn, adOpenStatic
    If R1.RecordCount > 0 Then
        Do Until R1.EOF
            cboMain.AddItem DecryptIt(R1![UserName], 11)
        R1.MoveNext
        Loop
    End If
    R1.Close
    Set R1 = Nothing
End Sub

'=== Combo Yes/No ==='
Sub addYesNo(cboMain As ComboBox)
  cboMain.Clear
  cboMain.AddItem "YES"
  cboMain.AddItem "NO"
End Sub

Function findPermID(parPerm As String) As Integer
  findPermID = 0
  Set R1 = New ADODB.Recordset
  strRecord = "SELECT PermID FROM Sys_Default_Permission WHERE PermName='" & parPerm & "';"
  R1.Open strRecord, MainConn, adOpenStatic
  If R1.RecordCount > 0 Then findPermID = R1![PermID]
  R1.Close
  Set R1 = Nothing
End Function

Public Sub prcMakeCenter(frm As Form)
    Dim x As Double, y As Double
    x = (Screen.Width - frm.Width) / 2
    y = (Screen.Height - frm.Height) / 2
    
    frm.Move x, y
End Sub

Public Function fncGotFocus(txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Function

Public Function findItemType(parItem As String) As String
    findItemType = ""
    Set R1 = New ADODB.Recordset
    strRecord = "SELECT itemType FROM new_Item_Name WHERE itemName='" & parItem & "';"
    R1.Open strRecord, MainConn, adOpenStatic
    If R1.RecordCount > 0 Then findItemType = R1![itemType]
    R1.Close
End Function

Public Sub addItemNameAll(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT itemName FROM new_Item_Name;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![itemName]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub
Public Sub addSupplierNameD(cboMain As DataCombo)
  cboMain.Text = ""
  Set R = New ADODB.Recordset
  strRecord = "SELECT supplierID,suppName FROM New_Supplier ORDER BY suppName;"
  R.Open strRecord, MainConn, adOpenStatic
  Set cboMain.RowSource = R
  cboMain.ListField = "suppName"
  cboMain.BoundColumn = "supplierID"
    
'    If R.RecordCount > 0 Then
'        Do Until R.EOF
'            cboMain.AddItem UCase(R![suppName])
'            cboMain.ItemData(cboMain.NewIndex) = R![supplierID]
'        R.MoveNext
'        Loop
'    End If
'    R.Close
    Set R = Nothing
End Sub

Public Sub addSupplierName(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT supplierID,suppName FROM New_Supplier ORDER BY supplierID;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem UCase(R![suppName])
            cboMain.ItemData(cboMain.NewIndex) = R![supplierID]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addSupplier(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    
    
    strRecord = "SELECT DISTINCT new_Supplier.supplierID,new_Supplier.SuppName FROM Pro_Inv_Main  INNER JOIN Company_Information ON (Pro_Inv_Main.compID = Company_Information.compID) " _
               & " INNER JOIN new_Supplier ON (Pro_Inv_Main.SuppID = new_Supplier.SupplierID)  Inner Join new_PI ON (Pro_Inv_Main.PIID = " _
               & " new_PI.PIID)  INNER JOIN Pro_Inv_Sub ON (Pro_Inv_Main.PIID = Pro_Inv_Sub.PIID) Inner Join new_Item_Name On (Pro_Inv_Sub.ItemID = new_Item_Name.ItemID) " _
               & " WHERE  Pro_Inv_Main.compID =" & bytCompID & " AND Pro_Inv_Sub.DeleteRow = 0 order by new_Supplier.SuppName;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem UCase(R![suppName])
            cboMain.ItemData(cboMain.NewIndex) = R![supplierID]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addSupplierNameList(cboMainList As DataCombo)
  Set rsDC = New ADODB.Recordset
  strRecord = "SELECT supplierID,suppName FROM New_Supplier ORDER BY suppName;"
  rsDC.Open strRecord, MainConn, adOpenStatic
  Set cboMainList.RowSource = rsDC
  cboMainList.ListField = "suppName"
  cboMainList.BoundColumn = "supplierID"
End Sub

'Public Sub addSuppPINo(cboMain As ComboBox, parSupplier As String)
'    cboMain.Clear
'    Set R = New ADODB.Recordset
'    strRecord = "SELECT Pro_Inv_Main.PINO FROM new_Supplier INNER JOIN Pro_Inv_Main ON new_Supplier.supplierID = Pro_Inv_Main.suppID" _
'        & " WHERE (((new_Supplier.suppName)='" & parSupplier & "') AND ((Pro_Inv_Main.compID)=" & bytCompID & "));"
'    R.Open strRecord, MainConn, adOpenStatic
'    If R.RecordCount > 0 Then
'        Do Until R.EOF
'            cboMain.AddItem R![PINO]
'        R.MoveNext
'        Loop
'    End If
'    R.Close
'    Set R = Nothing
'End Sub
'Public Sub addSuppPINo(cboMain As ComboBox, parSupplier As String)
'    cboMain.Clear
'    Set R = New ADODB.Recordset
'        strRecord = "SELECT NEW_PI.PIID, NEW_PI.PINo FROM NEW_PI INNER JOIN Pro_Inv_Main ON NEW_PI.PIID = Pro_Inv_Main.PIID INNER JOIN new_Supplier ON new_Supplier.supplierID = Pro_Inv_Main.suppID" _
'            & " WHERE (((new_Supplier.suppName)='" & parSupplier & "') AND ((Pro_Inv_Main.compID)=" & bytCompID & "));"
'        R.Open strRecord, MainConn, adOpenStatic
'        If R.RecordCount > 0 Then
'            Do Until R.EOF
'                cboMain.AddItem R![PINo]
'                cboMain.ItemData(cboMain.NewIndex) = R![PIId]
'                R.MoveNext
'            Loop
'        End If
'        R.Close
'    Set R = Nothing
'End Sub

Public Sub addSuppPINo(cboMain As ComboBox, parSupplier As String)
cboMain.Clear
Set R = New ADODB.Recordset
strRecord = "SELECT NEW_PI.PIID,NEW_PI.PINo FROM NEW_PI INNER JOIN Pro_Inv_Main ON NEW_PI.PIID = Pro_Inv_Main.PIID INNER JOIN new_Supplier ON new_Supplier.supplierID = Pro_Inv_Main.suppID" _
          & " WHERE new_Supplier.suppName='" & parSupplier & "' AND Pro_Inv_Main.compID=" & bytCompID & ";"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  Do Until R.EOF
    cboMain.AddItem R![PINo]
    cboMain.ItemData(cboMain.NewIndex) = R![PIId]
    R.MoveNext
  Loop
End If
R.Close
Set R = Nothing
End Sub

Public Sub EditSuppPINo(cboMain As ComboBox, parSupplier As Long)
cboMain.Clear
Set R = New ADODB.Recordset
strRecord = "Select PONO,POID From ThreadInfo WHERE SUPPID='" & parSupplier & "' AND compID=" & bytCompID & " AND DeleteRow = 0;"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  Do Until R.EOF
    cboMain.AddItem R![PONO]
    cboMain.ItemData(cboMain.NewIndex) = R![POID]
    R.MoveNext
  Loop
End If
R.Close
Set R = Nothing
End Sub
Public Sub addSuppItem(cboMain As ComboBox, parSupplier As String)
cboMain.Clear
Set R = New ADODB.Recordset
'strRecord = "Select PONO,POID From ThreadInfo WHERE compID=" & bytCompID & " AND DeleteRow = 0;"
strRecord = " SELECT DISTINCT S.itemID, N.itemName, P.suppName, M.suppID " _
          & " FROM  dbo.MRR_Challan_Sub S INNER JOIN dbo.new_Item_Name N ON S.itemID = N.itemID INNER JOIN dbo.MRR_Challan_Main M ON S.MRRID = M.MRRID INNER JOIN dbo.new_Supplier P ON M.suppID = P.supplierID WHERE suppName='" & parSupplier & "' ORDER BY itemName "
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  Do Until R.EOF
    cboMain.AddItem R![itemName]
    cboMain.ItemData(cboMain.NewIndex) = R![ITEMID]
    R.MoveNext
  Loop
End If
R.Close
Set R = Nothing
End Sub
Public Sub addThreadPONo(cboMain As ComboBox)
cboMain.Clear
Set R = New ADODB.Recordset
strRecord = "Select PONO,POID From ThreadInfo WHERE compID=" & bytCompID & " AND DeleteRow = 0;"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  Do Until R.EOF
    cboMain.AddItem R![PONO]
    cboMain.ItemData(cboMain.NewIndex) = R![POID]
    R.MoveNext
  Loop
End If
R.Close
Set R = Nothing
End Sub
Public Sub addSuppPINoOC(cboMain As ComboBox, parSupplier As String)
cboMain.Clear
Set R = New ADODB.Recordset
strRecord = "SELECT NEW_PI.PIID, NEW_PI.PINo FROM NEW_PI INNER JOIN Pro_Inv_MainOC ON NEW_PI.PIID = Pro_Inv_MainOC.PIID INNER JOIN new_Supplier ON new_Supplier.supplierID = Pro_Inv_MainOC.suppID" _
          & " WHERE new_Supplier.suppName='" & parSupplier & "' AND Pro_Inv_MainOC.compID=" & bytCompID & " AND NEW_PI.DeleteRow = 0;"
R.Open strRecord, MainConn, adOpenStatic
If R.RecordCount > 0 Then
  Do Until R.EOF
    cboMain.AddItem R![PINo]
    cboMain.ItemData(cboMain.NewIndex) = R![PIId]
    R.MoveNext
  Loop
End If
R.Close
Set R = Nothing
End Sub


Public Sub addSuppPIMRRNo(cboManin As ComboBox, cboMRRMain As ComboBox, parSupplier As String)
    cboMRRMain.Clear
    
    Set R = New ADODB.Recordset
    
     strRecord = "Select Distinct New_MRR.MRRID,New_MRR.MRRNO From NEW_PI" _
        & " Inner Join Pro_Inv_Main On NEW_PI.PIID = Pro_Inv_Main.PIID " _
        & " Inner Join MRR_Challan_Sub On MRR_Challan_Sub.PIID=Pro_Inv_Main.PIID " _
        & " Inner Join MRR_Challan_Main On MRR_Challan_Main.MRRID=MRR_Challan_Sub.MRRID And MRR_Challan_Main.DeleteRow=MRR_Challan_Sub.DeleteRow " _
        & " Inner Join New_MRR On New_MRR.MRRID=MRR_Challan_Sub.MRRID And New_MRR.DeleteRow=MRR_Challan_Sub.DeleteRow " _
        & " INNER JOIN new_Supplier ON new_Supplier.supplierID = MRR_Challan_Main.suppID " _
        & " WHERE new_Supplier.suppName='" & parSupplier & "' AND NEW_PI.PINo='" & cboManin.Text & "' And Pro_Inv_Main.compID=" & bytCompID & " AND NEW_PI.DeleteRow = 0 And  NEW_MRR.DeleteRow = 0 " _
        & " And New_MRR.DeleteRow= 0 And MRR_Challan_Sub.DeleteRow=0;"
    
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then
            Do Until R.EOF
                cboMRRMain.AddItem R![MRRNo]
                cboMRRMain.ItemData(cboMRRMain.NewIndex) = R![MRRID]
                R.MoveNext
            Loop
        End If
        R.Close
    Set R = Nothing
End Sub

Public Function findSupplierID(parSupplier As String) As Integer
    Set R = New ADODB.Recordset
    strRecord = "SELECT supplierID FROM new_Supplier WHERE suppName='" & parSupplier & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findSupplierID = R![supplierID]
    R.Close
End Function

Public Function findSupplierDetails(parSupplier As String) As String
    Set R = New ADODB.Recordset
    strRecord = "SELECT suppAddress FROM New_Supplier WHERE suppName='" & parSupplier & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findSupplierDetails = R![suppAddress]
    R.Close
    Set R = Nothing
End Function


Public Sub addItemType(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT itemType FROM new_Item_Type;"
    R.Open strRecord, MainConn, adOpenStatic
  
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![itemType]
            R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addSuppId(cboMain As ComboBox, a As Integer)
cboMain.Clear
Dim tmprst As New Recordset
'        If a = 1 Then
'        Mysql = "Select Distinct ACCNO From New_ACCID Where ComID = " & bytCompID & " and TransType = 'Cash' and Deleterow=0;"
'        ElseIf a = 2 Then
'        Mysql = "Select Distinct ACCNO From New_ACCID Where ComID = " & bytCompID & " and TransType = 'Credit' and Deleterow=0;"
'        End If

        If a = 1 Then
        Mysql = "Select Distinct SCNo From New_SCID Where ComID = " & bytCompID & " and TransType = 'Cash' and Deleterow=0;"
        ElseIf a = 2 Then
        Mysql = "Select Distinct SCNo From New_SCID Where ComID = " & bytCompID & " and TransType = 'Credit' and Deleterow=0;"
        End If
        tmprst.Open Mysql, MainConn, adOpenStatic
        If tmprst.EOF = False And tmprst.BOF = False Then
            cboMain.Clear
            tmprst.MoveFirst
            Do While Not tmprst.EOF
                cboMain.AddItem tmprst("SCNo")
                tmprst.MoveNext
            Loop
        Else
            cboMain.Clear
        End If
    Set tmprst = Nothing
End Sub


Public Sub addItemName(cboMain As ComboBox, parItGroup As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT itemID,itemName FROM new_Item_Name WHERE itemType = '" & parItGroup & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![itemName]
            cboMain.ItemData(cboMain.NewIndex) = R![ITEMID]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub
Public Sub addUnit(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT UnitName FROM new_Unit_Name;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![UnitName]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Function findItemID(parItem As String) As Integer
    Set R = New ADODB.Recordset
    strRecord = "SELECT itemID FROM new_Item_Name WHERE itemName='" & parItem & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findItemID = R![ITEMID]
    R.Close
End Function
Public Sub addCountry(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT CountryName FROM new_Country;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![CountryName]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPFNo(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT PFNo FROM new_PF_No;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![PFNO]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub


Public Sub addPrNO(cboMain As ComboBox, cbopono As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
   
    strRecord = "SELECT  DISTINCT   dbo.Pro_Inv_Sub.PRNo FROM dbo.New_PI INNER JOIN " _
                      & " dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID where dbo.Pro_Inv_Sub.DeleteRow=0 And dbo.New_PI.PINo ='" & cbopono & "';"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![PRNo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPrNOOC(cboMain As ComboBox, cbopono As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
   
    strRecord = "SELECT  DISTINCT   dbo.Pro_Inv_SubOC.PRNo FROM dbo.New_PI INNER JOIN " _
                      & " dbo.Pro_Inv_SubOC ON dbo.New_PI.PIID = dbo.Pro_Inv_SubOC.PIID where dbo.Pro_Inv_SubOC.DeleteRow=0 And dbo.New_PI.PINo ='" & cbopono & "';"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![PRNo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPOITEM(cboMain As ComboBox, cbopono As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
   
    strRecord = "SELECT DISTINCT dbo.new_Item_Name.itemName" _
                & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID INNER JOIN" _
                & " dbo.new_Item_Name ON dbo.Pro_Inv_Sub.itemID = dbo.new_Item_Name.itemID where dbo.Pro_Inv_Sub.DeleteRow=0 And dbo.New_PI.PINo ='" & cbopono & "';"

    
    
'    strRecord = " SELECT DISTINCT dbo.new_Item_Name.itemName " _
'              & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_Sub ON dbo.New_PI.PIID = dbo.Pro_Inv_Sub.PIID INNER JOIN dbo.new_Item_Name ON " _
'              & " dbo.Pro_Inv_Sub.itemID = dbo.new_Item_Name.itemID INNER JOIN dbo.Pro_Inv_Main ON dbo.New_PI.PIID = dbo.Pro_Inv_Main.PIID INNER JOIN " _
'              & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID where dbo.Pro_Inv_Sub.DeleteRow=0 And  dbo.new_Supplier.suppName='" & cmbSupplier & "' AND  dbo.New_PI.PINo ='" & cbopono & "' "
'
    
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![itemName]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPOITEMOC(cboMain As ComboBox, cbopono As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
   
    strRecord = "SELECT DISTINCT dbo.new_Item_Name.itemName" _
                & " FROM dbo.New_PI INNER JOIN dbo.Pro_Inv_SubOC ON dbo.New_PI.PIID = dbo.Pro_Inv_SubOC.PIID INNER JOIN" _
                & " dbo.new_Item_Name ON dbo.Pro_Inv_SubOC.itemID = dbo.new_Item_Name.itemID where dbo.Pro_Inv_SubOC.DeleteRow=0 And dbo.New_PI.PINo ='" & cbopono & "';"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![itemName]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPINO(cboMain As ComboBox, cboSupplier As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
    'strRecord = "SELECT PINO FROM New_PI  ;"
    strRecord = "SELECT dbo.New_PI.PINo FROM dbo.Pro_Inv_Main INNER JOIN" _
                     & " dbo.New_PI ON dbo.Pro_Inv_Main.PIID = dbo.New_PI.PIID AND dbo.Pro_Inv_Main.compID = dbo.New_PI.ComID INNER JOIN" _
                      & " dbo.new_Supplier ON dbo.Pro_Inv_Main.suppID = dbo.new_Supplier.supplierID  where dbo.new_Supplier.suppName='" & cboSupplier & "';"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![PINo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPINOOC(cboMain As ComboBox, cboSupplier As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
    'strRecord = "SELECT PINO FROM New_PI  ;"
    strRecord = "SELECT dbo.New_PI.PINo FROM dbo.Pro_Inv_MainOC INNER JOIN" _
                     & " dbo.New_PI ON dbo.Pro_Inv_MainOC.PIID = dbo.New_PI.PIID AND dbo.Pro_Inv_MainOC.compID = dbo.New_PI.ComID INNER JOIN" _
                      & " dbo.new_Supplier ON dbo.Pro_Inv_MainOC.suppID = dbo.new_Supplier.supplierID  where dbo.new_Supplier.suppName='" & cboSupplier & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
        
            cboMain.AddItem R![PINo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addCurrency(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    strRecord = "SELECT CurrName FROM new_Currency;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![CurrName]
            R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Function findDCurrency() As String
    Set R = New ADODB.Recordset
    strRecord = "SELECT CurrName FROM new_Currency WHERE setDefault=1;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findDCurrency = R![CurrName]
    R.Close
    Set R = Nothing
End Function

Public Sub prcAddSuppMrrID(cboMain As ComboBox, parSupplier As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
    
    strRecord = " SELECT New_MRR.MRRNo FROM New_MRR INNER JOIN MRR_Challan_Main ON New_MRR.ComID = MRR_Challan_Main.compID AND " _
            & " New_MRR.MRRID = MRR_Challan_Main.MRRID INNER JOIN new_Supplier ON MRR_Challan_Main.suppID = new_Supplier.supplierID " _
            & " WHERE (((new_Supplier.suppName)='" & parSupplier & "') AND ((MRR_Challan_Main.compID)=" & bytCompID & ")) AND New_MRR.DELETEROW = 0 ORDER BY MRR_Year DESC, MRR_No DESC;"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
    
            cboMain.AddItem R![MRRNo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub prcAddSuppMrrIDOC(cboMain As ComboBox, parSupplier As String)
    cboMain.Clear
    Set R = New ADODB.Recordset
    
    strRecord = " SELECT New_MRR.MRRNo FROM New_MRR INNER JOIN MRR_Challan_MainOC ON New_MRR.ComID = MRR_Challan_MainOC.compID AND " _
            & " New_MRR.MRRID = MRR_Challan_MainOC.MRRID INNER JOIN new_Supplier ON MRR_Challan_MainOC.suppID = new_Supplier.supplierID " _
            & " WHERE (((new_Supplier.suppName)='" & parSupplier & "') AND ((MRR_Challan_MainOC.compID)=" & bytCompID & ")) AND New_MRR.DELETEROW = 0;"
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            
            cboMain.AddItem R![MRRNo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Function findPIID(parPINo As String) As Long
    Set R = New ADODB.Recordset
    strRecord = "SELECT PIID FROM new_PI WHERE PINO='" & parPINo & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findPIID = R![PIId]
    R.Close
End Function

Public Function findMRRID(parMRRID As String) As Integer
    Set R = New ADODB.Recordset
    strRecord = "SELECT MRRID FROM new_MRR WHERE MRRNO='" & parMRRID & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findMRRID = R![MRRID]
    R.Close
End Function

Function fncMRR(InputID As Integer) As String
    If Val(InputID) > 0 And Val(InputID) < 10 Then
        fncMRR = "0000" & Val(InputID)
    ElseIf Val(InputID) > 9 And Val(InputID) < 100 Then
        fncMRR = "000" & Val(InputID)
    ElseIf Val(InputID) > 99 And Val(InputID) < 1000 Then
        fncMRR = "00" & Val(InputID)
    ElseIf Val(InputID) > 999 And Val(InputID) < 10000 Then
        fncMRR = "0" & Val(InputID)
    ElseIf Val(InputID) > 9999 And Val(InputID) < 100000 Then
        fncMRR = Val(InputID)
    End If
End Function

Public Sub prcAddSuppMrrno(cboMain2 As TextBox, cmbMRRID As String, parSupplier As String)
    Set R = New ADODB.Recordset
    
'    strRecord = " SELECT New_MRR.MRRid FROM New_MRR where DeleteRow =0 AND mrrno= '" & cmbMRRID & "';"
''02-JUN-2016
    strRecord = "SELECT  dbo.New_MRR.MRRID " _
              & " FROM dbo.New_MRR INNER JOIN  dbo.MRR_Challan_Main ON dbo.New_MRR.MRRID = dbo.MRR_Challan_Main.MRRID INNER JOIN  dbo.new_Supplier ON dbo.MRR_Challan_Main.suppID = dbo.new_Supplier.supplierID " _
              & " WHERE(dbo.New_MRR.DeleteRow = 0) AND (dbo.New_MRR.MRRNo = '" & cmbMRRID & "')AND new_Supplier.suppName='" & parSupplier & "' "
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
       Do Until R.EOF
           cboMain2 = R![MRRID]
        R.MoveNext
       Loop
    End If
    R.Close
    Set R = Nothing
End Sub


Public Sub prcAddSuppMrrnoOC(cboMain2 As TextBox, cmbMRRID As String, parSupplier As String)
    Set R = New ADODB.Recordset
    
'    strRecord = " SELECT New_MRR.MRRid FROM New_MRR where DeleteRow =0 AND mrrno= '" & cmbMRRID & "';"
''02-JUN-2016
    strRecord = "SELECT  dbo.New_MRR.MRRID " _
              & " FROM dbo.New_MRR INNER JOIN  dbo.MRR_Challan_MainOC ON dbo.New_MRR.MRRID = dbo.MRR_Challan_MainOC.MRRID INNER JOIN  dbo.new_Supplier ON dbo.MRR_Challan_MainOC.suppID = dbo.new_Supplier.supplierID " _
              & " WHERE(dbo.New_MRR.DeleteRow = 0) AND (dbo.New_MRR.MRRNo = '" & cmbMRRID & "')AND new_Supplier.suppName='" & parSupplier & "' "
    
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
       Do Until R.EOF
           cboMain2 = R![MRRID]
        R.MoveNext
       Loop
    End If
    R.Close
    Set R = Nothing
End Sub
Public Sub prCMAX(vmrrid As Double)
    
    Set R = New ADODB.Recordset
    strRecord = "SELECT MAX(SLNO) AS MAX FROM Mrr_Challan_SUB WHERE MRRID=" & vmrrid & " "
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
       'Do Until R.EOF
           intSLNO = R![Max]
       ' R.MoveNext
       'Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub prCMAXOC(vmrrid As Double)
    
    Set R = New ADODB.Recordset
    strRecord = "SELECT MAX(SLNO) AS MAX FROM Mrr_Challan_SUBOC WHERE MRRID=" & vmrrid & " "
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
       'Do Until R.EOF
           intSLNO = R![Max]
       ' R.MoveNext
       'Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Sub addPONO(cboMain As ComboBox)
    cboMain.Clear
    Set R = New ADODB.Recordset
    'strRecord = "SELECT PINO FROM New_PI  ;"
    strRecord = "SELECT Distinct PINo FROM New_PI Where DeleteRow=0 AND PI_Year>2015 ;"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then
        Do Until R.EOF
            cboMain.AddItem R![PINo]
        R.MoveNext
        Loop
    End If
    R.Close
    Set R = Nothing
End Sub

Public Function findempID(CompId As Integer, EmpId As String, paremp As String) As Double
    
    Set R = New ADODB.Recordset
    strRecord = "SELECT empID FROM infoemp WHERE CompID=" & CompId & " and Empcode='" & EmpId & "' and empName='" & paremp & "';"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findempID = R![EmpId]
    R.Close

End Function
Public Function FindSCID(SCID As String) As Integer
Set RS = New ADODB.Recordset
                    strRecord = "SELECT SCID FROM New_SCID Where ComID = " & bytCompID & " And SCNO='" & SCID & "';"
                    MainComm.CommandText = strRecord
                RS.Open strRecord, MainConn, adOpenStatic
                    If Not IsNull(RS![SCID]) Then
                        FindSCID = RS.Fields(0).Value
                    End If
                    RS.Close
End Function
