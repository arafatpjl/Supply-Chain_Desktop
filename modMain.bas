Attribute VB_Name = "modMain"
Global MainConn As ADODB.Connection
Global MainComm As ADODB.Command
Global R As ADODB.Recordset
Global RS As ADODB.Recordset
Global R1 As ADODB.Recordset
Global strRecord As String
Global strRecord1 As String
Global Mysql As String
Global Appl As New CRAXDRT.Application
Global REPORT As New CRAXDRT.REPORT
'----------- Varriables --------'
'==============================='

Global bytCompID As Byte
Global bytUserID As Byte
Global strUser As String
Global bytEditSalary As Byte
Global strUserStatus As String
Global strComNameTxt As String
Public WSName As String                                 ''  Server For Work
Public WDBName As String                                ''  Database For Work
Public WDBPrefix As String                              ''  Prefix For Working Table
Public Const cnstPrimaryDB As String = "InventoryPJL"   ''  Check primary Database Connection or not
Public ReportPath As String                             ''  For Report Path


Global boUserEdit As Boolean
Global isPrimaryDB As Boolean
Global CheckComCon As Boolean
Global strDataPath As String
Global strRptPath As String
Global strMark As String
Global strImagePath As String

Global inti As Integer
Global intYear As Integer
Global boValidDt As Boolean
Global boEmptyText As Boolean
Global CheckCompanyText As Boolean
Global strBegDate As String
Global strEndDate As String
Global strNullDate As String
Global intSLNO As Integer

Global lngLastID As Long

'----------- Constants ---------'
'==============================='
'Global Const cnstDtFrmtI As String = "MMM-DD-YYYY"
'Global Const cnstDtFrmtQ As String = "MMM DD YYYY"
Global Const cnstDtFrmtI As String = "DD-MMM-YYYY"
Global Const cnstDtFrmtQ As String = "DD-MMM-YYYY"
Global Const cnstDtFrmtD As String = "DD-MMM-YYYY"

Global Const cnstMsgInfo As String = "Message"
Global Const cnstMsgErDB As String = "Error"
Global Const cnstMsgQ As String = "Confirmation"

'Public ReportPath As String
Public PeriodYear As Integer
'---------Global Variable For ComName & ComDriveVolumnno-------
Global getComName As String
Global getComVolName As String
Global getComVolID As Integer

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

''Function findServerName() As String
''    If boValidConn = True Then
''        Open App.Path + "\MainDataPathJ2L.ini" For Input As #1
''    Else
''        Open App.Path + "\MainDataPath.ini" For Input As #1
''    End If
''
''    Input #1, strDataPath
''    findServerName = strDataPath
''    Close #1
''
''    If Len(findServerName) = 0 Then
''        MsgBox "Server not found", vbCritical, cnstMsgErDB
''    End If
''End Function

Public Sub Main()
  
  isPrimaryDB = True
  
  Call ConnectDatabase
  
  Call prcCheckComCon
  
  If CheckComCon = True Then
    
    Set R = New ADODB.Recordset
    strRecord = "SELECT VersionNo FROM Version;"
    R.Open strRecord, MainConn, adOpenStatic
    If R![VersionNo] = 1.5 Then
      frmSplash.Show
    Else
      MsgBox "Software version not match." & vbCrLf & "Please Update or Run latest version from Server.", vbCritical, "Update Information"
    End If
        'frmSplash.Show
  End If
  
  isPrimaryDB = False
        
' MDIfrmMain.Show
' isPrimaryDB = False
End Sub

Function ServerName() As String
    Open App.Path + "\MainDataPath.ini" For Input As #1
    Input #1, ServerName
    strServer = ServerName
    Close #1
    
    If Len(ServerName) = 0 Then
        MsgBox "Server not found", vbCritical, cnstMsgErDB
    End If
End Function

Sub ConnectDatabase()

On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    
    Dim StrConn As String
    
    Set MainConn = New ADODB.Connection
    If MainConn.State = 1 Then MainConn.Close

    If isPrimaryDB = True Then
        StrConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=SA; PASSWORD=SOFT2O12; " _
            & " Initial Catalog=" & cnstPrimaryDB & "; Data Source='" & ServerName & "'"
    Else
        StrConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=SA; PASSWORD=SOFT2O12; " _
            & " Initial Catalog= " & WDBName & "; Data Source='" & WSName & "'"
    End If

''    If isPrimaryDB = True Then
'''        StrConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=SVR2010; PASSWORD=SOFT2O12; " _
'''            & " Initial Catalog=" & cnstPrimaryDB & "; Data Source='" & findServerName & "'"
''
''      StrConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=SVR2010; PASSWORD=SOFT2O12; " _
''              & " Initial Catalog=" & cnstPrimaryDB & "; Data Source='" & findServerName & "'"
''    Else
''      StrConn = "Provider=SQLOLEDB.1; Persist Security Info=False; User ID=SVR2010; PASSWORD=SOFT2O12; " _
''              & " Initial Catalog= " & WDBName & "; Data Source='" & WSName & "'"
''    End If
    
    MainConn.Open StrConn
    
    Screen.MousePointer = vbDefault

Exit Sub
ErrorHandler:
    If err.Number = -2147467259 Then
        MsgBox "Database file not found." & vbCrLf & "Please Contact to ICT (Software) Department.", vbCritical, cnstMsgErDB
    ElseIf err.Number = 53 Then
        MsgBox "Server File not found." & vbCrLf & "Please Contact to ICT (Software) Department.", vbCritical, cnstMsgErDB
    Else
        MsgBox err.Description, vbCritical, cnstMsgErDB
    End If
    
    Screen.MousePointer = vbDefault
    End
End Sub

Public Sub prcConReport(obj As Object)
    With obj
'        .LogonInfo(0) = "DSN= " & WSName & "; UID=SVR2010; PWD=SOFT20$12; DSQ=" & WDBName & ""
'        .LogOnServer "p2ssql.dll", "" & WSName & "", "" & WDBName & "", "SVR2010", "SOFT20$12"
        '.LogOnServer "p2sSQL.dll", "" & WSName & "", "" & WDBName & "", "SVR2010", "SOFT20$12"
    End With
End Sub

Function FormatDate(txtMain As TextBox)
    If Len(txtMain) = 0 Then Exit Function
    If Len(txtMain) = 11 And (txtMain.Text Like "*-*-*") = True Then
        FormatDate = txtMain
    Exit Function
    End If

    Dim strDate As String
    boValidDt = True

    On Error GoTo ErrorHandler

    If Len(txtMain) = 6 Then
        For inti = 1 To 6 Step 3
            strDate = strDate + Left(txtMain, 2) + "-"
            txtMain = Right(txtMain, 4)
        Next inti
        strDate = strDate + Right(txtMain, 2)
        FormatDate = Format(CDate(strDate), cnstDtFrmtD)
        txtMain.Text = FormatDate
    Else
        MsgBox "Enter 6 Digit", vbInformation, cnstMsgInfo
        txtMain.Text = ""
        txtMain.SetFocus
        boValidDt = False
    End If
  
ErrorHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbCritical, cnstMsgErDB
        txtMain.Text = ""
        txtMain.SetFocus
        boValidDt = False
    End If
End Function

Function frmtTime(txtMain As TextBox) As String
    If Len(txtMain) = 0 Then Exit Function
    If Len(txtMain) = 5 And txtMain Like "*:*" = True Then
        frmtTime = txtMain
    Exit Function
    End If

    ValidDate = True
    If Len(txtMain) <> 4 Then
        MsgBox "Enter 4 Digits", vbInformation, cnstMsgInfo
        txtMain.SetFocus
        ValidDate = False
    Exit Function
    End If

    Dim strTime As String
    strTime = ""

On Error GoTo ErrorHandler
    strTime = Left(txtMain, 2) + ":" + Right(txtMain, 2)
    frmtTime = Format(CDate(strTime), "HH:MM")
    txtMain.Text = frmtTime

ErrorHandler:
    If err.Number <> 0 Then
        If err.Number = 13 Then
            MsgBox "Invalid 'Hour' or 'Minute'", vbInformation, cnstMsgInfo
        Else
            MsgBox err.Description, vbCritical, cnstMsgErDB
        End If
        txtMain.Text = ""
        txtMain.SetFocus
        ValidDate = False
    err.Number = 0
    End If
End Function

Function DateFunc(KeyAscii As Integer) As String
    Dim strValid As String
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Function

Function CurrFunc(KeyAscii As Integer) As String
    Dim strValid As String
    strValid = "0123456789."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Function

Function CurrFuncForMrrRecv(KeyAscii As Integer) As String
    Dim strValid As String
    strValid = "0123456789."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Function

Function SingleCodeFunc(KeyAscii As Integer) As String
    If KeyAscii = 39 Then KeyAscii = 0
End Function

Sub txtGotFocus(txtMain As TextBox)
    With txtMain
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Sub cboGotFocus(cboMain As ComboBox)
    With cboMain
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'Function findCompID(parComp As String) As Byte
'    findCompID = 0
'    Set R1 = New ADODB.Recordset
'    strRecord = "SELECT CompID FROM Company_Information WHERE CompName='" & parComp & "';"
'    R1.Open strRecord, MainConn, adOpenStatic
'    If R1.RecordCount > 0 Then findCompID = R1![CompID]
'    R1.Close
'    Set R1 = Nothing
'End Function

Function findCompName() As String
    Set R = New ADODB.Recordset
    strRecord = "SELECT CompName FROM Company_Information WHERE CompID=" & bytCompID & ";"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findCompName = R![CompName] Else findCompName = ""
    R.Close
    Set R = Nothing
End Function

Function findCompAdd() As String
    Set R = New ADODB.Recordset
    strRecord = "SELECT CompAdd FROM Company_Information WHERE CompID=" & bytCompID & ";"
    R.Open strRecord, MainConn, adOpenStatic
    If R.RecordCount > 0 Then findCompAdd = R![CompAdd] Else findCompAdd = ""
    R.Close
    Set R = Nothing
End Function

Function DecryptIt(s As String, Crack As Integer) As String
    For i = 1 To Len(s)
        T = Mid(s, i, 1)
        Z = Z + Chr(Asc(T) - Crack)
    Next i
    DecryptIt = Z
End Function

Function EncryptIt(s As String, Crack As Integer) As String
    For i = 1 To Len(s)
        T = Mid(s, i, 1)
        Z = Z + Chr(Asc(T) + Crack)
    Next i
    EncryptIt = Z
End Function

'=== Display Multi Line into Single Line ==='
Function MultiToSingleLine(strMain As String) As String
    Dim strFinal As String, strTmp As String, intPos1 As Integer, intPos2 As Integer
   
    strTmp = strMain
    intPos1 = 1
    intPos2 = InStr(strTmp, vbCrLf)

    If intPos2 = 0 Then
        MultiToSingleLine = strTmp
        Exit Function
    End If

    If Right(strTmp, 1) <> vbCrLf Then strTmp = strTmp + vbCrLf

    strFinal = Mid$(strTmp, intPos1, intPos2 - intPos1)
    intPos1 = intPos2 + 2
    intPos2 = InStr(intPos1, strTmp, vbCrLf)

    Do While intPos2
        strFinal = strFinal + " " + Mid$(strTmp, intPos1, intPos2 - intPos1)
        intPos1 = intPos2 + 2
        intPos2 = InStr(intPos1, strTmp, vbCrLf)
    Loop
    MultiToSingleLine = strFinal
End Function

Public Sub prcMoveTab(KeyCode As Integer)
   If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Public Function fncComputerName() As String
    Dim ComputerName As String, temp As String
    ComputerName = Space(255)
    temp = GetComputerName(ComputerName, 255)
    If temp <> 0 Then
         ComputerName = Mid(Trim(ComputerName), 1, Len(Trim(ComputerName)) - 1)
    Else
        ComputerName = "#Error"
    End If
    
    fncComputerName = ComputerName
    getComName = ComputerName
End Function

''Use For System Drive Volumn No.
Public Function fncGetDriveVolumnNo() As String
    Dim serialNO As Long
    Dim res As Long
    Dim temp1 As String
    Dim temp2 As String
'    Dim GetVolSL As String
    temp1 = String$(255, Chr$(0))
    temp2 = String$(255, Chr$(0))
    res = GetVolumeInformation("c:\", temp1, Len(temp1), serialNO, 0, 0, temp2, Len(temp2))
    'MsgBox serialNO, vbOKOnly, "Volume Information"
    fncGetDriveVolumnNo = serialNO
    getComVolName = serialNO
End Function

''Public Sub Main()
''Set R = New ADODB.Recordset
''
''    isPrimaryDB = True
''        Call ConnectDatabase
''        Call prcCheckComCon
''
''    strRecord = "Select VersionNo From Version ;"
''    R.Open strRecord, MainConn, adOpenStatic
''
''    If CheckComCon = True Then
''
''        If R!VersionNo = 1.5 Then
''            frmSplash.Show
''        Else
''            MsgBox "Please update your software !!!", vbCritical, "Alert !"
''        Exit Sub
''        End If
''    End If
''    isPrimaryDB = False
''End Sub

Public Function prcCheckComCon() As Boolean

Set R1 = New ADODB.Recordset
''strRecord = "Select IDNo,ComVolumn,ComName,ComID From Sys_Configure_Status Where ((ComVolumn='" & EncryptIt(fncGetDriveVolumnNo, 11) & "') And (ComName='" & EncryptIt(fncComputerName, 11) & "')) ;"
''strRecord = "Select IDNo,ComVolumn,ComName,ComID From Sys_Configure_Status Where ((ComVolumn='" & EncryptIt(fncGetDriveVolumnNo, 11) & "') And (ComName='" & fncComputerName & "')) ;"
strRecord = "Select IDNo,ComVolumn,ComName,ComID From Sys_Configure_Status Where ComName = '" & fncComputerName & "';"
R1.Open strRecord, MainConn, adOpenStatic
If R1.RecordCount = 0 Then
    MsgBox "Computer Permission not found." & vbCrLf & "Please Contact to ICT (Software) Department.", vbCritical + vbOKOnly, "Error"
    CheckComCon = False
    Exit Function
Else
CheckComCon = True
'bytCompID = R1!ComID
getComVolID = R1!IDNo
End If
R1.Close
Set R1 = Nothing
End Function
Public Function fncNewID(ByVal Param As String)

On Error GoTo x

Set R1 = New ADODB.Recordset
If UCase(Param) = "PI" Then
  Mysql = "Select Max(PIID) AS NewID From New_PI;"
ElseIf UCase(Param) = "MRR" Then
  Mysql = "Select Max(MRRID) AS NewID From New_MRR;"
ElseIf UCase(Param) = "HEAD" Then
  Mysql = "Select Max(HEADID) AS NewID From AccountsHead;"
End If

R1.Open Mysql, MainConn, adOpenStatic
If IsNull(R1(0).Value) Then
  fncNewID = 1
Else
  fncNewID = R1(0).Value + 1
End If
R1.Close
Set R1 = Nothing

x:
If err.Number <> 0 Then MsgBox err.Description, vbCritical, AppTitle
End Function

Public Function fncNewID3(ByVal Param As String)
On Error GoTo x
Set RS = New ADODB.Recordset
If UCase(Param) = "PARTY" Then
        strRecord = "Select  Max(PartyID) AS NewID From Party;"
End If
    RS.Open strRecord, MainConn, adOpenStatic
    
    If IsNull(RS(0).Value) Then
        fncNewID3 = 1
    Else
        fncNewID3 = RS(0).Value + 1
    End If
    
    Exit Function
x:
MsgBox err.Description, vbCritical, AppTitle
End Function
Function findReportPath() As String
    Set R = New ADODB.Recordset
        strRecord = "SELECT ReportPath FROM Company_Information WHERE CompID=" & bytCompID & ";"
        R.Open strRecord, MainConn, adOpenStatic
        If R.RecordCount > 0 Then findReportPath = R![ReportPath] Else findReportPath = ""
        R.Close
    Set R = Nothing
End Function

Public Sub EnableFrame(InFrame As Frame, ByVal flag As Boolean)
  Dim Contrl As Control
'some controls don't have the Container.Name property, so instead of
'stopping the application with an error message, we ignore them.
    On Error Resume Next
'enable or disable the frame that passed as parameter.
    InFrame.Enabled = flag
'passing over all controls
    For Each Contrl In InFrame.Parent.Controls
'if the control is found in the frame
       If (Contrl.Container.Name = InFrame.Name) Then
'if the control is a frame, and it's not the frame that passed as parameter, i.e.
'other frame that found inside our frame, recursively run this sub with this frame,
'to enable or disable all the controls in it.
          If (TypeOf Contrl Is Frame) And Not (Contrl.Name = InFrame.Name) Then
             EnableFrame Contrl, flag
          Else
'enable or disable the control
             If Not (TypeOf Contrl Is Menu) Then Contrl.Enabled = flag
          End If
       End If
    Next
End Sub


