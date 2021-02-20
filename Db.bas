Attribute VB_Name = "DBase"
Option Explicit

Public AdoConn_Str As String
Public adoConn As New ADODB.Connection
Public adoHome As New ADODB.Connection
Public adoTemp As New ADODB.Connection

'// DB Open
Public Function DataBaseOpen(ByRef pAdoCon As ADODB.Connection) As Boolean

On Error GoTo Error_Result
    pAdoCon.ConnectionTimeout = 1
    pAdoCon.ConnectionString = AdoConn_Str
    pAdoCon.CursorLocation = adUseClient
    pAdoCon.Open
    DataBaseOpen = True
    Exit Function
Error_Result:
End Function

'// Db Close
Public Sub DataBaseClose(ByRef pAdoConn As ADODB.Connection)

On Error GoTo Error_Result
    pAdoConn.Close
    Set pAdoConn = Nothing
Exit Sub
Error_Result:
MsgBox Err.Description
End Sub

'// DB Open
Public Function HomeDB_Open(ByRef qAdoCon As ADODB.Connection) As Boolean

On Error GoTo Error_Result
    qAdoCon.ConnectionString = AdoHome_Str
    qAdoCon.CursorLocation = adUseClient
    qAdoCon.Open
    HomeDB_Open = True
    Exit Function

Error_Result:
    Call DataLogger(Err.Description)
End Function

'// Db Close
Public Sub HomeDB_Close(ByRef qAdoConn As ADODB.Connection)

On Error GoTo Error_Result
    qAdoConn.Close
    Set qAdoConn = Nothing
Error_Result:

End Sub

Public Function DataBaseOpenTemp(ByRef pAdoCon As ADODB.Connection, ByVal sConnStr As String) As Boolean

On Error GoTo Error_Result
    pAdoCon.ConnectionTimeout = 1
    pAdoCon.ConnectionString = sConnStr
    pAdoCon.CursorLocation = adUseClient
    pAdoCon.Open
    DataBaseOpenTemp = True
    Exit Function
Error_Result:

End Function


' DB ������ ���� ó������, ������Ʈ ���ݿ� �ִ� ���� ���๮�� �� ������ ����
' arg1: ���ڵ��
' arg2: Ŀ�ؼ�
' arg3: ��������
' arg4: ���ܱ� ����
Public Function DataBaseQuery(ByRef pRS As ADODB.Recordset, ByRef pAdoCon As ADODB.Connection, ByRef sQry As String, ByRef bGateOpen As Boolean, Optional ByVal iGateNo As Integer = -1) As Boolean
    
    On Error GoTo Err_p


'''    If (DB_Connect_F = True) Then
'        DataBaseQuery = False

        DataBaseQuery = True
        DB_Connect_F = True
        
        pRS.Open sQry, pAdoCon
        
'''    End If

    Exit Function

Err_p:
'    Call DataLogger("*[DataBase Query] " & Err.Description)
'    If (InStr(1, Err.Description, "MySQL server has gone away") > 0 Or InStr(1, Err.Description, "Lost connection to MySQL server during query") > 0 Or InStr(1, Err.Description, "t connect to MySQL server on") > 0) Then
'        If (bGateOpen = NWERR_GATE_OPEN) Then
'            Call Relay_Out(0, Glo_GateNo)
'        End If
'        Call DataLogger("QRY:" & sQry)
'        DB_Connect_F = False
'
'
'    ElseIf (InStr(1, Err.Description, "marked as crashed and should be repaired") > 0) Then
'        If (bGateOpen = NWERR_GATE_OPEN) Then
'            Call Relay_Out(0, Glo_GateNo)
'        End If
'
'        Dim sPos As Long
'        Dim ePos As Long
'        Dim sTb As String
'        sPos = InStr(1, Err.Description, "'") + 1
'        ePos = InStr(sPos, Err.Description, "'")
'        sTb = Mid(Err.Description, sPos, ePos - sPos)
'        Call DataLogger("*[DataBase Query]  " & "Trying to repair DB : " & sTb)
'        pAdoCon.Execute "repair table " & sTb & " "
'
'        DB_Connect_F = True
'        DataBaseQuery = True
'        Call DataLogger("*[DataBase Query]  " & "Repaired a DB : " & sTb)
'
'        pAdoCon.Execute sQry
'        Call DataLogger("*[DataBase ReQuery]  " & sQry)
'
'    Else
'        DB_Connect_F = False
'        DataBaseQuery = False
'        Call DataLogger("QRY:" & sQry)
'    End If
    On Error GoTo Err_Last_P
    
    Call DataLogger("*[DataBaseQuery] " & Err.Description)
    Call DebugLogger("*[DataBaseQuery] " & Err.Description)
    
    'DB ����
    If (InStr(1, Err.Description, "MySQL server has gone away") > 0 Or InStr(1, Err.Description, "Lost connection to MySQL server during query") > 0 Or InStr(1, Err.Description, "t connect to MySQL server on") > 0) Then

        pAdoCon.Close
        If DataBaseOpen(pAdoCon) Then
        End If
        
        If (pAdoCon.State = adStateOpen) Then
            Call DataLogger("[DataBaseQuery] DB Reconnection Success..!!")
            Call DebugLogger("[DataBaseQuery] DB Reconnection Success..!!")
            
            pAdoCon.Execute sQry
            Call DataLogger("[DataBaseQuery] ReQuery Success")
            Call DebugLogger("[DataBaseQuery] ReQuery Success")
            
        Else
            Call DataLogger("*[DataBaseQuery] DB Reconnection Fail..!!")
            Call DebugLogger("*[DataBaseQuery] DB Reconnection Fail..!!")
            
            Call DataLogger("*[DataBaseQuery] Lost Query : " & sQry)
            Call DebugLogger("*[DataBaseQuery] Lost Query : " & sQry)
        
            If (bGateOpen = NWERR_GATE_OPEN) Then
                Call Relay_Out(0, iGateNo)
            End If
            
            
            'Ȯ���� MySQL ��ü ���� ��쿡�� False ó���ؾ� ��
            DataBaseQuery = False
            DB_Connect_F = False
        End If
    

    '���̺� ũ����
    ElseIf (InStr(1, Err.Description, "marked as crashed and should be repaired") > 0) Then

        Dim sPos As Long
        Dim ePos As Long
        Dim sTb As String
        sPos = InStr(1, Err.Description, "'") + 1
        ePos = InStr(sPos, Err.Description, "'")
        sTb = Mid(Err.Description, sPos, ePos - sPos)
        
        Call DataLogger("[DataBaseQuery]  " & "Trying to repair DB : " & sTb)
        Call DebugLogger("[DataBaseQuery]  " & "Trying to repair DB : " & sTb)
        
        pAdoCon.Execute "repair table " & sTb & " "
        Call DataLogger("[DataBaseQuery]  " & "Repaired a DB : " & sTb)
        Call DebugLogger("[DataBaseQuery]  " & "Repaired a DB : " & sTb)
        
        pAdoCon.Execute sQry
        Call DataLogger("[DataBaseQuery] ReQuery Success")
        Call DebugLogger("[DataBaseQuery] ReQuery Success")
        
        
    'Ű �ߺ�
    ElseIf (InStr(1, Err.Description, "Duplicate entry") > 0) Then
        Call DataLogger("*[DataBaseQuery]  " & "Duplicate Errot : " & sQry)
        Call DebugLogger("*[DataBaseQuery]  " & "Duplicate Errot : " & sQry)
        
        
    '�������ϴ� MySQL ����
    Else
        '�������ϴ� MySQL �����̹Ƿ� ������ �׽�Ʈ �õ�
        pAdoCon.Close
        If DataBaseOpen(pAdoCon) Then
        End If
        
        If (pAdoCon.State = adStateOpen) Then
            Call DataLogger("[DataBaseQuery] DB Reconnection Success..!!(UnKnown)")
            Call DebugLogger("[DataBaseQuery] DB Reconnection Success..!!(UnKnown)")
            
            pAdoCon.Execute sQry
            Call DataLogger("[DataBaseQuery] ReQuery Success(UnKnown)")
            Call DebugLogger("[DataBaseQuery] ReQuery Success(UnKnown)")
            
        Else
            Call DataLogger("*[DataBaseQuery] DB Reconnection Fail..!!(UnKnown)")
            Call DebugLogger("*[DataBaseQuery] DB Reconnection Fail..!!(UnKnown)")
            
            Call DataLogger("*[DataBaseQuery] Lost Query(UnKnown) : " & sQry)
            Call DebugLogger("*[DataBaseQuery] Lost Query(UnKnown) : " & sQry)
        
            If (bGateOpen = NWERR_GATE_OPEN) Then
                Call Relay_Out(0, iGateNo)
            End If
            
            
            'Ȯ���� MySQL ��ü ���� ��쿡�� False ó���ؾ� ��
            DataBaseQuery = False
            DB_Connect_F = False
        End If

    End If
    
    Exit Function

Err_Last_P:

    'Ȯ���� MySQL ��ü ���� ��쿡�� False ó���ؾ� ��
    DB_Connect_F = False
    DataBaseQuery = False

    Call DataLogger("*[DataBaseQuery] Err_Last : " & Err.Description)
    Call DebugLogger("*[DataBaseQuery] Err_Last : " & Err.Description)
    Call DataLogger("*[DataBaseQuery] Err_Query : " & sQry)
    Call DebugLogger("*[DataBaseQuery] Err_Query : " & sQry)
    
    If (bGateOpen = NWERR_GATE_OPEN) Then
        Call Relay_Out(0, iGateNo)
    End If
    
End Function

' DB ������ ���� ó������, ������Ʈ ���ݿ� �ִ� ���� ���๮�� �� ������ ����
' arg1: Ŀ�ؼ�
' arg2: ��������
' arg3: ���ܱ� ����
Public Function DataBaseQueryExec(ByRef pAdoCon As ADODB.Connection, ByRef sQry As String, ByRef bGateOpen As Boolean, Optional ByVal iGateNo As Integer = -1) As Boolean

    On Error GoTo Err_p
    
    
    'If (DB_Connect_F = True) Then
        'DataBaseQueryExec = False
        
        DataBaseQueryExec = True
        DB_Connect_F = True
        
        pAdoCon.Execute sQry
        
    'End If
    
    Exit Function

Err_p:
    'Debug.Print Err.Description
    
On Error GoTo Err_Last_P
    
    Call DataLogger("*[DataBaseQueryExec] " & Err.Description)
    Call DebugLogger("*[DataBaseQueryExec] " & Err.Description)
    
    'DB ����
    If (InStr(1, Err.Description, "MySQL server has gone away") > 0 Or InStr(1, Err.Description, "Lost connection to MySQL server during query") > 0 Or InStr(1, Err.Description, "t connect to MySQL server on") > 0) Then

        pAdoCon.Close
        If DataBaseOpen(pAdoCon) Then
        End If
        
        If (pAdoCon.State = adStateOpen) Then
            Call DataLogger("[DataBaseQueryExec] DB Reconnection Success..!!")
            Call DebugLogger("[DataBaseQueryExec] DB Reconnection Success..!!")
            
            pAdoCon.Execute sQry
            Call DataLogger("[DataBaseQueryExec] ReQuery Success")
            Call DebugLogger("[DataBaseQueryExec] ReQuery Success")
            
        Else
            Call DataLogger("*[DataBaseQueryExec] DB Reconnection Fail..!!")
            Call DebugLogger("*[DataBaseQueryExec] DB Reconnection Fail..!!")
            
            Call DataLogger("*[DataBaseQueryExec] Lost Query : " & sQry)
            Call DebugLogger("*[DataBaseQueryExec] Lost Query : " & sQry)
        
            If (bGateOpen = NWERR_GATE_OPEN) Then
                Call Relay_Out(0, iGateNo)
            End If
            
            
            'Ȯ���� MySQL ��ü ���� ��쿡�� False ó���ؾ� ��
            DataBaseQueryExec = False
            DB_Connect_F = False
        End If
    

    '���̺� ũ����
    ElseIf (InStr(1, Err.Description, "marked as crashed and should be repaired") > 0) Then

        Dim sPos As Long
        Dim ePos As Long
        Dim sTb As String
        sPos = InStr(1, Err.Description, "'") + 1
        ePos = InStr(sPos, Err.Description, "'")
        sTb = Mid(Err.Description, sPos, ePos - sPos)
        
        Call DataLogger("[DataBaseQueryExec]  " & "Trying to repair DB : " & sTb)
        Call DebugLogger("[DataBaseQueryExec]  " & "Trying to repair DB : " & sTb)
        
        pAdoCon.Execute "repair table " & sTb & " "
        Call DataLogger("[DataBaseQueryExec]  " & "Repaired a DB : " & sTb)
        Call DebugLogger("[DataBaseQueryExec]  " & "Repaired a DB : " & sTb)
        
        pAdoCon.Execute sQry
        Call DataLogger("[DataBaseQueryExec] ReQuery Success")
        Call DebugLogger("[DataBaseQueryExec] ReQuery Success")
        
        
    'Ű �ߺ�
    ElseIf (InStr(1, Err.Description, "Duplicate entry") > 0) Then
        Call DataLogger("*[DataBaseQueryExec]  " & "Duplicate Errot : " & sQry)
        Call DebugLogger("*[DataBaseQueryExec]  " & "Duplicate Errot : " & sQry)
        
        
    '�������ϴ� MySQL ����
    Else
        '�������ϴ� MySQL �����̹Ƿ� ������ �׽�Ʈ �õ�
        pAdoCon.Close
        If DataBaseOpen(pAdoCon) Then
        End If
        
        If (pAdoCon.State = adStateOpen) Then
            Call DataLogger("[DataBaseQueryExec] DB Reconnection Success..!!(UnKnown):" & Err.Description)
            Call DebugLogger("[DataBaseQueryExec] DB Reconnection Success..!!(UnKnown):" & Err.Description)
            
            pAdoCon.Execute sQry
            Call DataLogger("[DataBaseQueryExec] ReQuery Success(UnKnown)")
            Call DebugLogger("[DataBaseQueryExec] ReQuery Success(UnKnown)")
            
        Else
            Call DataLogger("*[DataBaseQueryExec] DB Reconnection Fail..!!:" & Err.Description)
            Call DebugLogger("*[DataBaseQueryExec] DB Reconnection Fail..!!:" & Err.Description)
            
            Call DataLogger("*[DataBaseQueryExec] Lost Query(UnKnown) : " & sQry)
            Call DebugLogger("*[DataBaseQueryExec] Lost Query(UnKnown) : " & sQry)
        
            If (bGateOpen = NWERR_GATE_OPEN) Then
                Call Relay_Out(0, iGateNo)
            End If
            
            
            'Ȯ���� MySQL ��ü ���� ��쿡�� False ó���ؾ� ��
            DataBaseQueryExec = False
            DB_Connect_F = False
        End If

    End If
    
    Exit Function

Err_Last_P:

    'Ȯ���� MySQL ��ü ���� ��쿡�� False ó���ؾ� ��
    DB_Connect_F = False
    DataBaseQueryExec = False

    Call DataLogger("*[DataBaseQueryExec] Err_Last : " & Err.Description)
    Call DebugLogger("*[DataBaseQueryExec] Err_Last : " & Err.Description)
    Call DataLogger("*[DataBaseQueryExec] Err_Query : " & sQry)
    Call DebugLogger("*[DataBaseQueryExec] Err_Query : " & sQry)
    
    If (bGateOpen = NWERR_GATE_OPEN) Then
        Call Relay_Out(0, iGateNo)
    End If
    
End Function

Public Sub DBaseCheck()

Dim a As Boolean
    Dim rs As ADODB.Recordset
    Dim qry As String
    Dim str As String
    
On Error GoTo Error_Result

    Set rs = New ADODB.Recordset
    
    On Error GoTo Error_Result
    
    qry = "show tables"
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    Set rs = Nothing
    
    
    Exit Sub
    
Error_Result:


On Error GoTo Err_p

    Call DebugLogger("[DBaseCheck] DB Connection Lost : " & Err.Description)

    
    If (adoConn.State = adStateOpen) Then
        adoConn.Close
    End If
    
    If DataBaseOpen(adoConn) Then
    End If
    
    If (adoConn.State = adStateOpen) Then
        Call DataLogger("[DBaseCheck] DB Reconnection Success..!!")
        Call DebugLogger("[DBaseCheck] DB Reconnection Success..!!")
    Else
        Call DataLogger("[DBaseCheck] DB Reconnection Fail..!!")
        Call DebugLogger("[DBaseCheck] DB Reconnection Fail..!!")
    End If
    
    Set rs = Nothing
    
    Exit Sub

Err_p:

    Call DebugLogger("[DBaseCheck] Error : " & Err.Description)

End Sub

Public Sub MakeCSV(lv As ListView, CSVname As String)

    Dim intFileNum As Integer
    Dim ecdata As New ExcelFile
    Dim i, j As Long
    Dim tmpHeader As String
    Dim tmpRS As String

    tmpHeader = ""

    For i = 1 To lv.ColumnHeaders.Count
        If i = 1 Then
            tmpHeader = Trim(lv.ColumnHeaders.Item(1).text)
        Else
            tmpHeader = tmpHeader & "," & Trim(lv.ColumnHeaders.Item(i).text)
        End If
    Next i

    intFileNum = FreeFile()
    Open CSVname & ".CSV" For Append As #intFileNum
    Print #intFileNum, tmpHeader

    For i = 1 To lv.ListItems.Count
        For j = 1 To lv.ColumnHeaders.Count
            If j = 1 Then
                tmpRS = tmpRS & lv.ListItems(i).text
            Else
                tmpRS = tmpRS & "," & lv.ListItems(i).SubItems(j - 1)
            End If
        Next j
        'Debug.Print tmpRS
        Print #intFileNum, tmpRS
        tmpRS = ""
    Next i

    Close #intFileNum
    'MsgBox "������ �Ϸ�Ǿ����ϴ�."

End Sub


'������Ʈ > ���� > Microsoft Scripting Runtime üũ�ؾ� ��.
Public Function IsFile(strFile As String) As Boolean

    On Error GoTo ERR_RTN
    
    Dim fso As FileSystemObject
    Dim strLogMsg As String
    Dim strLogType As String
    Dim strFileName As String
    Dim PauseTime As Single
    Dim start  As Single
    
    
    PauseTime = 1
    start = Timer
    Do While Timer < start + PauseTime
        If (Timer < start) Then
            start = start - 86400
        End If
        
        If InStr(1, strFile, "\") > 0 Then
            strFileName = Mid(strFile, InStrRev(strFile, "\", -1) + 1)
        End If
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FileExists(strFile) Then
            IsFile = True
            Exit Function
        End If
    Loop
    
    IsFile = False
    
    Set fso = Nothing
    Exit Function
    
ERR_RTN:
    
    IsFile = False
    If Not fso Is Nothing Then Set fso = Nothing
    
End Function


Public Sub LISTBOX_PutString(ByVal lst As ListBox, ByVal msg As String)
    lst.AddItem Format(Now, "HH:NN:SS") & msg, 0
    If (lst.ListCount > MAX_LISTBOX_LINE) Then
        lst.RemoveItem (lst.ListCount - 1)  '��������� �׸� ����
        
    End If
End Sub



Public Sub DB_CFG_Init(ByVal sCategory As String)

    Dim rs As Recordset
    Dim bQryResult As Boolean
    
    ' ����
    If (sCategory = "����") Then
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'SOUND_YN' ", False): Glo_SOUND_YN = rs!Content: Set rs = Nothing
        
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane1_NoReg' ", False): Glo_SND_Lane1_Guest_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane1_NoRec' ", False): Glo_SND_Lane1_NoRec_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane1_BlackList' ", False): Glo_SND_Lane1_BlackList_YN = rs!Content: Set rs = Nothing
        
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane2_NoReg' ", False): Glo_SND_Lane2_Guest_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane2_NoRec' ", False): Glo_SND_Lane2_NoRec_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane2_BlackList' ", False): Glo_SND_Lane2_BlackList_YN = rs!Content: Set rs = Nothing
        
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane3_NoReg' ", False): Glo_SND_Lane3_Guest_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane3_NoRec' ", False): Glo_SND_Lane3_NoRec_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane3_BlackList' ", False): Glo_SND_Lane3_BlackList_YN = rs!Content: Set rs = Nothing
        
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane4_NoReg' ", False): Glo_SND_Lane5_Guest_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane4_NoRec' ", False): Glo_SND_Lane4_NoRec_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane4_BlackList' ", False): Glo_SND_Lane4_BlackList_YN = rs!Content: Set rs = Nothing
        
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane5_NoReg' ", False): Glo_SND_Lane5_Guest_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane5_NoRec' ", False): Glo_SND_Lane5_NoRec_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane5_BlackList' ", False): Glo_SND_Lane5_BlackList_YN = rs!Content: Set rs = Nothing
        
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane6_NoReg' ", False): Glo_SND_Lane6_Guest_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane6_NoRec' ", False): Glo_SND_Lane6_NoRec_YN = rs!Content: Set rs = Nothing
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'Lane6_BlackList' ", False): Glo_SND_Lane6_BlackList_YN = rs!Content: Set rs = Nothing
    End If
End Sub



Public Sub UnloadForms(ByVal own As Form)
    Dim frm As Form
    For Each frm In Forms
        If (frm.name <> own.name And frm.name <> "FrmTcpServer") Then
            Unload frm
            Set frm = Nothing
        End If
    Next
    'Debug.Print Forms.Count
End Sub




Public Function Able_WebDC() As Boolean
    Dim rs As Recordset
    Dim qry As String

    Able_WebDC = False
    
    On Error Resume Next

    Set rs = New ADODB.Recordset
    'qry = "SELECT Content FROM tb_config WHERE (NAME = 'PCWebDC' AND CONTENT = 'Y') OR (NAME = 'AppWebDC' AND CONTENT = 'Y') "
    qry = "SELECT Content FROM tb_config WHERE (NAME = 'WebDC' AND CONTENT = 'Y') "
    
    rs.Open qry, adoConn
    
    If (Not (rs.EOF)) Then
        Able_WebDC = True
    End If
    
    Set rs = Nothing
End Function



