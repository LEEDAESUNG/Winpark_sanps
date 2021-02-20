Attribute VB_Name = "LPR"
Option Explicit
'Public Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000


Public Function Get_Process(fileName As String) As Boolean
Dim cb As Long
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim i As Long
Dim tmp As String

cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop

NumElements = cbNeeded / 4
For i = 1 To NumElements
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
    If hProcess <> 0 Then
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
        If lRet <> 0 Then
            ModuleName = Space(MAX_PATH)
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
            
            tmp = ModuleName
            tmp = Mid(tmp, 1, InStr(tmp, "exe") + 2)
            
            If (InStr(tmp, fileName) <> 0) Then
                lRet = CloseHandle(hProcess)
                Get_Process = True
                Exit Function
            Else
                
            End If
            
            'tmp = Left(ModuleName, lRet)
            'If (Right(tmp, 11) = filename) Then
            '    Get_Process = True
            '    Exit Function
            'End If
        
        End If
    End If
    lRet = CloseHandle(hProcess)
Next
Get_Process = False

End Function

Public Sub Data_ReSearch(size As Integer)
End Sub

Public Sub Data_ReSearch_Unload()
End Sub

Public Sub BackLPR_Proc(ByVal sGateNo As String, ByVal sBackCarno As String, ByVal sPassDate As String, ByVal sImage As String)
    
    '�Ĺ�ī�޶� ����(Lane��ȣ 6~11)
    Dim sFrontCarno As String
    Dim sFrontPassDate As String
    Dim sFrontGateNo As String
    Dim sBackInout As String
    Dim sMODEL, sGUBUN, sName, sPHONE, sDEPT As String
    Dim sCLASS, sSDT, sEDT, sGATE, sInOut As String
    Dim sYN, sResult, sPassINOUT, sLPR_IP As String
    Dim sQry As String

On Error GoTo Err_p

    Select Case sGateNo
        Case 6
            If (Glo_Lane1_Back_YN <> "Y") Then Exit Sub
            If (LANE1_Inout = "�Ա�") Then sPassINOUT = "IN" Else sPassINOUT = "OUT"
            sFrontGateNo = "0"
            sFrontCarno = Glo_Lane1_Front_CarNo
            sFrontPassDate = Glo_Lane1_Front_PassDate
            'sLPR_IP = ""
        Case 7
            If (Glo_Lane2_Back_YN <> "Y") Then Exit Sub
            If (LANE2_Inout = "�Ա�") Then sPassINOUT = "IN" Else sPassINOUT = "OUT"
            sFrontGateNo = "1"
            sFrontCarno = Glo_Lane2_Front_CarNo
            sFrontPassDate = Glo_Lane2_Front_PassDate
        Case 8
            If (Glo_Lane3_Back_YN <> "Y") Then Exit Sub
            If (LANE3_Inout = "�Ա�") Then sPassINOUT = "IN" Else sPassINOUT = "OUT"
            sFrontGateNo = "2"
            sFrontCarno = Glo_Lane3_Front_CarNo
            sFrontPassDate = Glo_Lane3_Front_PassDate
        Case 9
            If (Glo_Lane4_Back_YN <> "Y") Then Exit Sub
            If (LANE4_Inout = "�Ա�") Then sPassINOUT = "IN" Else sPassINOUT = "OUT"
            sFrontGateNo = "3"
            sFrontCarno = Glo_Lane4_Front_CarNo
            sFrontPassDate = Glo_Lane4_Front_PassDate
        Case 10
            If (Glo_Lane5_Back_YN <> "Y") Then Exit Sub
            If (LANE5_Inout = "�Ա�") Then sPassINOUT = "IN" Else sPassINOUT = "OUT"
            sFrontGateNo = "4"
            sFrontCarno = Glo_Lane5_Front_CarNo
            sFrontPassDate = Glo_Lane5_Front_PassDate
        Case 11
            If (Glo_Lane6_Back_YN <> "Y") Then Exit Sub
            If (LANE6_Inout = "�Ա�") Then sPassINOUT = "IN" Else sPassINOUT = "OUT"
            sFrontGateNo = "5"
            sFrontCarno = Glo_Lane6_Front_CarNo
            sFrontPassDate = Glo_Lane6_Front_PassDate
    End Select
    
    If (sGateNo >= 6 And sGateNo < 12) Then
        
        adoConn.Execute "INSERT INTO tb_inout VALUES ('" & sBackCarno & "', '" & sBackCarno & "', '', '', '', '', '', '', '', '', '" & sGateNo & "', '" & sPassINOUT & "', '" & sPassDate & "', 'N', '�̵������', '" & sImage & "', '" & sLPR_IP & "', 0)"
        
        
        '���� ������ȣ�� "�νĽ���" �Ǵ� 6�ڸ� �����϶�,
        If ((sFrontCarno = "�νĽ���" Or LenH(sFrontCarno) <= 6)) Then
            
            '�Ĺ� ������ȣ �����϶�
            If ((sBackCarno <> "�νĽ���" And LenH(sBackCarno) >= 8)) Then

                sQry = "SELECT * FROM tb_reg WHERE CAR_NO = '" & sBackCarno & "' "
                Set rs = New ADODB.Recordset
                rs.Open sQry, adoConn
                
                If Not (rs.EOF) Then
                    sMODEL = "" & rs!CAR_MODEL:         sGUBUN = "" & rs!CAR_GUBUN:     sName = "" & rs!DRIVER_NAME:    sPHONE = "" & rs!DRIVER_PHONE:  sDEPT = "" & rs!DRIVER_DEPT:
                    sCLASS = "" & rs!DRIVER_CLASS:      sSDT = "" & rs!START_DATE:      sEDT = "" & rs!END_DATE:        sGATE = sGateNo:                sInOut = sPassINOUT:
                    sPassDate = sPassDate:              sYN = "N":                      sResult = "��������"
                Else
                    sMODEL = "":                        sGUBUN = "":                    sName = "":                     sPHONE = "":                    sDEPT = "":
                    sCLASS = "":                        sSDT = "":                      sEDT = "":                      sGATE = sGateNo:                sInOut = sPassINOUT:
                    sPassDate = sPassDate:              sYN = "N":                      sResult = "�̵������"
                End If
                Set rs = Nothing

                'adoConn.Execute "UPDATE tb_now   SET CAR_NO = '" & sBackCarno & "' WHERE PASS_DATE ='" & sFrontPassDate & "' AND PASS_GATE = '" & sFrontGateNo & "' "
                adoConn.Execute "UPDATE tb_inout SET CAR_NO = '" & sBackCarno & "' WHERE PASS_DATE ='" & sFrontPassDate & "' AND PASS_GATE = '" & sFrontGateNo & "' "
                'adoConn.Execute "INSERT INTO tb_inout VALUES ('" & sBackCarno & "', '" & sBackCarno & "', '" & sMODEL & "', '" & sGUBUN & "', '" & sNAME & "', '" & sPHONE & "', '" & sDEPT & "', '" & sCLASS & "', '" & sSDT & "', '" & sEDT & "', '" & sGATE & "', '" & sInOut & "', '" & sPassDate & "', '" & sYN & "', '" & sResult & "', '" & sImage & "', '" & sLPR_IP & "', 0)"
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�Ĺ�������ü', 'HOST','" & sFrontCarno & " -> " & sBackCarno & " ��ü',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                
                Call DataLogger("���� ������ȣ (" & sFrontCarno & "), �Ĺ� ������ȣ (" & sBackCarno & ")" & " ==> ������ȣ ��ü OK")
                
            Else
                Call DataLogger("���� ������ȣ (" & sFrontCarno & "), �Ĺ� ������ȣ (" & sBackCarno & ")" & " ==> ������ȣ ��ü PASS")
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�Ĺ�������ü', 'HOST','" & sFrontCarno & " -> " & sBackCarno & " ��ü PASS',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            End If
                
        Else
            Call DataLogger("���� ������ȣ (" & sFrontCarno & ") ==> �Ĺ� ������ȣ(" & sBackCarno & ") PASS")
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�Ĺ�������ü', 'HOST','" & sFrontCarno & " -> " & sBackCarno & " ��ü PASS',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            
        End If
        
        '���� ������ȣ �ʱ�ȭ
        Select Case sGateNo
            Case 6
                Glo_Lane1_Front_CarNo = ""
                Glo_Lane1_Front_PassDate = ""
            Case 7
                Glo_Lane2_Front_CarNo = ""
                Glo_Lane2_Front_PassDate = ""
            Case 8
                Glo_Lane3_Front_CarNo = ""
                Glo_Lane3_Front_PassDate = ""
            Case 9
                Glo_Lane4_Front_CarNo = ""
                Glo_Lane4_Front_PassDate = ""
            Case 10
                Glo_Lane5_Front_CarNo = ""
                Glo_Lane5_Front_PassDate = ""
            Case 11
                Glo_Lane6_Front_CarNo = ""
                Glo_Lane6_Front_PassDate = ""
        End Select
    End If
    
    Exit Sub
Err_p:
    Call DataLogger("[BackLPR_Proc] " & Err.Description)
End Sub
Public Sub LPRIn_Proc(ByVal carnum As String, ByVal ImageFile As String, ByVal sLPR_IP As String, ByVal sLane_INOUT As String, ByVal sFreePass As String, ByVal sBlackList As String, ByVal sNoRecOpen As String, ByVal sTaxiPass As String, ByVal iGateNo As Integer, ByVal sPassDate As String, ByRef sRef_GateOpen As String, ByRef sRef_GateStat As String, ByRef stSound As structSound, ByVal sNoWork As String, ByRef stEmerg As structEmerg)
'Public Sub LPRIn_Proc(ByVal carnum As String, ByVal ImageFile As String, ByVal stLPR As structLPR, ByRef sRef_GateOpen As String, ByRef sRef_GateStat As String)

    
    Dim Car_Num_Str As String
    Dim Car_i As Integer
    Dim i As Integer
    Dim Save_CarNum As String
    Dim RecStat As String
    Dim qry As String
    Dim Ret As Integer
    Dim Check_Flag As Boolean
    Dim Rec_CarNo As String
    Dim Proc_CarNo As String
    Dim HomeNet_Str As String

    Dim iRotatRes As Integer '�����
    Dim iDay As Integer '���� ��¥
    Dim sWeekday As String '���� ����
    Dim bWeekday As String
    Dim iCarEndNo As Integer
    Dim bQryResult As Boolean

    ' �߰�
    Dim sCAR_MODEL As String
    Dim sCAR_GUBUN As String
    Dim sDRIVER_NAME As String
    Dim sDRIVER_PHONE As String
    Dim sDRIVER_DEPT As String
    Dim sDRIVER_CLASS As String
    Dim sSTART_DATE As String
    Dim sEND_DATE As String
    Dim sPASS_GATE As String
    Dim sPASS_INOUT As String
    'Dim sPASS_DATE As String
    Dim sPass_YN As String
    Dim sPASS_RESULT As String
    Dim sPassLane As String
    Dim sGUESTREG_ID As String
    
    Dim sFee_Carno As String
    Dim sGateOPen_YN As String
    Dim sEmergency_Print As String
    Dim sCarReg_Kind As String
    
    Dim sWebDCResult As String
    
    '�湮���� ����
    Dim sGuestQry As String
    Dim sGuestRegAdminQry As String
    Dim nMaxParkCount As Long
    Dim nNowParkCount As Long
                                
    Dim sGuestPassDate As String
    
    ' �߰�

On Error GoTo Err_Proc
    
    sGuestPassDate = Left(sPassDate, 19)
    
    sGateOPen_YN = "N"
    sFee_Carno = ""
    sEmergency_Print = ""
    
    MissMatch_F = False
    Rec_CarNo = carnum
    Proc_CarNo = carnum
    iCarEndNo = Val(Right(carnum, 1))
    ImageFile = Slash_Conv(ImageFile)
    '����ī�޶� ���� ����
    If (iGateNo >= 0 And iGateNo < 6) Then
        Select Case iGateNo
            Case 0
                Glo_Lane1_Front_CarNo = carnum
                Glo_Lane1_Front_PassDate = sPassDate
            Case 1
                Glo_Lane2_Front_CarNo = carnum
                Glo_Lane2_Front_PassDate = sPassDate
            Case 2
                Glo_Lane3_Front_CarNo = carnum
                Glo_Lane3_Front_PassDate = sPassDate
            Case 3
                Glo_Lane4_Front_CarNo = carnum
                Glo_Lane4_Front_PassDate = sPassDate
            Case 4
                Glo_Lane5_Front_CarNo = carnum
                Glo_Lane5_Front_PassDate = sPassDate
            Case 5
                Glo_Lane6_Front_CarNo = carnum
                Glo_Lane6_Front_PassDate = sPassDate
        End Select
    End If
    
    '�Ĺ�ī�޶� ó��(Lane��ȣ 6~11)
    If (iGateNo >= 6 And iGateNo < 12) Then
        Call BackLPR_Proc(iGateNo, carnum, sPassDate, ImageFile)
        Exit Sub
    End If


With FrmG4Mini
    If (carnum = "�νĽ���") Then
        RecStat = "���ν�"
        GoTo No_Data
    End If


    Set rs = New ADODB.Recordset
    qry = "SELECT * FROM tb_reg WHERE CAR_NO = '" & carnum & "'"
    'rs.Open Qry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, qry, NWERR_GATE_OPEN, iGateNo)
    If (bQryResult = False) Then
        DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
        Exit Sub
    End If

    If Not (rs.EOF) Then

        
        RecStat = "�����ν�"
    
Seek_DB:
        Proc_CarNo = rs!CAR_NO
        
        
        sCAR_MODEL = "" & rs!CAR_MODEL
        sCAR_GUBUN = "" & rs!CAR_GUBUN
        sDRIVER_NAME = "" & rs!DRIVER_NAME
        sDRIVER_PHONE = "" & rs!DRIVER_PHONE
        sDRIVER_DEPT = "" & rs!DRIVER_DEPT
        sDRIVER_CLASS = "" & rs!DRIVER_CLASS
        sSTART_DATE = Format(rs!START_DATE, "yyyy-mm-dd hh:nn:ss")
        sEND_DATE = Format(rs!END_DATE, "yyyy-mm-dd hh:nn:ss")
        sPASS_GATE = iGateNo
        
        
        Check_Flag = True '�ϳ��� ������ ����� false �� �������� �����.



        '�Ⱓ��
        'If (rs!Start_Date <= Format(Now, "yyyymmdd") And rs!End_Date >= Format(Now, "yyyymmdd")) Then
        If (Format(rs!START_DATE, "yyyymmddhhnnss") <= Format(Now, "yyyymmddhhnnss") And Format(rs!END_DATE, "yyyymmddhhnnss") >= Format(Now, "yyyymmddhhnnss")) Then '����Ͼ� ����(jwt_sanps)
        Else
            Check_Flag = False
            '��ϱⰣ�� ���� ����ó��
            Glo_Disp1 = rs!CAR_NO
            Glo_Disp2 = "�Ⱓ����"
            Glo_Gate = "CLOSE"
            sRef_GateStat = "�Ⱓ����"
            sRef_GateOpen = "CLOSE"
        End If
        '�������� ����
        If (rs!CAR_GUBUN = "��������") Then
            '���ѽÿ��� ���ܱ� & ������ ǥ��
            '�������ѿ� ���� ����ó��
            If (sBlackList = "Y") Then
                Check_Flag = False
                Glo_Disp1 = rs!CAR_NO
                Glo_Disp2 = "��������"
                Glo_Gate = "CLOSE"
                sRef_GateStat = "��������"
                sRef_GateOpen = "CLOSE"
            Else
'                Glo_Disp2 = "�������"
'                Glo_Gate = "OPEN"
'                sRef_GateStat = "�������"
'                sRef_GateOpen = "OPEN"
            End If
        Else
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ���� ����(2����, 5����, 10����)
        If (Glo_ROTATION <> "������") Then      '��������� ��������
            If (rs!Rotation = "Y") Then         '����� ���������� ���ؼ� ���� ����

                iDay = Val(Format(Now, "d"))            '���� ����
                If (Glo_ROTATION = "2����") Then
                        If ((Int(iCarEndNo Mod 2)) = (Int(iDay Mod 2))) Then
                        Else
                            Check_Flag = False
                        End If
                ElseIf (Glo_ROTATION = "5����") Then
                        sWeekday = Format(Now, "dddd")      '�������
                        Select Case sWeekday
                        Case "Monday"
                            If (iCarEndNo = 1 Or iCarEndNo = 6) Then
                                Check_Flag = False
                            End If
                        Case "Tuesday"
                            If (iCarEndNo = 2 Or iCarEndNo = 7) Then
                                Check_Flag = False
                            End If
                        Case "Wednesday"
                            If (iCarEndNo = 3 Or iCarEndNo = 8) Then
                                Check_Flag = False
                            End If
                        Case "Thursday"
                            If (iCarEndNo = 4 Or iCarEndNo = 9) Then
                                Check_Flag = False
                            End If
                        Case "Friday"
                            If (iCarEndNo = 5 Or iCarEndNo = 0) Then
                                Check_Flag = False
                            End If
                    End Select
                ElseIf (Glo_ROTATION = "10����") Then
                        If (iCarEndNo = (iDay Mod 10)) Then
                            Check_Flag = False
                        End If
                End If

                If (Check_Flag = False) Then            '����  ����
                    Glo_Disp1 = rs!CAR_NO
                    Glo_Disp2 = "��������"
                    Glo_Gate = "CLOSE"
                    sRef_GateStat = "��������"
                    sRef_GateOpen = "CLOSE"
                Else
                End If
            Else
            End If
        Else
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
            '���� ���Ͽ��� ����(�������� �ƴҶ� ����)
            If (Glo_WEEK_YN = "Y") Then ' ���� ���Ͽ��� ����
                'If (Glo_Disp2 <> "��������") Then '�������� �ƴ�(���������̶��, ���Ͽ��� üũ�ʿ� ���ǹ�)
    
                    sWeekday = Format(Now, "dddd")      '�������
                    bWeekday = False
                
                    Select Case sWeekday
                        Case "Monday"
                            If (rs!WEEK1 = "Y") Then '�ش������� ���� ���� ���
                                bWeekday = True
                            End If
                        Case "Tuesday"
                            If (rs!WEEK2 = "Y") Then
                                bWeekday = True
                            End If
                        Case "Wednesday"
                            If (rs!WEEK3 = "Y") Then
                                bWeekday = True
                            End If
                        Case "Thursday"
                            If (rs!WEEK4 = "Y") Then
                                bWeekday = True
                            End If
                        Case "Friday"
                            If (rs!WEEK5 = "Y") Then
                                bWeekday = True
                            End If
                        Case "Saturday"
                            If (rs!WEEK6 = "Y") Then
                                bWeekday = True
                            End If
                        Case "Sunday"
                            If (rs!WEEK7 = "Y") Then
                                bWeekday = True
                            End If
                    End Select
        
                    If (bWeekday = False) Then
                        Check_Flag = False
                        Glo_Disp1 = rs!CAR_NO
                        Glo_Disp2 = "��������"
                        Glo_Gate = "CLOSE"
                        sRef_GateStat = "��������"
                        sRef_GateOpen = "CLOSE"
                    End If
                'End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If



        '���� �� ��� ó��
        If (Check_Flag = True) Then
            Glo_Disp1 = rs!CAR_NO
            Glo_Disp2 = "�������"
            Glo_Gate = "OPEN"
            sRef_GateStat = "�������"
            sRef_GateOpen = "OPEN"
            
            'ImageFile = Slash_Conv(ImageFile)
        
            
    
            '������ ���� ���
            If iGateNo = 0 Then
                sPassLane = rs!LANE1
            ElseIf iGateNo = 1 Then
                sPassLane = rs!LANE2
            ElseIf iGateNo = 2 Then
                sPassLane = rs!LANE3
            ElseIf iGateNo = 3 Then
                sPassLane = rs!LANE4
            ElseIf iGateNo = 4 Then
                sPassLane = rs!LANE5
            ElseIf iGateNo = 5 Then
                sPassLane = rs!LANE6
            Else
                sPassLane = "N"
            End If
            
            
            If (sPassLane = "Y") Then
                'INOUT ó�� ��������
'                If (Glo_Display = "������") Then
'                    Call GL_Emergency("[�������]", rs!CAR_NO, 0, 30, 10, 1, 2, 1, iGateNo)
'                ElseIf Glo_Display = "FND" Then
'                    Call FND_Display(Right(rs!CAR_NO, 4), iGateNo)
'                End If

                
                sEmergency_Print = "�������"
                
                If (sLane_INOUT = "�Ա�") Then
                    'Call Relay_Out(0, iGateNo)
                    sGateOPen_YN = "Y"
                    sPASS_INOUT = "IN":                    sPass_YN = "Y":                    sPASS_RESULT = "��������"
                    
                    '�����뺸
                    If (HomeNet_YN = "Y" And MissMatch_F = False) Or (HomeNet_YN = "Y" And MissMatch_F = True And MissMatch_HomeNet_YN = "Y") Then
                        If (IsNumeric(rs!DRIVER_DEPT) = True) And (IsNumeric(rs!DRIVER_CLASS) = True) And (rs!DAY_ROTATION_YN = "����") Then
                            Call SendHomenet(rs!DRIVER_DEPT, rs!DRIVER_CLASS, Proc_CarNo)
                        End If
                    End If
                Else
                   'INOUT ó�� ��������
                    'Call Relay_Out(0, iGateNo)
                    sGateOPen_YN = "Y"
                    sPASS_INOUT = "OUT":                        sPass_YN = "Y":                        sPASS_RESULT = "��������"
                End If
            Else
                If (sLane_INOUT = "�Ա�") Then
                    sPASS_INOUT = "IN":                         sPass_YN = "N":                        sPASS_RESULT = "������������"
                Else
                    sPASS_INOUT = "OUT":                        sPass_YN = "N":                        sPASS_RESULT = "������������"
                End If
                
'                If (Glo_Display = "������") Then
'                    Call GL_Emergency("[��������]", rs!CAR_NO, 0, 30, 10, 1, 2, 0, iGateNo)
'                ElseIf Glo_Display = "FND" Then
'                    Call FND_Display(Right(rs!CAR_NO, 4), iGateNo)
'                End If
                sEmergency_Print = "��������"

            End If

        Else 'REG ���� ó��
            'ImageFile = Slash_Conv(ImageFile)
            Select Case Trim(Glo_Disp2)
                Case "�Ⱓ����"

'                    If (Glo_Display = "������") Then
'                        Call GL_Emergency("[�Ⱓ����]", rs!CAR_NO, 0, 30, 10, 1, 2, 0, iGateNo)
'                    ElseIf Glo_Display = "FND" Then
'                        Call FND_Display(Right(rs!CAR_NO, 4), iGateNo)
'                    End If
                    'sEmergency_Print = "[�Ⱓ����]"
                    sEmergency_Print = "�Ⱓ����"
                    
                    If (sLane_INOUT = "�Ա�") Then
                        sPASS_INOUT = "IN":                         sPass_YN = "N":                         sPASS_RESULT = "�Ⱓ��������"
                    Else
                        sPASS_INOUT = "OUT":                        sPass_YN = "N":                         sPASS_RESULT = "�Ⱓ��������"
                    End If
                    
                Case "��������"
                    sEmergency_Print = "��������"
                    
                    'If (sBlackList = "N") And (sFreePass = "N") Then
                    If (sBlackList = "N") Then
                        'Call Relay_Out(0, iGateNo)
                        sGateOPen_YN = "Y"
                    End If
                    
                    If (sLane_INOUT = "�Ա�") Then
''''                        Call Sound_Out("������������.wav")
                        sPASS_INOUT = "IN":                        sPass_YN = sGateOPen_YN:                          sPASS_RESULT = "������������"
                    Else
                        sPASS_INOUT = "OUT":                       sPass_YN = sGateOPen_YN:                        sPASS_RESULT = "������������"
                    End If

'                    If (Glo_Display = "������") Then
'                        Call GL_Emergency("[��������]", rs!CAR_NO, 0, 30, 10, 1, 2, 0, iGateNo)
'                    ElseIf Glo_Display = "FND" Then
'                        Call FND_Display(Right(rs!CAR_NO, 4), iGateNo)
'                    End If
                    

                    
                Case "��������"
'                    If (Glo_Display = "������") Then
'                        Call GL_Emergency("[��������]", rs!CAR_NO, 0, 30, 10, 1, 2, 0, iGateNo)
'                    ElseIf Glo_Display = "FND" Then
'                        Call FND_Display(Right(rs!CAR_NO, 4), iGateNo)
'                    End If
                    sEmergency_Print = "��������"
                    
                    If (sLane_INOUT = "�Ա�") Then
                            sPASS_INOUT = "IN":                            sPass_YN = "N":                            sPASS_RESULT = "������������"
                    Else
                            sPASS_INOUT = "OUT":                            sPass_YN = "N":                            sPASS_RESULT = "������������"
                    End If
                Case "��������"
                    
'                    If (Glo_Display = "������") Then
'                        Call GL_Emergency("��������", rs!CAR_NO, 0, 30, 10, 1, 2, 0, iGateNo)
'                    ElseIf Glo_Display = "FND" Then
'                        Call FND_Display(Right(rs!CAR_NO, 4), iGateNo)
'                    End If
                    sEmergency_Print = "��������"
                    
                    If (sLane_INOUT = "�Ա�") Then
                            sPASS_INOUT = "IN":                            sPass_YN = "N":                            sPASS_RESULT = "������������"
                    Else
                            sPASS_INOUT = "OUT":                            sPass_YN = "N":                            sPASS_RESULT = "������������"
                    End If
            End Select
        End If  '���� �� �� ���� �Ǵ� ���� ����
    Else
        RecStat = "���ν�"
        '==========================================================================================================================================================================
        '�ѱ� ���͸�
        If (MissMatch_YN = "Y") Then
            Select Case LenH(carnum)
                Case 8
                    qry = "SELECT * FROM tb_reg WHERE CAR_NO Like '" & Left(carnum, 2) & "_" & Right(carnum, 4) & "'"
                Case 9
                    qry = "SELECT * FROM tb_reg WHERE CAR_NO Like '" & Left(carnum, 3) & "_" & Right(carnum, 4) & "'"
                Case 11
                    qry = "SELECT * FROM tb_reg WHERE CAR_NO Like '" & Left(carnum, 3) & "_" & Right(carnum, 4) & "'"
                Case 12
                    qry = "SELECT * FROM tb_reg WHERE CAR_NO Like '" & Left(carnum, 4) & "_" & Right(carnum, 4) & "'"
                
            End Select
            Set rs = New ADODB.Recordset
            'rs.Open Qry, adoConn
            bQryResult = DataBaseQuery(rs, adoConn, qry, NWERR_GATE_OPEN, iGateNo)
            If (bQryResult = False) Then
                DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
                Exit Sub
            End If
            
            If (rs.EOF) Then
            Else
                Call DataLogger("[LPRIn_Proc] �ѱ����͸� : " & carnum & "    ��ü��ȣ : " & rs!CAR_NO)
                MissMatch_F = True
                GoTo Seek_DB
            End If
        End If
        
        '��Ÿ�� ��ġ
            Set rs = New ADODB.Recordset
            qry = "SELECT * FROM tb_reg WHERE ETC Like '%" & carnum & "%'"
            'rs.Open Qry, adoConn
            bQryResult = DataBaseQuery(rs, adoConn, qry, NWERR_GATE_OPEN, iGateNo)
            If (bQryResult = False) Then
                DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
                Exit Sub
            End If
            
            If (rs.EOF) Then
            Else
                Call DataLogger("[LPRIn_Proc] ��Ÿ ���͸� : " & carnum & "    ��ü��ȣ : " & rs!CAR_NO)
                GoTo Seek_DB
            End If
        '==========================================================================================================================================================================
        'tb_reg�� ������ ���ٸ� ���ν� or ���ν�. ������ ���ν� ó���������Ƿ� ���⼭���� ���ν����� ó���ؾ� ��
        Select Case RecStat
           Case "���ν�"
No_Data:
                'ImageFile = Slash_Conv(ImageFile)
                
                If ((sNoRecOpen = "Y") And (sFreePass = "Y")) Then
                   'Call Relay_Out(0, iGateNo)
                    sGateOPen_YN = "Y"
                End If
                
                If (sLane_INOUT = "�Ա�") Then
                    sEmergency_Print = "���ν�����"
                    sPASS_INOUT = "IN":                                 sPass_YN = sGateOPen_YN:                         sPASS_RESULT = "���ν�����"
                Else
                    sEmergency_Print = "���ν�����"
                    sPASS_INOUT = "OUT":                            sPass_YN = sGateOPen_YN:                         sPASS_RESULT = "���ν�����"
                End If
                
                
            Case Else '������(�ù�,�ý�,ȭ��..), �湮����
                    Taxi_F = False
                    Select Case LenH(Trim(Proc_CarNo))
                        Case 11  '����ȣ2 : ����8��6800
                            Select Case Mid(Trim(Proc_CarNo), 4, 1)
                                Case "��", "��", "��", "��", "��", "ī", "Ÿ", "��"
                                    Taxi_F = True
                                Case Else
                                    Taxi_F = False
                            End Select
                        Case 12
                            Select Case Mid(Trim(Proc_CarNo), 5, 1)
                                Case "��", "��", "��", "��", "��"
                                    Taxi_F = True
                                Case Else
                                    Taxi_F = False
                            End Select
                        Case 8
                            Select Case Mid(Trim(Proc_CarNo), 3, 1)
                                Case "��", "��", "��", "��", "��", "ī", "Ÿ", "��", "��"
                                    Taxi_F = True
                                Case Else
                                    Taxi_F = False
                            End Select
                        Case 9
                            Select Case Mid(Trim(Proc_CarNo), 4, 1) '123��1234
                                Case "��", "��", "��", "��", "��", "ī", "Ÿ", "��", "��"
                                    Taxi_F = True
                                Case Else
                                    Taxi_F = False
                            End Select
                    End Select
                
                    If Taxi_F = True Then

                            sEmergency_Print = "��������"

                            'ImageFile = Slash_Conv(ImageFile)
                            
                            If (sLane_INOUT = "�Ա�") Then
                                If (sTaxiPass = "Y") Then
                                    sGateOPen_YN = "Y"
                                    sPASS_INOUT = "IN":                                         sPass_YN = sGateOPen_YN:                                         sPASS_RESULT = "����������"
                                Else
                                    sPASS_INOUT = "IN":                                         sPass_YN = sGateOPen_YN:                                         sPASS_RESULT = "����������"
                                End If
                            Else
                                If sTaxiPass = "Y" Then
                                    sGateOPen_YN = "Y"
                                    sPASS_INOUT = "OUT":                                        sPass_YN = sGateOPen_YN:                                         sPASS_RESULT = "����������"
                                Else
                                    sPASS_INOUT = "OUT":                                        sPass_YN = sGateOPen_YN:                                         sPASS_RESULT = "����������"
                                End If
                            End If
                    
                    Else
                    
                        '�����湮���� ���� ó��
                        sGuestQry = "SELECT * FROM tb_guestReg where CAR_NO = '" & Trim(Proc_CarNo) & "' AND START_DATE <= '" & sGuestPassDate & "' AND END_DATE >= '" & sGuestPassDate & "' "
                        Set rsGuestReg = New ADODB.Recordset
                        'rsGuestReg.Open Qry, adoConn
                        bQryResult = DataBaseQuery(rsGuestReg, adoConn, sGuestQry, NWERR_GATE_STAY)
                        If (bQryResult = False) Then
                            DataLogger ("[LPRIN_PROC GuestReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                            Exit Sub
                        End If
                        
                        '�湮��������
                        If Not (rsGuestReg.EOF) Then
                        
                            sEmergency_Print = "�湮��������"
                            sGateOPen_YN = "Y"
'
'                            If (sLane_INOUT = "�Ա�") Then
'                                sPASS_INOUT = "IN":                                     sPASS_YN = sGateOPen_YN:                                     sPASS_RESULT = "�湮��������"
'                            Else
'                                sPASS_INOUT = "OUT":                                    sPASS_YN = sGateOPen_YN:                                     sPASS_RESULT = "�湮��������"
'                            End If

                            sCAR_GUBUN = "�湮����"

                            sDRIVER_NAME = "" & rsGuestReg!DRIVER_NAME
                            sDRIVER_PHONE = "" & rsGuestReg!DRIVER_PHONE
                            sDRIVER_DEPT = "" & rsGuestReg!DRIVER_DEPT   '��
                            sDRIVER_CLASS = "" & rsGuestReg!DRIVER_CLASS 'ȣ��
                            'sSTART_DATE = "" & rsGuestReg!START_DATE
                            'sEND_DATE = "" & rsGuestReg!END_DATE
                            sSTART_DATE = Format(rsGuestReg!START_DATE, "yyyy-mm-dd hh:nn:ss")
                            sEND_DATE = Format(rsGuestReg!END_DATE, "yyyy-mm-dd hh:nn:ss")
                            sGUESTREG_ID = "" & rsGuestReg!GUESTREG_ID '�����湮��û ����ID
                            
                            
                            
                            
'''                            '�����湮�������� ��,ȣ���� �����Ǽ� üũ ==> �ۿ��� �Ǽ�/�ð� üũ �̻���� ��� ��� �� �Ǽ� ������
'''                            If (sLane_INOUT = "�Ա�") Then
'''
'''                                '���� ��,ȣ���� ���̵� ������ ���� �� ����. �����Ǽ�, �����ð� ������ ���� ��,ȣ���� �����ϰ� ������.
'''                                sGuestRegAdminQry = "SELECT MAXPARKCOUNT,NOWPARKCOUNT FROM tb_guestReg_admin where DRIVER_DEPT = '" & sDRIVER_DEPT & "' AND DRIVER_CLASS = '" & sDRIVER_CLASS & "' "
'''                                Set rsGuestRegAdmin = New ADODB.Recordset
'''                                rsGuestRegAdmin.Open sGuestRegAdminQry, adoConn
'''
'''                                If Not (rsGuestRegAdmin.EOF) Then
'''                                    nMaxParkCount = Int(0 & rsGuestRegAdmin!MAXPARKCOUNT) '�ִ������Ǽ�(��)
'''                                    If (nMaxParkCount > 0) Then '0:�����Ǽ� üũ����, >0:�����Ǽ� üũ
'''
'''                                        nNowParkCount = Int(0 & rsGuestRegAdmin!NOWPARKCOUNT)
'''                                        nNowParkCount = nNowParkCount + 1
'''
'''                                        adoConn.Execute "UPDATE tb_guestReg_admin SET NOWPARKCOUNT = " & nNowParkCount & " WHERE DRIVER_DEPT = '" & sDRIVER_DEPT & "' AND DRIVER_CLASS = '" & sDRIVER_CLASS & "' "
'''
'''                                        If (nNowParkCount > nMaxParkCount) Then
'''                                            '�Ǽ��ʰ�
'''                                            sEmergency_Print = "�湮���ุ��"
'''                                            sGateOPen_YN = "N"
'''                                            Call DataLogger("[�湮���ุ��] " & Proc_CarNo & ", ��:" & sDRIVER_DEPT & ", ȣ��:" & sDRIVER_CLASS & ", ����Ǽ�/�ִ������Ǽ�:" & nNowParkCount & "/" & nMaxParkCount)
'''                                        End If
'''                                    End If
'''
'''                                '    rsGuestRegAdmin.MoveNext
'''                                'Loop
'''                                End If
'''                            End If
                            
                            
                            
                            If (sLane_INOUT = "�Ա�") Then
                                sPASS_INOUT = "IN":                                     sPass_YN = sGateOPen_YN:                                     sPASS_RESULT = "�湮��������"
                            Else
                                sPASS_INOUT = "OUT":                                    sPass_YN = sGateOPen_YN:                                     sPASS_RESULT = "�湮��������"
                            End If

                            
                            
                            '�����뺸
                            If (sLane_INOUT = "�Ա�") Then
                                If (HomeNet_YN = "Y") Then
'                                    If (IsNumeric(rsGuestReg!DRIVER_DEPT) = True) And (IsNumeric(rsGuestReg!DRIVER_CLASS) = True) And (rsGuestReg!DAY_ROTATION_YN = "����") Then
'                                        Call SendHomenet(rsGuestReg!DRIVER_DEPT, rsGuestReg!DRIVER_CLASS, Proc_CarNo)
'                                    End If
                                    If (IsNumeric(sDRIVER_DEPT) = True) And (IsNumeric(sDRIVER_CLASS) = True) And (rsGuestReg!DAY_ROTATION_YN = "����") Then
                                        Call SendHomenet(sDRIVER_DEPT, sDRIVER_CLASS, Proc_CarNo)
                                    End If
                                End If
                            End If

                        
                        '�湮����
                        Else
                            sEmergency_Print = "�湮����"
                            If (sFreePass = "Y") Then
                                sGateOPen_YN = "Y"
                            End If
                            
                            If (sLane_INOUT = "�Ա�") Then
                                sPASS_INOUT = "IN":                                     sPass_YN = sGateOPen_YN:                                     sPASS_RESULT = "�̵������"
                            Else
                                sPASS_INOUT = "OUT":                                    sPass_YN = sGateOPen_YN:                                     sPASS_RESULT = "�̵������"
                            End If
                        
                        End If
                        Set rsGuestReg = Nothing
                        
                        
                        
                    End If
                    
        End Select
    End If
    
    
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '���� ó����
    If (stSound.sSnd_YN = "Y") Then
        If ((InStr(1, sEmergency_Print, "�湮��������") > 0)) Then
            Call Sound_Out(stSound.sSndFName_Reg)
        ElseIf ((InStr(1, sEmergency_Print, "���") > 0) And (stSound.sSndReg_YN = "Y")) Then
            Call Sound_Out(stSound.sSndFName_Reg)
        ElseIf ((InStr(1, sEmergency_Print, "���ν�") > 0) And (stSound.sSndNoRec_YN = "Y")) Then
            Call Sound_Out(stSound.sSndFName_NoRec)
        ElseIf ((InStr(1, sEmergency_Print, "�湮����") > 0) And (stSound.sSndGuest_YN = "Y")) Then
            Call Sound_Out(stSound.sSndFName_Guest)
        ElseIf ((InStr(1, sEmergency_Print, "��������") > 0) And (stSound.sSndBlackList_YN = "Y")) Then
            Call Sound_Out(stSound.sSndFName_BlackList)
        ElseIf ((InStr(1, sEmergency_Print, "����") > 0) And (stSound.sSndTaxi_YN = "Y")) Then
            Call Sound_Out(stSound.sSndFName_Taxi)
        ElseIf ((InStr(1, sEmergency_Print, "��������") > 0) And (stSound.sSndDay_YN = "Y")) Then
            Call Sound_Out(stSound.sSndFName_Day)
        ElseIf (InStr(1, sEmergency_Print, "�Ⱓ����") > 0) And (stSound.sSndRegExpDate_YN = "Y") Then
            Call Sound_Out(stSound.sSndFName_RegExpDate)
        Else
             Call Sound_Out(App.Path & "\sound\Bell.wav")
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    ' ����Ƚ������ ���� ��ȸ
    Set rs = New ADODB.Recordset
    qry = "SELECT * From tb_guest_limit WHERE CAR_NO = '" & carnum & "'"
    rs.Open qry, adoConn
    
    If Not (rs.EOF) Then
        Dim iMaxInPark As Integer
        Dim iNowInPark As Integer
        iMaxInPark = rs!MAXINPARK:         iNowInPark = rs!NOWINPARK
        
        If (rs!MAXINPARK > rs!NOWINPARK) Then
            adoConn.Execute "UPDATE tb_guest_limit SET NOWINPARK = " & rs!NOWINPARK + 1 & " WHERE CAR_NO = '" & carnum & "' "
        Else
            sEmergency_Print = "������������"
            sPass_YN = "N"
            sGateOPen_YN = "N"
            
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�����湮����', 'HOST','" & carnum & " -> ����Ƚ������ ����(���ܱ� �� ����)',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            Call DataLogger("[INFO] " & carnum & " : ����Ƚ������ ����(������: �ִ� " & rs!MAXINPARK & "ȸ(��)")
        End If
        
        
    End If
    Set rs = Nothing

    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' �ڸ���� ���(������ ���ܱ� ����)
    If (sNoWork = "�ڸ����") Then
        sPass_YN = "Y"
        sGateOPen_YN = "Y"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '���� ����Ⱓ ���ĺ��� ���ܱ� ���� ����
    If (Glo_Certify = enumCertify.eCertTry) Then
        If (Glo_Cert_LimitDate < Format(Now, "yyyy-mm-dd")) Then
            sGateOPen_YN = "N"
            Call DataLogger("[WARNING!!] �����Ⱓ ����� ���� ���ܱ� �� ����")
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ���ܱ� ó����
    ' ���� : DB ó���� ���ʿ� ��ġ�ؾ� ��. DBó���ο��� �����߻��� ����ó���� ������ �Ǿ� ���ܱ� �ȿ��� �� ����.
    If (Glo_ParkFull_YN = "Y") Then
        
        '������� �켱 ó���ǹǷ�, ���ܱ��ڵ����� ��� ����ȵ�
        If (ParkFull_Proc(sLane_INOUT, sGateOPen_YN, sPASS_RESULT) = True) Then
            Call Relay_Out(0, iGateNo)
        End If
        
    Else
        If (sGateOPen_YN = "Y") Then
    '        If (Glo_ApsYN <> "Y" And Glo_PreApsYN <> "Y") Then 'ȣ��Ʈ �ܵ� ����� ���
                Call Relay_Out(0, iGateNo)
    '        End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ������ ó����
    If (LenH(sEmergency_Print) > 0) Then
        If (sLane_INOUT = "�Ա�") Then
                If (Glo_Display = "������" Or Glo_Display = "������(Ǯ�÷�)" Or Glo_Display = "������(Ǯ�÷�)_FW7") Then
                    ''''''''''''''''''''''''''''''''''''''''''''''''
                    ' ������� ó��
                    If (Glo_ParkFull_YN = "Y") Then '������� ���
                        Call ParkFull_Display(carnum, sEmergency_Print, iGateNo)
                        
                    Else
                    
                        'Call GL_Emergency(sEmergency_Print, carnum, 0, 30, 20, 1, 2, 1, iGateNo) ' ���� ����-���� �Ʒ���,ǥ�ýð� ����(5��° �Ķ���� - 10:5��, 20:10�� ���)
                        
                        
                        '���� ���� ������ ���� �̿��� ���� �����ϵ��� ��
                        If (InStr(1, sEmergency_Print, "�湮��������") > 0) Then
                            Call GL_Emergency("�湮��������", carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorReg, stEmerg.iDisp2EmergColorReg, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "�湮���ุ��") > 0) Then
                            Call GL_Emergency("�湮���ุ��", carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorReg, stEmerg.iDisp2EmergColorBKList, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "���") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergReg, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorReg, stEmerg.iDisp2EmergColorReg, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "���ν�") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergNoRec, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorNoRec, stEmerg.iDisp2EmergColorNoRec, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "�湮����") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergGuest, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorGuest, stEmerg.iDisp2EmergColorGuest, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "��������") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergBlackList, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorBKList, stEmerg.iDisp2EmergColorBKList, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "����") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergTaxi, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorTaxi, stEmerg.iDisp2EmergColorTaxi, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "��������") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergDay, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorDay, stEmerg.iDisp2EmergColorDay, iGateNo)
                        ElseIf (InStr(1, sEmergency_Print, "�Ⱓ����") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergRegExpDate, carnum, 0, 30, 20, 1, stEmerg.iDisp1EmergColorRegExpDate, stEmerg.iDisp2EmergColorRegExpDate, iGateNo)
                        Else
                            Call DebugLogger("������ �Ա� ���� �������� : " & sEmergency_Print)
                        End If
                        '���� ��
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    
                ElseIf Glo_Display = "FND" Then
                    If (Trim(Proc_CarNo) = "�νĽ���") Then
                        Call FND_Display(Right("----", 4), iGateNo)
                    Else
                        Call FND_Display(Right(Trim(Proc_CarNo), 4), iGateNo)
                    End If
                End If
                    
        '�ⱸ
        Else
                If (Glo_ApsYN <> "Y" And Glo_PreApsYN <> "Y") Then 'ȣ��Ʈ �ܵ� ����� ���
                
                    If (Glo_Display = "������" Or Glo_Display = "������(Ǯ�÷�)" Or Glo_Display = "������(Ǯ�÷�)_FW7") Then
                        '''Call GL_Emergency(carnum, sEmergency_Print, 0, 30, 10, 1, 2, 1, iGateNo)
                            If (InStr(1, sEmergency_Print, "�湮��������") > 0) Then
                                Call GL_Emergency("�湮��������", carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "�湮���ุ��") > 0) Then
                                Call GL_Emergency("�湮���ุ��", carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "���") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergReg, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "���ν�") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergNoRec, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "�湮����") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergGuest, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "��������") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergBlackList, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "����") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergTaxi, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "��������") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergDay, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            ElseIf (InStr(1, sEmergency_Print, "�Ⱓ����") > 0) Then
                                Call GL_Emergency(stEmerg.sEmergRegExpDate, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                            Else
                                Call DebugLogger("������ �ⱸ ���� �������� : " & sEmergency_Print & "_" & stEmerg.sEmergReg)
                            End If

                    ElseIf Glo_Display = "FND" Then
                        If (Trim(Proc_CarNo) = "�νĽ���") Then
                            Call FND_Display(Right("----", 4), iGateNo)
                        Else
                            Call FND_Display(Right(Trim(Proc_CarNo), 4), iGateNo)
                        End If
                    End If
                
                ElseIf (Glo_ApsYN = "Y" Or Glo_PreApsYN = "Y") Then '�������α�, �ⱸ���α� ���
                    If (sGateOPen_YN = "Y") Then '���ܱ���� ����
                        Call GL_Emergency(sEmergency_Print, carnum, 0, 30, 20, 1, 2, 1, iGateNo) ' ���� ����-���� �Ʒ���,ǥ�ýð� ����(10->15:��7��)
                    Else
                        If (InStr(1, sEmergency_Print, "��������") > 0) Then
                            Call GL_Emergency(stEmerg.sEmergBlackList, carnum, 0, 30, 10, 1, 2, 1, iGateNo)
                        End If
                    End If
                End If
        End If
    End If



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DBó����(������)
    ' tb_inout : ��� ���� ��������� ����
    qry = "INSERT INTO tb_inout VALUES ('" & Trim(Proc_CarNo) & "', '" & Trim(Rec_CarNo) & "', '" & sCAR_MODEL & "', '" & sCAR_GUBUN & "', '" & sDRIVER_NAME & "', '" & sDRIVER_PHONE & "', '" & sDRIVER_DEPT & "', '" & sDRIVER_CLASS & "', '" & sSTART_DATE & "', '" & sEND_DATE & "', '" & iGateNo & "', '" & sPASS_INOUT & "', '" & sPassDate & "', '" & sPass_YN & "', '" & sPASS_RESULT & "', '" & ImageFile & "', '" & sLPR_IP & "', 0)"
    bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY, iGateNo)
    If (bQryResult = False) Then
        DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
        Set rs = Nothing
        Exit Sub
    End If

    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DBó����(����)
    ' tb_now :�������α� �Ǵ� �ⱸ���α� ����� ��쿡�� �����ؾ� ��
    If (sLane_INOUT = "�Ա�") Then

            'If (Glo_ApsYN = "Y" Or Glo_PreApsYN = "Y") Then
                qry = "Delete From tb_now Where CAR_NO= '" & carnum & "'"
                bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_OPEN, iGateNo)
                If (bQryResult = False) Then
                    DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
                    Exit Sub
                End If
    
                'Qry = "INSERT INTO tb_now   VALUES ('" & Trim(Proc_CarNo) & "', '" & Trim(Rec_CarNo) & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & iGateno & "', '" & "IN" & "', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "', '" & "Y" & "', '" & sPASS_RESULT & "', '" & ImageFile & "', '" & sLPR_IP & "', '0', '')"
                'Qry = "INSERT INTO tb_now   VALUES ('" & Trim(Proc_CarNo) & "', '" & Trim(Rec_CarNo) & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & " " & "', '" & iGateNo & "', '" & "IN" & "', '" & sPassDate & "', '" & "Y" & "', '" & sPASS_RESULT & "', '" & ImageFile & "', '" & sLPR_IP & "', '0', '')"
                qry = "INSERT INTO tb_now   VALUES ('" & Trim(Proc_CarNo) & "', '" & Trim(Rec_CarNo) & "', '" & sCAR_MODEL & "', '" & sCAR_GUBUN & "', '" & sDRIVER_NAME & "', '" & sDRIVER_PHONE & "', '" & sDRIVER_DEPT & "', '" & sDRIVER_CLASS & "', '" & sSTART_DATE & "', '" & sEND_DATE & "', '" & iGateNo & "', '" & sPASS_INOUT & "', '" & sPassDate & "', '" & sPass_YN & "', '" & sPASS_RESULT & "', '" & ImageFile & "', '" & sLPR_IP & "', 0, '')"
                bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_OPEN, iGateNo)
                If (bQryResult = False) Then
                    DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
                    Exit Sub
                End If
            'End If
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DBó����(�湮��������)
    ' tb_guestReg_inout : "�湮��������" ������ ����
'''    If (InStr(1, sEmergency_Print, "�湮����") > 0) Then
'''        qry = "INSERT INTO tb_guestReg_inout (CAR_NO,REC_NO,CAR_GUBUN,DRIVER_NAME,DRIVER_PHONE,DRIVER_DEPT,DRIVER_CLASS,START_DATE,END_DATE,PASS_GATE,PASS_INOUT,PASS_DATE,PASS_YN,PASS_RESULT,PASS_IMAGE) VALUES ('" & Trim(Proc_CarNo) & "', '" & Trim(Rec_CarNo) & "', '�湮����', '" & sDRIVER_NAME & "', '" & sDRIVER_PHONE & "', '" & sDRIVER_DEPT & "', '" & sDRIVER_CLASS & "', '" & sSTART_DATE & "', '" & sEND_DATE & "', '" & iGateNo & "', '" & sPASS_INOUT & "', '" & sPassDate & "', '" & sPass_YN & "', '" & sPASS_RESULT & "', '" & ImageFile & "')"
'''        bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
'''
'''        adoConn.Execute "UPDATE  tb_guestReg SET PASS_YN = 'Y' WHERE CAR_NO = '" & Trim(Proc_CarNo) & "' AND START_DATE <= '" & sGuestPassDate & "' AND END_DATE >= '" & sGuestPassDate & "' "
'''    End If


    
    If (InStr(1, sEmergency_Print, "�湮����") > 0) Then

        adoConn.Execute "UPDATE  tb_guestReg SET PASS_YN = 'Y' WHERE CAR_NO = '" & Trim(Proc_CarNo) & "' AND START_DATE <= '" & sGuestPassDate & "' AND END_DATE >= '" & sGuestPassDate & "' " '��/���� ���ó��

        If (sLane_INOUT = "�ⱸ") Then
        
            Dim nParkTime As Integer
            Dim bResult As Boolean
            Dim rsNow As Recordset
            Set rsNow = New ADODB.Recordset
            bResult = DataBaseQuery(rsNow, adoConn, "SELECT * FROM tb_now Where CAR_NO= '" & carnum & "' ORDER BY pass_date desc LIMIT 1", False)
            If (Not rsNow.EOF) Then
                nParkTime = DateDiff("n", Left(rsNow!PASS_DATE, 19), Left(sPassDate, 19))
                '�����ð� ����
                adoConn.Execute "INSERT INTO tb_guestreg_daily (CAR_NO, DRIVER_DEPT, DRIVER_CLASS, IN_TIME, OUT_TIME, PARKTIME, DRIVER_NAME, DRIVER_PHONE, REG_DATE) VALUES ('" & rsNow!CAR_NO & "','" & sDRIVER_DEPT & "','" & sDRIVER_CLASS & "', '" & Left(rsNow!PASS_DATE, 19) & "', '" & Left(sPassDate, 19) & "', " & nParkTime & ", '" & sDRIVER_NAME & "', '" & sDRIVER_PHONE & "', '" & Left(sPassDate, 19) & "')"
                adoConn.Execute "UPDATE tb_guestreg_admin SET NOWPARKTIME = NOWPARKTIME + " & nParkTime & " WHERE ID = '" & sGUESTREG_ID & "' "
            End If
            Set rsNow = Nothing
        End If
        
        
        
    End If
    
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�ⱸ���α� ó������ ����
    'tb_now ���� �� ó������ �����ؾ���
    If (sLane_INOUT = "�ⱸ") Then
        If (Glo_ApsYN <> "Y" And Glo_PreApsYN <> "Y") Then 'ȣ��Ʈ �ܵ� ���
            qry = "Delete From tb_now Where CAR_NO= '" & carnum & "'"
            bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_OPEN, iGateNo)
            If (bQryResult = False) Then
                Call DebugLogger("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
                Exit Sub
            End If
        Else
            If (Glo_ApsYN = "Y") Then   '1����:�ⱸ���α� ���
                If (sGateOPen_YN <> "Y") Then '���ܱ� �����ʴ� ��쿡 ���ؼ� �ⱸ���α⿡�� ��ݰ���ؾ� ��.
                    If (InStr(1, sEmergency_Print, "��������") > 0) Then '�������� ������ ���α⿡�� ó������(��ܿ��� ������ �������� ǥ���� ���ܱ� �ȿ���)
                    Else
                        Glo_APS_Str = carnum
                        Call APS_Connect
                        Call FrmAccnt.APS_PutImage(Proc_CarNo, ImageFile) '���⿡ �ⱸ���� �������� �۽�
                    End If
                End If
            Else                        '2����:�������α���(�ⱸ���α� �̻��)
                If (Glo_PreApsYN = "Y") Then
                    If (sGateOPen_YN <> "Y") Then '���ܱ� �����ʴ� ���
                        Call PreAps_Proc(carnum, iGateNo, sPASS_RESULT, sPassDate)
                    End If
                End If
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End With
Set rs = Nothing

Exit Sub

Err_Proc:
    Call DataLogger("[LPRIn_Proc] " & Err.Description)

End Sub

Public Sub SendHomenet(Dong As String, Ho As String, carno As String)
    
    On Error Resume Next
    
    HomeNet_Dong = Dong
    HomeNet_Ho = Ho
    HomeNet_CarNo = carno
    
    HomeNet_Str = HomeNet_Dong & HomeNet_Ho & HomeNet_CarNo
    
    If (FrmTcpServer.HomeSock.State = sckClosed) Then
        
        FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
        FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
        FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
        
        FrmTcpServer.HomeSock.SendData (HomeNet_Str)
        Call DataLogger("[HomeNet UDP ����]  DATA = " & HomeNet_Str)
    Else
        FrmTcpServer.HomeSock.SendData (HomeNet_Str)
        Call DataLogger("[HomeNet UDP ����]  DATA = " & HomeNet_Str)
    End If
    
End Sub
Public Function ParkFull_Proc(sInOut As String, sGateOpen As String, sResult As String)
    On Error GoTo Err_p

        ParkFull_Proc = False
        
        If (sInOut = "�Ա�") Then
            If (sGateOpen = "Y") Then

                    If (Glo_ParkNow_Count < Glo_ParkFull_Count) Then ' �������ɻ���
                        Glo_ParkNow_Count = Glo_ParkNow_Count + 1
                        'Call Relay_Out(0, iGateno)
                        ParkFull_Proc = True
                        Call DataLogger("[��    ��] " & Glo_ParkNow_Count & " ��° ���� ����ó��")

                    Else    ' ��������
                        If (sResult = "��������") Then
                            
                            If (Glo_ParkRegIn_YN = "Y") Then
                                Glo_ParkNow_Count = Glo_ParkNow_Count + 1
                                'Call Relay_Out(0, iGateno)
                                ParkFull_Proc = True
                                Call DataLogger("[��    ��] ��������:������� ����ó��")
                            Else
                                Call DataLogger("[��    ��] ��������:������� ��������!!")
                            End If
                            
                        Else
                            Call DataLogger("[��    ��] ��������:������� �� ��������!!")
                        End If
                    End If

                    '����üũ
                    If (Glo_ParkNow_Count < Glo_ParkFull_Count) Then
                        Glo_ParkFull_Status = pkfStayNML '�������
                    ElseIf (Glo_ParkNow_Count = Glo_ParkFull_Count) Then
                        Glo_ParkFull_Status = pkfChangeFULL '����->������ ����
                    ElseIf (Glo_ParkNow_Count > Glo_ParkFull_Count) Then '��������
                        Glo_ParkFull_Status = pkfStayFULL '��������
                    End If
            
            Else
                If (Glo_ParkFull_Status = pkfChangeFULL) Then '���� ����->������ �����ߴٸ�
                    Glo_ParkFull_Status = pkfStayFULL '�̹� �������� �������� ����
                Else
                    Glo_ParkFull_Status = pkfStayNML '�������
                End If
            End If
            
        Else
            If (sGateOpen = "Y") Then
                Glo_ParkNow_Count = Glo_ParkNow_Count - 1
                If (Glo_ParkNow_Count < 0) Then
                    Glo_ParkNow_Count = 0
                End If
            
                'Call Relay_Out(0, iGateno)
                ParkFull_Proc = True

                
                Call DataLogger("[��    ��] ����ó��")
            End If
            
            
            '����üũ
            If (Glo_ParkNow_Count >= Glo_ParkFull_Count) Then
                Glo_ParkFull_Status = pkfStayFULL '��������
            ElseIf (Glo_ParkNow_Count + 1 = Glo_ParkFull_Count) Then
                Glo_ParkFull_Status = pkfChangeNML '����->������ ����
            ElseIf (Glo_ParkNow_Count < Glo_ParkFull_Count) Then '��������
                Glo_ParkFull_Status = pkfStayNML '����->������ ����
            End If
        End If
        
        Call Put_Ini("System Config", "ParkNow_Count", CStr(Glo_ParkNow_Count))
        
        Exit Function
Err_p:
        Call DataLogger("[��    ��]  ������ ó�� ����:" & Err.Description)
End Function

Public Sub ParkFull_Display(sCarnum As String, sEmergency As String, iGateNo As Integer)
    On Error GoTo Err_p

    '������ȣ ������ ��� ���� ����
    '1.������� 2.����->���� ���� 3.����->���� ����
    If (Glo_ParkFull_Status = pkfStayNML Or Glo_ParkFull_Status = pkfChangeNML Or Glo_ParkFull_Status = pkfChangeFULL) Then
        Call GL_Emergency(sEmergency, sCarnum, 0, 30, 20, 1, 2, 1, iGateNo) ' ���� ����-���� �Ʒ���,ǥ�ýð� ����(10->15:��7��)
    End If

    
    If (Glo_ParkFull_Status = pkfChangeFULL) Then '����->���� ����
        Call GL_Nomal("[��  ��]", "��������", 128, 0, 0, FrmTcpServer.cmb_Disp1(iGateNo).ListIndex, FrmTcpServer.cmb_Disp2(iGateNo).ListIndex, iGateNo) '����ȭ��
    ElseIf (Glo_ParkFull_Status = pkfChangeNML) Then     '����->���� ����
        Call GL_Nomal(FrmTcpServer.txt_Disp1(iGateNo), FrmTcpServer.txt_Disp2(iGateNo), 129, 70, 0, FrmTcpServer.cmb_Disp1(iGateNo).ListIndex, FrmTcpServer.cmb_Disp2(iGateNo).ListIndex, iGateNo)
    End If
    
    Exit Sub
Err_p:
    Call DataLogger("[��    ��]  ������ ��� ����:" & Err.Description)
End Sub

Public Sub ParkFull_PutNMLDisplay(iGateNo As Integer)
    On Error GoTo Err_p
    
    If (Glo_ParkNow_Count < Glo_ParkFull_Count) Then
        Call GL_Nomal(FrmTcpServer.txt_Disp1(iGateNo), FrmTcpServer.txt_Disp2(iGateNo), 129, 70, 0, FrmTcpServer.cmb_Disp1(iGateNo).ListIndex, FrmTcpServer.cmb_Disp2(iGateNo).ListIndex, iGateNo)
    Else
        Call GL_Nomal("[��  ��]", "��������", 128, 0, 0, FrmTcpServer.cmb_Disp1(iGateNo).ListIndex, FrmTcpServer.cmb_Disp2(iGateNo).ListIndex, iGateNo) '����ȭ��
    End If
    
    Exit Sub
Err_p:
    Call DataLogger("[��    ��]  ������ ��� ����:" & Err.Description)
End Sub

Public Sub ParkFull_GetState(iGateNo As Integer, sInOut As String)
    On Error GoTo Err_p

        If (sInOut = "�Ա�") Then
                    
            Glo_ParkNow_Count = Glo_ParkNow_Count + 1
            
            '����üũ
            If (Glo_ParkNow_Count < Glo_ParkFull_Count) Then
                Glo_ParkFull_Status = pkfStayNML '�������
            ElseIf (Glo_ParkNow_Count = Glo_ParkFull_Count) Then
                Glo_ParkFull_Status = pkfChangeFULL '����->������ ����
            ElseIf (Glo_ParkNow_Count > Glo_ParkFull_Count) Then '��������
                Glo_ParkFull_Status = pkfStayFULL '��������
            End If
        
        Else
            Glo_ParkNow_Count = Glo_ParkNow_Count - 1
            If (Glo_ParkNow_Count < 0) Then
                Glo_ParkNow_Count = 0
            End If
            
            '����üũ
            If (Glo_ParkNow_Count >= Glo_ParkFull_Count) Then
                Glo_ParkFull_Status = pkfStayFULL '��������
            ElseIf (Glo_ParkNow_Count + 1 = Glo_ParkFull_Count) Then
                Glo_ParkFull_Status = pkfChangeNML '����->������ ����
            ElseIf (Glo_ParkNow_Count < Glo_ParkFull_Count) Then '��������
                Glo_ParkFull_Status = pkfStayNML '����->������ ����
            End If
        End If
        
        Call Put_Ini("System Config", "ParkNow_Count", CStr(Glo_ParkNow_Count))
        
        Exit Sub
Err_p:
        Call DataLogger("[��    ��]  ������ ó�� ����:" & Err.Description)
End Sub
Public Sub ParkFull_Show()
    If Glo_Screen_No = 6 Then
        FrmG6_23.lbl_ParkFull.Caption = "������Ȳ : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
    ElseIf Glo_Screen_No = 4 Then
        FrmG4Mini.lbl_ParkFull.Caption = "������Ȳ : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
    ElseIf Glo_Screen_No = 2 Then
        Jung.lbl_ParkFull.Caption = "������Ȳ : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
    ElseIf Glo_Screen_No = 1 Then
        FrmG1.lbl_ParkFull.Caption = "������Ȳ : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
    End If
End Sub

Public Sub ParkFull_Visible(bVisible As Boolean)
    If Glo_Screen_No = 6 Then
        FrmG6_23.lbl_ParkFull.Visible = bVisible
    ElseIf Glo_Screen_No = 4 Then
        FrmG4Mini.lbl_ParkFull.Visible = bVisible
    ElseIf Glo_Screen_No = 2 Then
        Jung.lbl_ParkFull.Visible = bVisible
    ElseIf Glo_Screen_No = 1 Then
        FrmG1.lbl_ParkFull.Visible = bVisible
    End If
End Sub

Public Sub ParkFull_Set()
    Dim i As Integer
    Dim bVisible As Boolean
    Dim sLane_YN(MAX_LANE_COUNT) As String
    Dim sLane_INOUT(MAX_LANE_COUNT) As String
    
    If (Glo_ParkFull_YN = "Y") Then
        Call ParkFull_Visible(True)
        Call ParkFull_Show
    Else
        Call ParkFull_Visible(False)
    End If
        
    
    If (Glo_ParkFull_YN = "Y") Then
        sLane_YN(0) = LANE1_YN: sLane_INOUT(0) = LANE1_Inout
        sLane_YN(1) = LANE2_YN: sLane_INOUT(1) = LANE2_Inout
        sLane_YN(2) = LANE3_YN: sLane_INOUT(2) = LANE3_Inout
        sLane_YN(3) = LANE4_YN: sLane_INOUT(3) = LANE4_Inout
        sLane_YN(4) = LANE5_YN: sLane_INOUT(4) = LANE5_Inout
        sLane_YN(5) = LANE6_YN: sLane_INOUT(5) = LANE6_Inout

        For i = 0 To Glo_Screen_No - 1
            If (sLane_YN(i) = "Y" And sLane_INOUT(i) = "�Ա�") Then
                Call ParkFull_PutNMLDisplay(i) '�Ϲݹ������
           End If
        Next i
    End If

End Sub


Public Sub ParkFullLight_Set()
    If (Glo_ParkFullLIGHT_YN = "Y") Then
        If (FrmTcpServer.Timer_ParkFullLight.Enabled = False) Then
            FrmTcpServer.Timer_ParkFullLight.Enabled = True
        End If
    Else
        FrmTcpServer.Timer_ParkFullLight.Enabled = False
    End If
End Sub

Public Function Slash_Conv(str As String) As String
Dim i As Integer
Dim tmp As String
Dim Ret As Boolean

tmp = "\\\\"

For i = 3 To LenH(str) Step 1
    If (Mid(str, i, 1) = "\") Then
        tmp = tmp & "\\" & Mid(str, i, 1)
    Else
        tmp = tmp & Mid(str, i, 1)
    End If
Next i

Slash_Conv = tmp

End Function


