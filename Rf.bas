Attribute VB_Name = "Module3"
Option Explicit

Public DataSync_F(7) As Boolean

Public Sync_Timer_Cnt(7) As Integer


Public Am_Time As String
Public Pm_Time As String
Public MsgRet As Boolean



Public Sub RemoteIn_Proc(CarNum As String, Comm_Port As Integer, FTP_File As String)
Dim TagNum As String * 10
Dim tmp As Long
Dim Time_Band As String

Dim tmp_str As String
Dim JIO_Status As String * 8
Dim JIO_IOFlag As String * 1

Dim idx As Integer

Dim Pass_Mode As Integer
Dim Ret
Dim t As String
Dim Dir_File As String
Dim Q_Cnt As Integer
Dim Car_Num_Str As String
Dim Car_i As Integer
Dim i As Integer
Dim Save_CarNum As String
Dim RecStat As String
Dim JungRs As ADODB.Recordset
Dim JungIORs As ADODB.Recordset
Dim Qry As String


On Error GoTo Err_Proc


t = MidH(In_Img_Folder(Comm_Port), 14, 50)
t = MidH(t, 1, LenH(t) - 5)
t = t & "�Ա�"


Save_CarNum = CarNum
Anal_InCnt = Anal_InCnt + 1
With main

idx = Comm_Port

.LblRecNum(Comm_Port * 2).Caption = "�νĹ�ȣ : " & Save_CarNum
.LblTime(Comm_Port * 2).Caption = "ó������ : " & Format(Now, "yy-mm-dd")
.LblEtc(Comm_Port * 2).Caption = "ó���ð� : " & Format(Now, "hh:nn:ss")

Set JungRs = New ADODB.Recordset
Qry = "SELECT * FROM regcar WHERE ������ȣ ='" & Save_CarNum & "'"
JungRs.Open Qry, AdoConn

If Not (JungRs.EOF) Then
    .LblRecStat(Comm_Port * 2).Caption = "�νİ�� : �����ν�"
    RecStat = "�����ν�"
    Anal_OkCnt = Anal_OkCnt + 1
Seek_DB:
    Dir_File = Dir(In_Img_Folder(Comm_Port) & "*.jpg")
    If (Dir_File <> "") Then
        Name In_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
    End If
    .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & JungRs!������ȣ
    .LblRecName(Comm_Port * 2).Caption = "��      �� : " & JungRs!�̸�
    .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & JungRs!�Ҽ�
    If ((JungRs!������ > Format(Now, "yyyy-mm-dd")) Or (JungRs!������ < Format(Now, "yyyy-mm-dd"))) Then
        '----------------------------------------- �Ⱓ�ʰ� ó�� -------------------------------------------------------------
        .LblRecProc(Comm_Port * 2).Caption = "ó����� : �Ⱓ�ʰ�"
        AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
        "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�Ⱓ�ʰ�', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
        
        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�Ⱓ�ʰ� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  ������ : " & JungRs!������ & "]"
        If (.Check1.Value = 1) Then
            .List1.ListIndex = .List1.ListCount - 1
        End If
    Else
        Pass_Mode = JungRs!�Ա�1
        Select Case Pass_Mode
               Case 0 '���ϻ����
                    .LblRecProc(Comm_Port * 2).Caption = "ó����� : ��������"
                '----------------------------------------- ���� ó�� -------------------------------------------------------------
                        AdoConn.Execute "UPDATE regcar SET ���������� = 1, ������ = ������ + 1, �������� = " & "'" & Format(Now, "yyyy-mm-dd") & "', �����ð� = " & "'" & Format(Now, "hh:nn") & "'" & "WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
                        AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                        "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�������� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
               Case 1 '�ְ������
                        .LblRecProc(Comm_Port * 2).Caption = "ó����� : �ְ����"
                        If ((Format(Now, "hh:nn") >= Mid(Am_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Am_Time, 7, 5))) Then
                            JIO_Status = "�ְ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�ְ���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "�ְ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '���Ա���', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�ְ���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 2 '�߰������
                        .LblRecProc(Comm_Port * 2).Caption = "ó����� : �߰����"
                        If ((Format(Now, "hh:nn") >= Mid(Pm_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Pm_Time, 7, 5))) Then
                            JIO_Status = "�߰�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�߰���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "�߰�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '���Ա���', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�߰���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 3 '�ָ������
                        .LblRecProc(Comm_Port * 2).Caption = "ó����� : �ָ����"
                        
                        If (Weekday(Now) = 7) Then
                            JIO_Status = "�ָ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�ָ���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "�ָ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '���Ա���', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�ָ���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
        End Select
        For i = 0 To 9
            If (IsNull(JungRs(30 + (i * 2)).Value) Or (JungRs(30 + (i * 2)).Value = "")) Then
                AdoConn.Execute "UPDATE regcar SET �νĹ�ȣ" & CStr(i + 1) & "='" & Save_CarNum & "', ó�����" & CStr(i + 1) & "='" & RecStat & "' WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
                Exit For
            End If
        Next i
        If (i = 10) Then
            For i = 10 To 2 Step -1
                AdoConn.Execute "UPDATE regcar SET �νĹ�ȣ" & CStr(i) & "=�νĹ�ȣ" & CStr(i - 1) & ", ó�����" & CStr(i) & "=ó�����" & CStr(i - 1) & " WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
            Next i
            AdoConn.Execute "UPDATE regcar SET �νĹ�ȣ1='" & Save_CarNum & "', ó�����1='" & RecStat & "' WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
        End If
    End If
Else
    If (CarNum = "�νĽ���") Then
        '////////////////////// �νĽ��� //////////////////////////////////////
            .LblRecStat(Comm_Port * 2).Caption = "�νİ�� : �νĽ���"
            RecStat = "�νĽ���"
            Anal_FailCnt = Anal_FailCnt + 1
    Else
            Q_Cnt = IsChar(CarNum)
            If (Q_Cnt > 0) Then
                RecStat = "�κ��ν�"
                Anal_HalfCnt = Anal_HalfCnt + 1
                CarNum = XToQ(CarNum)
                If (Half_Rec_Mode) Then
                    '////////////////////// �κ��ν� //////////////////////////////////////
                    RecStat = "�κ��ν�"
                    .LblRecStat(Comm_Port * 2).Caption = "�νİ�� : �κ��ν�"
                    If (Q_Cnt <= Half_Cnt) Then
                       If ((LenH(CarNum) = 7) Or (LenH(CarNum) = 8)) Then
                            CarNum = QToAll(CarNum)
                       Else
                            '���12��1234 , ���1��1234 , 12��1234
                             If (MidH(CarNum, 1, 1) = "%") Then
                                 CarNum = "%%%" & CarNum
                             End If
                             
                             Car_Num_Str = RightH(CarNum, 5)
                             If (MidH(Car_Num_Str, 1, 1) = "%") Then
                                 CarNum = MidH(CarNum, 1, (LenH(CarNum) - 5)) & "%%" & RightH(CarNum, 4)
                             End If
                       End If
                        Set JungRs = New ADODB.Recordset
                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & CarNum & "'"
                        JungRs.Open Qry, AdoConn
                        If Not (JungRs.EOF) Then
Record_Found:
                            JungRs.MoveLast
                            If (JungRs.RecordCount > 1) Then
                                    Select Case Half_Rec_DoubleRecord_Process
                                              Case "M"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [�κ��ν� ���͸�ó�� >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                        'Call Data_ReSearch(LenH(CarNum))
                                                        Exit Sub
                                              Case "Y"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [�κ��ν� ���͸�ó�� >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                       'Call DRelay(Comm_Port, 0)
                                              Case "N"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [�κ��ν� ���͸�ó�� >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                    End Select
                            Else
                                If (Half_Rec_OneRecord_Process) Then
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [�κ��ν� ���͸�ó�� >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻����=" & JungRs!������ȣ
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    GoTo Seek_DB
                                Else
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [�κ��ν� ���͸�ó�� >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻����=" & JungRs!������ȣ
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    '.Text2.Text = JungRs!������ȣ
                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                    Exit Sub
                                End If
                            End If
                        Else
                            If (LenH(CarNum) = 11) Then '���12��7890 ����  ���1��7890  "2"�� �ҽǵȰ�� ���1?��7890  �� ġȯ�Ͽ� ��˻��Ѵ�
                                CarNum = MidH(CarNum, 1, 5) & "%" & RightH(CarNum, 6)
                                Set JungRs = New ADODB.Recordset
                                Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & CarNum & "'"
                                JungRs.Open Qry, AdoConn
                                If Not (JungRs.EOF) Then
                                    GoTo Record_Found
                                End If
                            End If
                        End If
                    Else
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [�κ��ν� ���͸� ��������(" & Half_Cnt & ") �ʰ�] �νĹ�ȣ=" & Save_CarNum
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                End If
            Else '////////////////////// �̵�� �Ǵ� ���ν� //////////////////////////////////////
                If (No_Rec_Mode) Then
                    .LblRecStat(Comm_Port * 2).Caption = "�νİ�� : ���ν�"
                    RecStat = "���ν�"
                    Anal_XXCnt = Anal_XXCnt + 1
                    Select Case Rec_Level
                              Case 1
                                    Select Case LenH(CarNum)
                                           Case 11
                                                    For Car_i = 1 To 16
                                                        Select Case Car_i
                                                               Case 1
                                                                    Car_Num_Str = "%%%%" & MidH(CarNum, 5, 7)
                                                               Case 2
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%" & MidH(CarNum, 6, 6)
                                                               Case 3
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%%" & MidH(CarNum, 8, 4)
                                                               Case 4
                                                                    Car_Num_Str = MidH(CarNum, 1, 7) & "%" & MidH(CarNum, 9, 3)
                                                               Case 5
                                                                    Car_Num_Str = MidH(CarNum, 1, 8) & "%" & MidH(CarNum, 10, 2)
                                                               Case 6
                                                                    Car_Num_Str = MidH(CarNum, 1, 9) & "%" & MidH(CarNum, 11, 1)
                                                               Case 7
                                                                    Car_Num_Str = MidH(CarNum, 1, 10) & "%"
                                                               Case 8
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 6)
                                                               Case 9 '���12��7890 : 12�� 2�� �ν��� �ȵȰ��(���ڴ���)
                                                                    Car_Num_Str = "%%%%" & MidH(CarNum, 5, 1) & "%" & MidH(CarNum, 6, 6)
                                                               Case 10
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%%" & MidH(CarNum, 6, 6)
                                                               Case 11
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%%%" & MidH(CarNum, 8, 4)
                                                               Case 12
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 2) & "%" & MidH(CarNum, 9, 3)
                                                               Case 13
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 3) & "%" & MidH(CarNum, 10, 2)
                                                               Case 14
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 4) & "%" & MidH(CarNum, 11, 1)
                                                               Case 15
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 5) & "%"
                                                        End Select
                                                        Set JungRs = New ADODB.Recordset
                                                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & Car_Num_Str & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!������ȣ
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 12 '����ȣ2 : ����81��6800
                                                    For Car_i = 1 To 8
                                                        Select Case Car_i
                                                               Case 1
                                                                    Car_Num_Str = "%%%%" & MidH(CarNum, 5, 8)
                                                               Case 2
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%" & MidH(CarNum, 6, 7)
                                                               Case 3
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 7, 6)
                                                               Case 4
                                                                    Car_Num_Str = MidH(CarNum, 1, 6) & "%" & MidH(CarNum, 9, 4)
                                                               Case 5
                                                                    Car_Num_Str = MidH(CarNum, 1, 8) & "%" & MidH(CarNum, 10, 3)
                                                               Case 6
                                                                    Car_Num_Str = MidH(CarNum, 1, 9) & "%" & MidH(CarNum, 11, 2)
                                                               Case 7
                                                                    Car_Num_Str = MidH(CarNum, 1, 10) & "%" & MidH(CarNum, 12, 1)
                                                               Case 8
                                                                    Car_Num_Str = MidH(CarNum, 1, 11) & "%"
                                                        End Select
                                                        Set JungRs = New ADODB.Recordset
                                                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & Car_Num_Str & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!������ȣ
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 8 '�Ź�ȣ1 : 81��7849
                                                    For Car_i = 1 To 7
                                                        Select Case Car_i
                                                               Case 1
                                                                    Car_Num_Str = "%" & MidH(CarNum, 2, 7)
                                                               Case 2
                                                                    Car_Num_Str = MidH(CarNum, 1, 1) & "%" & MidH(CarNum, 3, 6)
                                                               Case 3
                                                                    Car_Num_Str = MidH(CarNum, 1, 2) & "%%" & MidH(CarNum, 5, 4)
                                                               Case 4
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%" & MidH(CarNum, 6, 3)
                                                               Case 5
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 7, 2)
                                                               Case 6
                                                                    Car_Num_Str = MidH(CarNum, 1, 6) & "%" & MidH(CarNum, 8, 1)
                                                               Case 7
                                                                    Car_Num_Str = MidH(CarNum, 1, 7) & "%"
                                                        End Select
                                                        Set JungRs = New ADODB.Recordset
                                                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & Car_Num_Str & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!������ȣ
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = Car_Num_Str
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                    End Select
                              Case 2
                                        Set JungRs = New ADODB.Recordset
                                        Qry = "SELECT * FROM regcar WHERE (len(������ȣ ) = " & LenH(CarNum) & ") AND (������ȣ  Like '%" & RightH(CarNum, 6) & "')"
                                        JungRs.Open Qry, AdoConn
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����2 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����2 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!������ȣ
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����2 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   '.Text2.Text = RightH(CarNum, 6)
                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����2 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����2 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                              Case 3
                                        Set JungRs = New ADODB.Recordset
                                        Qry = "SELECT * FROM regcar WHERE (len(������ȣ ) = " & LenH(CarNum) & ") AND (������ȣ  Like '%" & RightH(CarNum, 4) & "')"
                                        JungRs.Open Qry, AdoConn
                                        
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    RecStat = "���ν�"
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����3 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����3 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!������ȣ
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����3 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = RightH(CarNum, 4)
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����3 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա� [���ν� ���͸�ó�� >> ���͸�����3 >> �ߺ����ڵ� >> ó������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                    End Select
                    '.SSPanel3(4).ForeColor = vbBlue
                    RecStat = "���ν�"
                Else
                    '.SSPanel3(4).ForeColor = vbWhite
                    RecStat = "���ν�"
                End If
            .LblRecStat(Comm_Port * 2).Caption = "�νİ�� : �̵��"
            End If
    End If
    Select Case RecStat
           Case "�νĽ���"
                .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & "�νĽ���"
                .LblRecName(Comm_Port * 2).Caption = "��      �� : " & "�νĽ���"
                .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & "�νĽ���"
                .LblRecProc(Comm_Port * 2).Caption = "ó����� : �νĽ���"
                Dir_File = Dir(In_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name In_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & "�νĽ���" & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "�νĽ���" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�νĽ��� �����Դϴ�. ]"
                If (.Check1.Value = 1) Then
                    .List1.ListIndex = .List1.ListCount - 1
                End If
           
           Case "���ν�", "�κ��ν�"
                Dir_File = Dir(In_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name In_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                If (RecStat = "�κ��ν�") Then
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & JungRs!������ȣ
                            .LblRecName(Comm_Port * 2).Caption = "��      �� : " & JungRs!�̸�
                            .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & JungRs!�Ҽ�
                            .LblRecProc(Comm_Port * 2).Caption = "ó����� : �����ߺ�"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�����ߺ�', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�����ߺ� �߻�! ��������ó�� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & "�̵��"
                            .LblRecName(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                            .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                            .LblRecProc(Comm_Port * 2).Caption = "ó����� : �̵��"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & "�̵��"
                        .LblRecName(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                        .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                        .LblRecProc(Comm_Port * 2).Caption = "ó����� : �̵��"
                        AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                        "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                        
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                Else
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & JungRs!������ȣ
                            .LblRecName(Comm_Port * 2).Caption = "��      �� : " & JungRs!�̸�
                            .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & JungRs!�Ҽ�
                            .LblRecProc(Comm_Port * 2).Caption = "ó����� : �����ߺ�"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�����ߺ�', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�����ߺ� �߻�! ��������ó�� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                            
                        Else
                            .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & "�̵��"
                            .LblRecName(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                            .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                            .LblRecProc(Comm_Port * 2).Caption = "ó����� : �̵��"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        .LblRecCarNum(Comm_Port * 2).Caption = "������ȣ : " & "�̵��"
                        .LblRecName(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                        .LblRecSosok(Comm_Port * 2).Caption = "��      �� : " & "�̵��"
                        .LblRecProc(Comm_Port * 2).Caption = "ó����� : �̵��"
                        AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                        "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�Ϲݱ�', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������Ա�" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                End If
    End Select
End If

If (.Frame3.Visible = True) Then
    .Label3.Caption = "���������:" & Anal_InCnt & Space(6 - Len(Anal_InCnt)) & "�����ν�:" & Anal_OkCnt & Space(6 - Len(Anal_OkCnt)) & "�κ��ν�:" & Anal_HalfCnt & Space(6 - Len(Anal_HalfCnt)) & "�νĽ���:" & Anal_FailCnt & Space(6 - Len(Anal_FailCnt)) & "���ν�:" & Anal_XXCnt & Space(6 - Len(Anal_XXCnt)) & "�νķ� = " & Int((Anal_OkCnt / Anal_InCnt) * 100) & "%" & Space(3) & "������ = " & Int(((Anal_OkCnt + Anal_HalfCnt + Anal_XXCnt) / Anal_InCnt) * 100) & "%"
    Put_Ini "�νķ��м�", "�������", CStr(Anal_InCnt)
    Put_Ini "�νķ��м�", "�����ν�", CStr(Anal_OkCnt)
    Put_Ini "�νķ��м�", "�κ��ν�", CStr(Anal_HalfCnt)
    Put_Ini "�νķ��м�", "�νĽ���", CStr(Anal_FailCnt)
    Put_Ini "�νķ��м�", "���ν�", CStr(Anal_XXCnt)
End If
End With
Set JungRs = Nothing
Exit Sub
Err_Proc:

Report_Write 0, "���о���", "ȣ��Ʈ", "**********", "RemoteIn_Proc �����߻�!", True
Report_Write 0, "��⳻��", "ȣ��Ʈ", Err.Number, Err.Description, True
Call Err_doc("ȣ��Ʈ : RemoteIn_Proc �����߻�!  Error��ȣ = " & Err.Number & "       Error���� = " & Err.Description)
End Sub


Public Sub RemoteOut_Proc(CarNum As String, Comm_Port As Integer, FTP_File As String)
Dim TagNum As String * 10
Dim tmp As Long
Dim Time_Band As String
Dim JungRs As ADODB.Recordset
Dim JungIORs As ADODB.Recordset
Dim JIO_Status As String * 8
Dim idx As Integer
Dim Pass_Mode As Integer
Dim Ret
Dim t As String
Dim Dir_File As String
Dim Q_Cnt As Integer
Dim Car_Num_Str As String
Dim Car_i As Integer
Dim i As Integer
Dim Save_CarNum As String
Dim RecStat As String

Dim Qry As String

On Error GoTo Err_Proc

t = MidH(In_Img_Folder(Comm_Port), 14, 50)
t = MidH(t, 1, LenH(t) - 5)
t = t & "�ⱸ"

Save_CarNum = CarNum
Anal_InCnt = Anal_InCnt + 1
With main

.LblRecNum(Comm_Port * 2 + 1).Caption = "�νĹ�ȣ : " & Save_CarNum
.LblTime(Comm_Port * 2 + 1).Caption = "ó���ð� : " & Format(Now, "yy-mm-dd")
.LblEtc(Comm_Port * 2 + 1).Caption = "ó���ð� : " & Format(Now, "hh:nn:ss")


Set JungRs = New ADODB.Recordset
Qry = "SELECT * FROM regcar WHERE ������ȣ ='" & Save_CarNum & "'"
JungRs.Open Qry, AdoConn



If Not (JungRs.EOF) Then
    .LblRecStat(Comm_Port * 2 + 1).Caption = "�νİ�� : �����ν�"
    RecStat = "�����ν�"
    Anal_OkCnt = Anal_OkCnt + 1
Seek_DB:
    Dir_File = Dir(Out_Img_Folder(Comm_Port) & "*.jpg")
    If (Dir_File <> "") Then
        Name Out_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
    End If
    .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & JungRs!������ȣ
    .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & JungRs!�̸�
    .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & JungRs!�Ҽ�
    If ((JungRs!������ > Format(Now, "yyyy-mm-dd")) Or (JungRs!������ < Format(Now, "yyyy-mm-dd"))) Then
        '----------------------------------------- �Ⱓ�ʰ� ó�� -------------------------------------------------------------
        .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �Ⱓ�ʰ�"
        AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
        "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�Ⱓ�ʰ�', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
        
        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�Ⱓ�ʰ� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  ������ : " & JungRs!������ & "]"
        If (.Check1.Value = 1) Then
            .List1.ListIndex = .List1.ListCount - 1
        End If
    Else
        Pass_Mode = JungRs!�Ա�1
        Select Case Pass_Mode
               Case 0 '���ϻ����
                    .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : ��������"
                '----------------------------------------- ���� ó�� -------------------------------------------------------------
                        AdoConn.Execute "UPDATE regcar SET ���������� = 2, ������ = ������ + 1, �������� = " & "'" & Format(Now, "yyyy-mm-dd") & "', �����ð� = " & "'" & Format(Now, "hh:nn") & "'" & "WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
                        AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                        "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�������� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
               Case 1 '�ְ������
                        .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �ְ����"
                        If ((Format(Now, "hh:nn") >= Mid(Am_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Am_Time, 7, 5))) Then
                            JIO_Status = "�ְ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�������� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "�ְ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '���Ա���', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�ְ���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 2 '�߰������
                        .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �߰����"
                        If ((Format(Now, "hh:nn") >= Mid(Pm_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Pm_Time, 7, 5))) Then
                            JIO_Status = "�߰�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�������� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "�߰�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '���Ա���', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�߰���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 3 '�ָ������
                        .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �ָ����"
                        If (Weekday(Now) = 7) Then
                            JIO_Status = "�ָ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '��������', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�������� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "�ָ�����"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & JungRs!�������� & "', '" & JungRs!�����ð� & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '���Ա���', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!��ȭ��ȣ & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�ָ���� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
        End Select
        
        For i = 0 To 9
            If (IsNull(JungRs(30 + (i * 2)).Value) Or (JungRs(30 + (i * 2)).Value = "")) Then
                AdoConn.Execute "UPDATE regcar SET �νĹ�ȣ" & CStr(i + 1) & "='" & Save_CarNum & "', ó�����" & CStr(i + 1) & "='" & RecStat & "' WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
                Exit For
            End If
        Next i
        If (i = 10) Then
            For i = 10 To 2 Step -1
                AdoConn.Execute "UPDATE regcar SET �νĹ�ȣ" & CStr(i) & "=�νĹ�ȣ" & CStr(i - 1) & ", ó�����" & CStr(i) & "=ó�����" & CStr(i - 1) & " WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
            Next i
            AdoConn.Execute "UPDATE regcar SET �νĹ�ȣ1='" & Save_CarNum & "', ó�����1='" & RecStat & "' WHERE ������ȣ= " & "'" & JungRs!������ȣ & "'"
        End If
    End If
Else
    If (CarNum = "�νĽ���") Then
        '////////////////////// �νĽ��� //////////////////////////////////////
            .LblRecStat(Comm_Port * 2 + 1).Caption = "�νİ�� : �νĽ���"
            RecStat = "�νĽ���"
            Anal_FailCnt = Anal_FailCnt + 1
    Else
            Q_Cnt = IsChar(CarNum)
            If (Q_Cnt > 0) Then
                RecStat = "�κ��ν�"
                Anal_HalfCnt = Anal_HalfCnt + 1
                CarNum = XToQ(CarNum)
                If (Half_Rec_Mode) Then
                    '////////////////////// �κ��ν� //////////////////////////////////////
                    RecStat = "�κ��ν�"
                    .LblRecStat(Comm_Port * 2 + 1).Caption = "�νİ�� : �κ��ν�"
                    If (Q_Cnt <= Half_Cnt) Then
                       If ((LenH(CarNum) = 7) Or (LenH(CarNum) = 8)) Then
                            CarNum = QToAll(CarNum)
                       Else
                            '���12��1234 , ���1��1234 , 12��1234
                             If (MidH(CarNum, 1, 1) = "%") Then
                                 CarNum = "%%%" & CarNum
                             End If
                             
                             Car_Num_Str = RightH(CarNum, 5)
                             If (MidH(Car_Num_Str, 1, 1) = "%") Then
                                 CarNum = MidH(CarNum, 1, (LenH(CarNum) - 5)) & "%%" & RightH(CarNum, 4)
                             End If
                       End If
                        Set JungRs = New ADODB.Recordset
                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & CarNum & "'"
                        JungRs.Open Qry, AdoConn
                        If Not (JungRs.EOF) Then
Record_Found:
                            JungRs.MoveLast
                            If (JungRs.RecordCount > 1) Then
                                    Select Case Half_Rec_DoubleRecord_Process
                                              Case "M"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [�κ��ν� ���͸�ó�� >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                        'Call Data_ReSearch(LenH(CarNum))
                                                        Exit Sub
                                              Case "Y"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [�κ��ν� ���͸�ó�� >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                       'Call DRelay(Comm_Port, 0)
                                              Case "N"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [�κ��ν� ���͸�ó�� >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                    End Select
                            Else
                                If (Half_Rec_OneRecord_Process) Then
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [�κ��ν� ���͸�ó�� >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻����=" & JungRs!������ȣ
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    GoTo Seek_DB
                                Else
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [�κ��ν� ���͸�ó�� >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & CarNum & " ,�˻����=" & JungRs!������ȣ
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            If (LenH(CarNum) = 11) Then '���12��7890 ����  ���1��7890  "2"�� �ҽǵȰ�� ���1?��7890  �� ġȯ�Ͽ� ��˻��Ѵ�
                                CarNum = MidH(CarNum, 1, 5) & "%" & RightH(CarNum, 6)
                                Set JungRs = New ADODB.Recordset
                                Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & CarNum & "'"
                                JungRs.Open Qry, AdoConn
                                If Not (JungRs.EOF) Then
                                    GoTo Record_Found
                                End If
                            End If
                        End If
                    Else
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [�κ��ν� ���͸� ��������(" & Half_Cnt & ") �ʰ�] �νĹ�ȣ=" & Save_CarNum
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                End If
            Else '////////////////////// �̵�� �Ǵ� ���ν� //////////////////////////////////////
                If (No_Rec_Mode) Then
                    .LblRecStat(Comm_Port * 2 + 1).Caption = "�νİ�� : ���ν�"
                    RecStat = "���ν�"
                    Anal_XXCnt = Anal_XXCnt + 1
                    Select Case Rec_Level
                              Case 1
                                    Select Case LenH(CarNum)
                                           Case 11
                                                    For Car_i = 1 To 16
                                                        Select Case Car_i
                                                               Case 1
                                                                    Car_Num_Str = "%%%%" & MidH(CarNum, 5, 7)
                                                               Case 2
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%" & MidH(CarNum, 6, 6)
                                                               Case 3
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%%" & MidH(CarNum, 8, 4)
                                                               Case 4
                                                                    Car_Num_Str = MidH(CarNum, 1, 7) & "%" & MidH(CarNum, 9, 3)
                                                               Case 5
                                                                    Car_Num_Str = MidH(CarNum, 1, 8) & "%" & MidH(CarNum, 10, 2)
                                                               Case 6
                                                                    Car_Num_Str = MidH(CarNum, 1, 9) & "%" & MidH(CarNum, 11, 1)
                                                               Case 7
                                                                    Car_Num_Str = MidH(CarNum, 1, 10) & "%"
                                                               Case 8
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 6)
                                                               Case 9 '���12��7890 : 12�� 2�� �ν��� �ȵȰ��(���ڴ���)
                                                                    Car_Num_Str = "%%%%" & MidH(CarNum, 5, 1) & "%" & MidH(CarNum, 6, 6)
                                                               Case 10
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%%" & MidH(CarNum, 6, 6)
                                                               Case 11
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%%%" & MidH(CarNum, 8, 4)
                                                               Case 12
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 2) & "%" & MidH(CarNum, 9, 3)
                                                               Case 13
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 3) & "%" & MidH(CarNum, 10, 2)
                                                               Case 14
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 4) & "%" & MidH(CarNum, 11, 1)
                                                               Case 15
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 6, 5) & "%"
                                                        End Select
                                                        Set JungRs = New ADODB.Recordset
                                                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & Car_Num_Str & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!������ȣ
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 12 '����ȣ2 : ����81��6800
                                                    For Car_i = 1 To 8
                                                        Select Case Car_i
                                                               Case 1
                                                                    Car_Num_Str = "%%%%" & MidH(CarNum, 5, 8)
                                                               Case 2
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%" & MidH(CarNum, 6, 7)
                                                               Case 3
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 7, 6)
                                                               Case 4
                                                                    Car_Num_Str = MidH(CarNum, 1, 6) & "%%" & MidH(CarNum, 9, 4)
                                                               Case 5
                                                                    Car_Num_Str = MidH(CarNum, 1, 8) & "%" & MidH(CarNum, 10, 3)
                                                               Case 6
                                                                    Car_Num_Str = MidH(CarNum, 1, 9) & "%" & MidH(CarNum, 11, 2)
                                                               Case 7
                                                                    Car_Num_Str = MidH(CarNum, 1, 10) & "%" & MidH(CarNum, 12, 1)
                                                               Case 8
                                                                    Car_Num_Str = MidH(CarNum, 1, 11) & "%"
                                                        End Select
                                                        Set JungRs = New ADODB.Recordset
                                                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & Car_Num_Str & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!������ȣ
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 8 '�Ź�ȣ1 : 81��7849
                                                    For Car_i = 1 To 7
                                                        Select Case Car_i
                                                               Case 1
                                                                    Car_Num_Str = "%" & MidH(CarNum, 2, 7)
                                                               Case 2
                                                                    Car_Num_Str = MidH(CarNum, 1, 1) & "%" & MidH(CarNum, 3, 6)
                                                               Case 3
                                                                    Car_Num_Str = MidH(CarNum, 1, 2) & "%%" & MidH(CarNum, 5, 4)
                                                               Case 4
                                                                    Car_Num_Str = MidH(CarNum, 1, 4) & "%" & MidH(CarNum, 6, 3)
                                                               Case 5
                                                                    Car_Num_Str = MidH(CarNum, 1, 5) & "%" & MidH(CarNum, 7, 2)
                                                               Case 6
                                                                    Car_Num_Str = MidH(CarNum, 1, 6) & "%" & MidH(CarNum, 8, 1)
                                                               Case 7
                                                                    Car_Num_Str = MidH(CarNum, 1, 7) & "%"
                                                        End Select
                                                        Set JungRs = New ADODB.Recordset
                                                        Qry = "SELECT * FROM regcar WHERE ������ȣ  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=" & Car_Num_Str & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻����=" & JungRs!������ȣ
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!������ȣ
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = Car_Num_Str
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����1 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(Car_Num_Str, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                    End Select
                              Case 2
                                        Set JungRs = New ADODB.Recordset
                                        Qry = "SELECT * FROM regcar WHERE (len(������ȣ ) = " & LenH(CarNum) & ") AND (������ȣ  Like '%" & RightH(CarNum, 6) & "')"
                                        JungRs.Open Qry, AdoConn
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����2 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����2 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!������ȣ
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����2 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   '.Text2.Text = RightH(CarNum, 6)
                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����2 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����2 >> �ߺ����ڵ� >> ���ܱ� �������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 6) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                              Case 3
                                        Set JungRs = New ADODB.Recordset
                                        Qry = "SELECT * FROM regcar WHERE (len(������ȣ ) = " & LenH(CarNum) & ") AND (������ȣ  Like '%" & RightH(CarNum, 4) & "')"
                                        JungRs.Open Qry, AdoConn
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    RecStat = "���ν�"
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����3 >> ���Ϸ��ڵ� >> �ڵ�ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����3 >> ���Ϸ��ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻����=" & JungRs!������ȣ
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!������ȣ
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����3 >> �ߺ����ڵ� >> ����ó��] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = RightH(CarNum, 4)
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����3 >> �ߺ����ڵ� >> ���ܱ� �ڵ�����] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ [���ν� ���͸�ó�� >> ���͸�����3 >> �ߺ����ڵ� >> ó������] �νĹ�ȣ=" & Save_CarNum & " ,���͸�����=*" & RightH(CarNum, 4) & " ,�˻��Ǽ�=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                    End Select
                    '.SSPanel3(4).ForeColor = vbBlue
                    RecStat = "���ν�"
                Else
                    '.SSPanel3(4).ForeColor = vbWhite
                    RecStat = "���ν�"
                End If
                .LblRecStat(Comm_Port * 2 + 1).Caption = "�νİ�� : �̵��"
            End If
    End If
    Select Case RecStat
           Case "�νĽ���"
                .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & "�νĽ���"
                .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�νĽ���"
                .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�νĽ���"
                .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �νĽ���"
                Dir_File = Dir(Out_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name Out_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                
                AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & "�νĽ���" & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "�νĽ���" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�νĽ��� �����Դϴ�. ]"
                If (.Check1.Value = 1) Then
                    .List1.ListIndex = .List1.ListCount - 1
                End If
           Case "���ν�", "�κ��ν�"
                Dir_File = Dir(Out_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name Out_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                If (RecStat = "�κ��ν�") Then
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & JungRs!������ȣ
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & JungRs!�̸�
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & JungRs!�Ҽ�
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �����ߺ�"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�����ߺ�', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�����ߺ� �߻�! ��������ó�� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If

                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & "�̵��"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �̵��"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        Set JungIORs = New ADODB.Recordset
                        Qry = "SELECT * FROM  regcarinout WHERE (�νĹ�ȣ='" & Save_CarNum & "') AND (���ⱸ��=False)"
                        JungIORs.Open Qry, AdoConn
                        If Not (JungIORs.EOF) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & Save_CarNum
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�Ϲݱ�"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�Ϲݱ�"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �Ϲݱ�"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�Ϲݱ� �����Դϴ�. ������ȣ :  " & Save_CarNum & "    �����Ͻ� : " & JungIORs!�������� & " " & JungIORs!�����ð� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                            AdoConn.Execute "DELETE FROM  regcarinout WHERE (�νĹ�ȣ='" & Save_CarNum & "') AND (���ⱸ��=False)"
                            
                            'If Not (UnRefresh_f) Then
                            '    .DataJungIn.Refresh
                            '    If (.DataJungIn.Recordset.BOF And .DataJungIn.Recordset.EOF) Then
                            '    Else
                            '        .DataJungIn.Recordset.MoveLast
                            '    End If
                            'End If
                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & "�̵��"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �̵��"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    End If
                Else
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & JungRs!������ȣ
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & JungRs!�̸�
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & JungRs!�Ҽ�
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �����ߺ�"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�����ߺ�', '" & JungRs!������ȣ & "', '" & JungRs!�̸� & "', '" & JungRs!������ & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!�Ҽ� & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�����ߺ� �߻�! ��������ó�� �����Դϴ�. " & "  ������ȣ : " & JungRs!������ȣ & "  �̸� : " & JungRs!�̸� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & "�̵��"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �̵��"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�̵��', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        Set JungIORs = New ADODB.Recordset
                        Qry = "SELECT * FROM  regcarinout WHERE (�νĹ�ȣ='" & Save_CarNum & "') AND (���ⱸ��=False)"
                        JungIORs.Open Qry, AdoConn
                        If Not (JungIORs.EOF) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & Save_CarNum
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�Ϲݱ�"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�Ϲݱ�"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �Ϲݱ�"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�Ϲݱ� �����Դϴ�. ������ȣ :  " & Save_CarNum & "    �����Ͻ� : " & JungIORs!�������� & " " & JungIORs!�����ð� & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                            AdoConn.Execute "DELETE FROM  regcarinout WHERE (�νĹ�ȣ='" & Save_CarNum & "') AND (���ⱸ��=False)"
                            'If Not (UnRefresh_f) Then
                            '    .DataJungIn.Refresh
                            '    If (.DataJungIn.Recordset.BOF And .DataJungIn.Recordset.EOF) Then
                            '    Else
                            '        .DataJungIn.Recordset.MoveLast
                            '    End If
                            'End If
                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "������ȣ : " & "�̵��"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "��      �� : " & "�̵��"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "ó����� : �̵��"
                            AdoConn.Execute "INSERT INTO regcarinout (��������, �����ð�, ��������, �����ð�, �������, ������ȣ, �̸�, ������, ó���Ͻ�, ���ⱸ��, �νĹ�ȣ, �Ҽ�, �νĻ���, �̹�����, ��ȭ��ȣ, ����Ʈ����) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '�Ϲݱ�', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "�̵��" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " ������ⱸ" & " [�̵�� �����Դϴ�. ������ȣ :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    End If
                End If
    End Select
End If

If (.Frame3.Visible = True) Then
    .Label3.Caption = "���������:" & Anal_InCnt & Space(6 - Len(Anal_InCnt)) & "�����ν�:" & Anal_OkCnt & Space(6 - Len(Anal_OkCnt)) & "�κ��ν�:" & Anal_HalfCnt & Space(6 - Len(Anal_HalfCnt)) & "�νĽ���:" & Anal_FailCnt & Space(6 - Len(Anal_FailCnt)) & "���ν�:" & Anal_XXCnt & Space(6 - Len(Anal_XXCnt)) & "�νķ� = " & Int((Anal_OkCnt / Anal_InCnt) * 100) & "%" & Space(3) & "������ = " & Int(((Anal_OkCnt + Anal_HalfCnt + Anal_XXCnt) / Anal_InCnt) * 100) & "%"
Put_Ini "�νķ��м�", "�������", CStr(Anal_InCnt)
Put_Ini "�νķ��м�", "�����ν�", CStr(Anal_OkCnt)
Put_Ini "�νķ��м�", "�κ��ν�", CStr(Anal_HalfCnt)
Put_Ini "�νķ��м�", "�νĽ���", CStr(Anal_FailCnt)
Put_Ini "�νķ��м�", "���ν�", CStr(Anal_XXCnt)
End If
End With

Set JungIORs = Nothing
Set JungRs = Nothing

Exit Sub
Err_Proc:
Report_Write 0, "���о���", "ȣ��Ʈ", "**********", "RemoteOut_Proc �����߻�!", True
Report_Write 0, "��⳻��", "ȣ��Ʈ", Err.Number, Err.Description, True
Call Err_doc("ȣ��Ʈ : RemoteOut_Proc �����߻�!  Error��ȣ = " & Err.Number & "       Error���� = " & Err.Description)

End Sub
