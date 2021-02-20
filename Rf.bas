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
t = t & "입구"


Save_CarNum = CarNum
Anal_InCnt = Anal_InCnt + 1
With main

idx = Comm_Port

.LblRecNum(Comm_Port * 2).Caption = "인식번호 : " & Save_CarNum
.LblTime(Comm_Port * 2).Caption = "처리일자 : " & Format(Now, "yy-mm-dd")
.LblEtc(Comm_Port * 2).Caption = "처리시간 : " & Format(Now, "hh:nn:ss")

Set JungRs = New ADODB.Recordset
Qry = "SELECT * FROM regcar WHERE 차량번호 ='" & Save_CarNum & "'"
JungRs.Open Qry, AdoConn

If Not (JungRs.EOF) Then
    .LblRecStat(Comm_Port * 2).Caption = "인식결과 : 완전인식"
    RecStat = "완전인식"
    Anal_OkCnt = Anal_OkCnt + 1
Seek_DB:
    Dir_File = Dir(In_Img_Folder(Comm_Port) & "*.jpg")
    If (Dir_File <> "") Then
        Name In_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
    End If
    .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & JungRs!차량번호
    .LblRecName(Comm_Port * 2).Caption = "이      름 : " & JungRs!이름
    .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & JungRs!소속
    If ((JungRs!시작일 > Format(Now, "yyyy-mm-dd")) Or (JungRs!종료일 < Format(Now, "yyyy-mm-dd"))) Then
        '----------------------------------------- 기간초과 처리 -------------------------------------------------------------
        .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 기간초과"
        AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
        "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '기간초과', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
        
        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [기간초과 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  만료일 : " & JungRs!종료일 & "]"
        If (.Check1.Value = 1) Then
            .List1.ListIndex = .List1.ListCount - 1
        End If
    Else
        Pass_Mode = JungRs!입구1
        Select Case Pass_Mode
               Case 0 '전일사용자
                    .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 정상입차"
                '----------------------------------------- 입차 처리 -------------------------------------------------------------
                        AdoConn.Execute "UPDATE regcar SET 입출차구분 = 1, 입차수 = 입차수 + 1, 입차일자 = " & "'" & Format(Now, "yyyy-mm-dd") & "', 입차시간 = " & "'" & Format(Now, "hh:nn") & "'" & "WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
                        AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                        "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상입차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [정상입차 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
               Case 1 '주간사용자
                        .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 주간사용"
                        If ((Format(Now, "hh:nn") >= Mid(Am_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Am_Time, 7, 5))) Then
                            JIO_Status = "주간정상"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상입차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [주간사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "주간위반"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '출입금지', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [주간사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 2 '야간사용자
                        .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 야간사용"
                        If ((Format(Now, "hh:nn") >= Mid(Pm_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Pm_Time, 7, 5))) Then
                            JIO_Status = "야간정상"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상입차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [야간사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "야간위반"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '출입금지', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [야간사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 3 '주말사용자
                        .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 주말사용"
                        
                        If (Weekday(Now) = 7) Then
                            JIO_Status = "주말정상"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상입차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [주말사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "주말위반"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!출차일자 & "', '" & JungRs!출차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '출입금지', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [주말사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
        End Select
        For i = 0 To 9
            If (IsNull(JungRs(30 + (i * 2)).Value) Or (JungRs(30 + (i * 2)).Value = "")) Then
                AdoConn.Execute "UPDATE regcar SET 인식번호" & CStr(i + 1) & "='" & Save_CarNum & "', 처리결과" & CStr(i + 1) & "='" & RecStat & "' WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
                Exit For
            End If
        Next i
        If (i = 10) Then
            For i = 10 To 2 Step -1
                AdoConn.Execute "UPDATE regcar SET 인식번호" & CStr(i) & "=인식번호" & CStr(i - 1) & ", 처리결과" & CStr(i) & "=처리결과" & CStr(i - 1) & " WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
            Next i
            AdoConn.Execute "UPDATE regcar SET 인식번호1='" & Save_CarNum & "', 처리결과1='" & RecStat & "' WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
        End If
    End If
Else
    If (CarNum = "인식실패") Then
        '////////////////////// 인식실패 //////////////////////////////////////
            .LblRecStat(Comm_Port * 2).Caption = "인식결과 : 인식실패"
            RecStat = "인식실패"
            Anal_FailCnt = Anal_FailCnt + 1
    Else
            Q_Cnt = IsChar(CarNum)
            If (Q_Cnt > 0) Then
                RecStat = "부분인식"
                Anal_HalfCnt = Anal_HalfCnt + 1
                CarNum = XToQ(CarNum)
                If (Half_Rec_Mode) Then
                    '////////////////////// 부분인식 //////////////////////////////////////
                    RecStat = "부분인식"
                    .LblRecStat(Comm_Port * 2).Caption = "인식결과 : 부분인식"
                    If (Q_Cnt <= Half_Cnt) Then
                       If ((LenH(CarNum) = 7) Or (LenH(CarNum) = 8)) Then
                            CarNum = QToAll(CarNum)
                       Else
                            '경기12가1234 , 경기1가1234 , 12가1234
                             If (MidH(CarNum, 1, 1) = "%") Then
                                 CarNum = "%%%" & CarNum
                             End If
                             
                             Car_Num_Str = RightH(CarNum, 5)
                             If (MidH(Car_Num_Str, 1, 1) = "%") Then
                                 CarNum = MidH(CarNum, 1, (LenH(CarNum) - 5)) & "%%" & RightH(CarNum, 4)
                             End If
                       End If
                        Set JungRs = New ADODB.Recordset
                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & CarNum & "'"
                        JungRs.Open Qry, AdoConn
                        If Not (JungRs.EOF) Then
Record_Found:
                            JungRs.MoveLast
                            If (JungRs.RecordCount > 1) Then
                                    Select Case Half_Rec_DoubleRecord_Process
                                              Case "M"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [부분인식 필터링처리 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색건수=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                        'Call Data_ReSearch(LenH(CarNum))
                                                        Exit Sub
                                              Case "Y"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [부분인식 필터링처리 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색건수=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                       'Call DRelay(Comm_Port, 0)
                                              Case "N"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [부분인식 필터링처리 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색건수=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                    End Select
                            Else
                                If (Half_Rec_OneRecord_Process) Then
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [부분인식 필터링처리 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색결과=" & JungRs!차량번호
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    GoTo Seek_DB
                                Else
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [부분인식 필터링처리 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색결과=" & JungRs!차량번호
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    '.Text2.Text = JungRs!차량번호
                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                    Exit Sub
                                End If
                            End If
                        Else
                            If (LenH(CarNum) = 11) Then '경기12가7890 에서  경기1가7890  "2"가 소실된경우 경기1?가7890  로 치환하여 재검색한다
                                CarNum = MidH(CarNum, 1, 5) & "%" & RightH(CarNum, 6)
                                Set JungRs = New ADODB.Recordset
                                Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & CarNum & "'"
                                JungRs.Open Qry, AdoConn
                                If Not (JungRs.EOF) Then
                                    GoTo Record_Found
                                End If
                            End If
                        End If
                    Else
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [부분인식 필터링 문자제한(" & Half_Cnt & ") 초과] 인식번호=" & Save_CarNum
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                End If
            Else '////////////////////// 미등록 또는 오인식 //////////////////////////////////////
                If (No_Rec_Mode) Then
                    .LblRecStat(Comm_Port * 2).Caption = "인식결과 : 오인식"
                    RecStat = "오인식"
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
                                                               Case 9 '경기12가7890 : 12중 2가 인식이 안된경우(문자누락)
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
                                                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & Car_Num_Str & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!차량번호
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 12 '구번호2 : 서울81나6800
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
                                                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & Car_Num_Str & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!차량번호
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 8 '신번호1 : 81마7849
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
                                                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & Car_Num_Str & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!차량번호
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = Car_Num_Str
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
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
                                        Qry = "SELECT * FROM regcar WHERE (len(차량번호 ) = " & LenH(CarNum) & ") AND (차량번호  Like '%" & RightH(CarNum, 6) & "')"
                                        JungRs.Open Qry, AdoConn
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨2 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨2 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!차량번호
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨2 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   '.Text2.Text = RightH(CarNum, 6)
                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨2 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨2 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                              Case 3
                                        Set JungRs = New ADODB.Recordset
                                        Qry = "SELECT * FROM regcar WHERE (len(차량번호 ) = " & LenH(CarNum) & ") AND (차량번호  Like '%" & RightH(CarNum, 4) & "')"
                                        JungRs.Open Qry, AdoConn
                                        
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    RecStat = "오인식"
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨3 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨3 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!차량번호
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨3 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = RightH(CarNum, 4)
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨3 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구 [오인식 필터링처리 >> 필터링레벨3 >> 중복레코드 >> 처리안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                    End Select
                    '.SSPanel3(4).ForeColor = vbBlue
                    RecStat = "오인식"
                Else
                    '.SSPanel3(4).ForeColor = vbWhite
                    RecStat = "오인식"
                End If
            .LblRecStat(Comm_Port * 2).Caption = "인식결과 : 미등록"
            End If
    End If
    Select Case RecStat
           Case "인식실패"
                .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & "인식실패"
                .LblRecName(Comm_Port * 2).Caption = "이      름 : " & "인식실패"
                .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & "인식실패"
                .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 인식실패"
                Dir_File = Dir(In_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name In_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & "인식실패" & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "인식실패" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [인식실패 차량입니다. ]"
                If (.Check1.Value = 1) Then
                    .List1.ListIndex = .List1.ListCount - 1
                End If
           
           Case "오인식", "부분인식"
                Dir_File = Dir(In_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name In_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                If (RecStat = "부분인식") Then
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & JungRs!차량번호
                            .LblRecName(Comm_Port * 2).Caption = "이      름 : " & JungRs!이름
                            .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & JungRs!소속
                            .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 차번중복"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '차번중복', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [차번중복 발생! 정상입차처리 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & "미등록"
                            .LblRecName(Comm_Port * 2).Caption = "이      름 : " & "미등록"
                            .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & "미등록"
                            .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 미등록"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & "미등록"
                        .LblRecName(Comm_Port * 2).Caption = "이      름 : " & "미등록"
                        .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & "미등록"
                        .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 미등록"
                        AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                        "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                        
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                Else
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & JungRs!차량번호
                            .LblRecName(Comm_Port * 2).Caption = "이      름 : " & JungRs!이름
                            .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & JungRs!소속
                            .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 차번중복"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '차번중복', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [차번중복 발생! 정상입차처리 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                            
                        Else
                            .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & "미등록"
                            .LblRecName(Comm_Port * 2).Caption = "이      름 : " & "미등록"
                            .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & "미등록"
                            .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 미등록"
                            AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        .LblRecCarNum(Comm_Port * 2).Caption = "차량번호 : " & "미등록"
                        .LblRecName(Comm_Port * 2).Caption = "이      름 : " & "미등록"
                        .LblRecSosok(Comm_Port * 2).Caption = "소      속 : " & "미등록"
                        .LblRecProc(Comm_Port * 2).Caption = "처리결과 : 미등록"
                        AdoConn.Execute "INSERT INTO regcarinout (출차일자, 출차시간, 입차일자, 입차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                        "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '일반권', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 0, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권입구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                End If
    End Select
End If

If (.Frame3.Visible = True) Then
    .Label3.Caption = "총입차대수:" & Anal_InCnt & Space(6 - Len(Anal_InCnt)) & "완전인식:" & Anal_OkCnt & Space(6 - Len(Anal_OkCnt)) & "부분인식:" & Anal_HalfCnt & Space(6 - Len(Anal_HalfCnt)) & "인식실패:" & Anal_FailCnt & Space(6 - Len(Anal_FailCnt)) & "오인식:" & Anal_XXCnt & Space(6 - Len(Anal_XXCnt)) & "인식률 = " & Int((Anal_OkCnt / Anal_InCnt) * 100) & "%" & Space(3) & "가독률 = " & Int(((Anal_OkCnt + Anal_HalfCnt + Anal_XXCnt) / Anal_InCnt) * 100) & "%"
    Put_Ini "인식률분석", "입차대수", CStr(Anal_InCnt)
    Put_Ini "인식률분석", "완전인식", CStr(Anal_OkCnt)
    Put_Ini "인식률분석", "부분인식", CStr(Anal_HalfCnt)
    Put_Ini "인식률분석", "인식실패", CStr(Anal_FailCnt)
    Put_Ini "인식률분석", "오인식", CStr(Anal_XXCnt)
End If
End With
Set JungRs = Nothing
Exit Sub
Err_Proc:

Report_Write 0, "구분없음", "호스트", "**********", "RemoteIn_Proc 오류발생!", True
Report_Write 0, "상기내용", "호스트", Err.Number, Err.Description, True
Call Err_doc("호스트 : RemoteIn_Proc 오류발생!  Error번호 = " & Err.Number & "       Error내용 = " & Err.Description)
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
t = t & "출구"

Save_CarNum = CarNum
Anal_InCnt = Anal_InCnt + 1
With main

.LblRecNum(Comm_Port * 2 + 1).Caption = "인식번호 : " & Save_CarNum
.LblTime(Comm_Port * 2 + 1).Caption = "처리시간 : " & Format(Now, "yy-mm-dd")
.LblEtc(Comm_Port * 2 + 1).Caption = "처리시간 : " & Format(Now, "hh:nn:ss")


Set JungRs = New ADODB.Recordset
Qry = "SELECT * FROM regcar WHERE 차량번호 ='" & Save_CarNum & "'"
JungRs.Open Qry, AdoConn



If Not (JungRs.EOF) Then
    .LblRecStat(Comm_Port * 2 + 1).Caption = "인식결과 : 완전인식"
    RecStat = "완전인식"
    Anal_OkCnt = Anal_OkCnt + 1
Seek_DB:
    Dir_File = Dir(Out_Img_Folder(Comm_Port) & "*.jpg")
    If (Dir_File <> "") Then
        Name Out_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
    End If
    .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & JungRs!차량번호
    .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & JungRs!이름
    .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & JungRs!소속
    If ((JungRs!시작일 > Format(Now, "yyyy-mm-dd")) Or (JungRs!종료일 < Format(Now, "yyyy-mm-dd"))) Then
        '----------------------------------------- 기간초과 처리 -------------------------------------------------------------
        .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 기간초과"
        AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
        "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '기간초과', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
        
        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [기간초과 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  만료일 : " & JungRs!종료일 & "]"
        If (.Check1.Value = 1) Then
            .List1.ListIndex = .List1.ListCount - 1
        End If
    Else
        Pass_Mode = JungRs!입구1
        Select Case Pass_Mode
               Case 0 '전일사용자
                    .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 정상출차"
                '----------------------------------------- 입차 처리 -------------------------------------------------------------
                        AdoConn.Execute "UPDATE regcar SET 입출차구분 = 2, 출차수 = 출차수 + 1, 출차일자 = " & "'" & Format(Now, "yyyy-mm-dd") & "', 출차시간 = " & "'" & Format(Now, "hh:nn") & "'" & "WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
                        AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                        "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상출차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [정상출차 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
               Case 1 '주간사용자
                        .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 주간사용"
                        If ((Format(Now, "hh:nn") >= Mid(Am_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Am_Time, 7, 5))) Then
                            JIO_Status = "주간정상"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상출차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [정상출차 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "주간위반"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '출입금지', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [주간사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 2 '야간사용자
                        .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 야간사용"
                        If ((Format(Now, "hh:nn") >= Mid(Pm_Time, 1, 5)) And (Format(Now, "hh:nn") <= Mid(Pm_Time, 7, 5))) Then
                            JIO_Status = "야간정상"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상출차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [정상출차 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "야간위반"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '출입금지', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [야간사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
               Case 3 '주말사용자
                        .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 주말사용"
                        If (Weekday(Now) = 7) Then
                            JIO_Status = "주말정상"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '정상출차', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [정상출차 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            JIO_Status = "주말위반"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & JungRs!입차일자 & "', '" & JungRs!입차시간 & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '출입금지', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & JungRs!전화번호 & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [주말사용 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
        End Select
        
        For i = 0 To 9
            If (IsNull(JungRs(30 + (i * 2)).Value) Or (JungRs(30 + (i * 2)).Value = "")) Then
                AdoConn.Execute "UPDATE regcar SET 인식번호" & CStr(i + 1) & "='" & Save_CarNum & "', 처리결과" & CStr(i + 1) & "='" & RecStat & "' WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
                Exit For
            End If
        Next i
        If (i = 10) Then
            For i = 10 To 2 Step -1
                AdoConn.Execute "UPDATE regcar SET 인식번호" & CStr(i) & "=인식번호" & CStr(i - 1) & ", 처리결과" & CStr(i) & "=처리결과" & CStr(i - 1) & " WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
            Next i
            AdoConn.Execute "UPDATE regcar SET 인식번호1='" & Save_CarNum & "', 처리결과1='" & RecStat & "' WHERE 차량번호= " & "'" & JungRs!차량번호 & "'"
        End If
    End If
Else
    If (CarNum = "인식실패") Then
        '////////////////////// 인식실패 //////////////////////////////////////
            .LblRecStat(Comm_Port * 2 + 1).Caption = "인식결과 : 인식실패"
            RecStat = "인식실패"
            Anal_FailCnt = Anal_FailCnt + 1
    Else
            Q_Cnt = IsChar(CarNum)
            If (Q_Cnt > 0) Then
                RecStat = "부분인식"
                Anal_HalfCnt = Anal_HalfCnt + 1
                CarNum = XToQ(CarNum)
                If (Half_Rec_Mode) Then
                    '////////////////////// 부분인식 //////////////////////////////////////
                    RecStat = "부분인식"
                    .LblRecStat(Comm_Port * 2 + 1).Caption = "인식결과 : 부분인식"
                    If (Q_Cnt <= Half_Cnt) Then
                       If ((LenH(CarNum) = 7) Or (LenH(CarNum) = 8)) Then
                            CarNum = QToAll(CarNum)
                       Else
                            '경기12가1234 , 경기1가1234 , 12가1234
                             If (MidH(CarNum, 1, 1) = "%") Then
                                 CarNum = "%%%" & CarNum
                             End If
                             
                             Car_Num_Str = RightH(CarNum, 5)
                             If (MidH(Car_Num_Str, 1, 1) = "%") Then
                                 CarNum = MidH(CarNum, 1, (LenH(CarNum) - 5)) & "%%" & RightH(CarNum, 4)
                             End If
                       End If
                        Set JungRs = New ADODB.Recordset
                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & CarNum & "'"
                        JungRs.Open Qry, AdoConn
                        If Not (JungRs.EOF) Then
Record_Found:
                            JungRs.MoveLast
                            If (JungRs.RecordCount > 1) Then
                                    Select Case Half_Rec_DoubleRecord_Process
                                              Case "M"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [부분인식 필터링처리 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색건수=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                        'Call Data_ReSearch(LenH(CarNum))
                                                        Exit Sub
                                              Case "Y"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [부분인식 필터링처리 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색건수=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                                       'Call DRelay(Comm_Port, 0)
                                              Case "N"
                                                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [부분인식 필터링처리 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색건수=" & JungRs.RecordCount
                                                        If (.Check1.Value = 1) Then
                                                            .List1.ListIndex = .List1.ListCount - 1
                                                        End If
                                    End Select
                            Else
                                If (Half_Rec_OneRecord_Process) Then
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [부분인식 필터링처리 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색결과=" & JungRs!차량번호
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    GoTo Seek_DB
                                Else
                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [부분인식 필터링처리 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & CarNum & " ,검색결과=" & JungRs!차량번호
                                    If (.Check1.Value = 1) Then
                                        .List1.ListIndex = .List1.ListCount - 1
                                    End If
                                    Exit Sub
                                End If
                            End If
                        Else
                            If (LenH(CarNum) = 11) Then '경기12가7890 에서  경기1가7890  "2"가 소실된경우 경기1?가7890  로 치환하여 재검색한다
                                CarNum = MidH(CarNum, 1, 5) & "%" & RightH(CarNum, 6)
                                Set JungRs = New ADODB.Recordset
                                Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & CarNum & "'"
                                JungRs.Open Qry, AdoConn
                                If Not (JungRs.EOF) Then
                                    GoTo Record_Found
                                End If
                            End If
                        End If
                    Else
                        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [부분인식 필터링 문자제한(" & Half_Cnt & ") 초과] 인식번호=" & Save_CarNum
                        If (.Check1.Value = 1) Then
                            .List1.ListIndex = .List1.ListCount - 1
                        End If
                    End If
                End If
            Else '////////////////////// 미등록 또는 오인식 //////////////////////////////////////
                If (No_Rec_Mode) Then
                    .LblRecStat(Comm_Port * 2 + 1).Caption = "인식결과 : 오인식"
                    RecStat = "오인식"
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
                                                               Case 9 '경기12가7890 : 12중 2가 인식이 안된경우(문자누락)
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
                                                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & Car_Num_Str & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!차량번호
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 12 '구번호2 : 서울81나6800
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
                                                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & Car_Num_Str & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!차량번호
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = RightH(Car_Num_Str, 6)
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                    Exit For
                                                                End Select
                                                            End If
                                                        End If
                                                   Next Car_i
                                           Case 8 '신번호1 : 81마7849
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
                                                        Qry = "SELECT * FROM regcar WHERE 차량번호  Like '" & Car_Num_Str & "'"
                                                        JungRs.Open Qry, AdoConn
                                                        If Not (JungRs.EOF) Then
                                                            JungRs.MoveLast
                                                            If (JungRs.RecordCount = 1) Then
                                                                If (No_Rec_OneRecord_Process) Then
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=" & Car_Num_Str & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    GoTo Seek_DB
                                                                Else
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색결과=" & JungRs!차량번호
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = JungRs!차량번호
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                                End If
                                                            Else
                                                                Select Case No_Rec_DoubleRecord_Process
                                                                          Case "M"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   '.Text2.Text = Car_Num_Str
                                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                                    Exit Sub
                                                                          Case "Y"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                                    If (.Check1.Value = 1) Then
                                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                                    End If
                                                                                   'Call DRelay(Comm_Port, 0)
                                                                                   Exit For
                                                                          Case "N"
                                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨1 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(Car_Num_Str, 6) & " ,검색건수=" & JungRs.RecordCount
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
                                        Qry = "SELECT * FROM regcar WHERE (len(차량번호 ) = " & LenH(CarNum) & ") AND (차량번호  Like '%" & RightH(CarNum, 6) & "')"
                                        JungRs.Open Qry, AdoConn
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨2 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨2 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!차량번호
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨2 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   '.Text2.Text = RightH(CarNum, 6)
                                                                    'Call Data_ReSearch(LenH(CarNum))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨2 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨2 >> 중복레코드 >> 차단기 개방안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 6) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                              Case 3
                                        Set JungRs = New ADODB.Recordset
                                        Qry = "SELECT * FROM regcar WHERE (len(차량번호 ) = " & LenH(CarNum) & ") AND (차량번호  Like '%" & RightH(CarNum, 4) & "')"
                                        JungRs.Open Qry, AdoConn
                                        If Not (JungRs.EOF) Then
                                            JungRs.MoveLast
                                            If (JungRs.RecordCount = 1) Then
                                                If (No_Rec_OneRecord_Process) Then
                                                    RecStat = "오인식"
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨3 >> 단일레코드 >> 자동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    GoTo Seek_DB
                                                Else
                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨3 >> 단일레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색결과=" & JungRs!차량번호
                                                    If (.Check1.Value = 1) Then
                                                        .List1.ListIndex = .List1.ListCount - 1
                                                    End If
                                                    '.Text2.Text = JungRs!차량번호
                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                    Exit Sub
                                                End If
                                            Else
                                                Select Case No_Rec_DoubleRecord_Process
                                                          Case "M"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨3 >> 중복레코드 >> 수동처리] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                    '.Text2.Text = RightH(CarNum, 4)
                                                                    'Call Data_ReSearch(LenH(.Text2.Text))
                                                                    Exit Sub
                                                          Case "Y"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨3 >> 중복레코드 >> 차단기 자동개방] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                                   'Call DRelay(Comm_Port, 0)
                                                          Case "N"
                                                                    .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구 [오인식 필터링처리 >> 필터링레벨3 >> 중복레코드 >> 처리안함] 인식번호=" & Save_CarNum & " ,필터링문자=*" & RightH(CarNum, 4) & " ,검색건수=" & JungRs.RecordCount
                                                                    If (.Check1.Value = 1) Then
                                                                        .List1.ListIndex = .List1.ListCount - 1
                                                                    End If
                                                End Select
                                            End If
                                        End If
                    End Select
                    '.SSPanel3(4).ForeColor = vbBlue
                    RecStat = "오인식"
                Else
                    '.SSPanel3(4).ForeColor = vbWhite
                    RecStat = "오인식"
                End If
                .LblRecStat(Comm_Port * 2 + 1).Caption = "인식결과 : 미등록"
            End If
    End If
    Select Case RecStat
           Case "인식실패"
                .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & "인식실패"
                .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "인식실패"
                .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "인식실패"
                .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 인식실패"
                Dir_File = Dir(Out_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name Out_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                
                AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & "인식실패" & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "인식실패" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [인식실패 차량입니다. ]"
                If (.Check1.Value = 1) Then
                    .List1.ListIndex = .List1.ListCount - 1
                End If
           Case "오인식", "부분인식"
                Dir_File = Dir(Out_Img_Folder(Comm_Port) & "*.jpg")
                If (Dir_File <> "") Then
                    Name Out_Img_Folder(Comm_Port) & FTP_File As "c:\Winpark\Image\" & RecStat & "\" & FTP_File
                End If
                If (RecStat = "부분인식") Then
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & JungRs!차량번호
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & JungRs!이름
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & JungRs!소속
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 차번중복"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '차번중복', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [차번중복 발생! 정상출차처리 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If

                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & "미등록"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "미등록"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "미등록"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 미등록"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        Set JungIORs = New ADODB.Recordset
                        Qry = "SELECT * FROM  regcarinout WHERE (인식번호='" & Save_CarNum & "') AND (입출구분=False)"
                        JungIORs.Open Qry, AdoConn
                        If Not (JungIORs.EOF) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & Save_CarNum
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "일반권"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "일반권"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 일반권"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [일반권 차량입니다. 차량번호 :  " & Save_CarNum & "    입차일시 : " & JungIORs!입차일자 & " " & JungIORs!입차시간 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                            AdoConn.Execute "DELETE FROM  regcarinout WHERE (인식번호='" & Save_CarNum & "') AND (입출구분=False)"
                            
                            'If Not (UnRefresh_f) Then
                            '    .DataJungIn.Refresh
                            '    If (.DataJungIn.Recordset.BOF And .DataJungIn.Recordset.EOF) Then
                            '    Else
                            '        .DataJungIn.Recordset.MoveLast
                            '    End If
                            'End If
                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & "미등록"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "미등록"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "미등록"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 미등록"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    End If
                Else
                    If (Half_Rec_DoubleRecord_Process = "Y") Then
                        If (JungRs.RecordCount >= 2) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & JungRs!차량번호
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & JungRs!이름
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & JungRs!소속
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 차번중복"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '차번중복', '" & JungRs!차량번호 & "', '" & JungRs!이름 & "', '" & JungRs!종료일 & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & JungRs!소속 & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [차번중복 발생! 정상출차처리 차량입니다. " & "  차량번호 : " & JungRs!차량번호 & "  이름 : " & JungRs!이름 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & "미등록"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "미등록"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "미등록"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 미등록"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '미등록', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    Else
                        Set JungIORs = New ADODB.Recordset
                        Qry = "SELECT * FROM  regcarinout WHERE (인식번호='" & Save_CarNum & "') AND (입출구분=False)"
                        JungIORs.Open Qry, AdoConn
                        If Not (JungIORs.EOF) Then
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & Save_CarNum
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "일반권"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "일반권"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 일반권"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [일반권 차량입니다. 차량번호 :  " & Save_CarNum & "    입차일시 : " & JungIORs!입차일자 & " " & JungIORs!입차시간 & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                            AdoConn.Execute "DELETE FROM  regcarinout WHERE (인식번호='" & Save_CarNum & "') AND (입출구분=False)"
                            'If Not (UnRefresh_f) Then
                            '    .DataJungIn.Refresh
                            '    If (.DataJungIn.Recordset.BOF And .DataJungIn.Recordset.EOF) Then
                            '    Else
                            '        .DataJungIn.Recordset.MoveLast
                            '    End If
                            'End If
                        Else
                            .LblRecCarNum(Comm_Port * 2 + 1).Caption = "차량번호 : " & "미등록"
                            .LblRecName(Comm_Port * 2 + 1).Caption = "이      름 : " & "미등록"
                            .LblRecSosok(Comm_Port * 2 + 1).Caption = "소      속 : " & "미등록"
                            .LblRecProc(Comm_Port * 2 + 1).Caption = "처리결과 : 미등록"
                            AdoConn.Execute "INSERT INTO regcarinout (입차일자, 입차시간, 출차일자, 출차시간, 입출상태, 차량번호, 이름, 종료일, 처리일시, 입출구분, 인식번호, 소속, 인식상태, 이미지명, 전화번호, 게이트구분) VALUES (" & _
                            "'" & "----------" & "', '" & "-----" & "', '" & Format(Now, "yyyy-mm-dd") & "', '" & Format(Now, "hh:nn") & "', '일반권', '" & CarNum & "', '" & "------------" & "', '" & "----------" & "', '" & Format(Now, "yyyymmddhhnnss") & "', 1, '" & Save_CarNum & "','" & "미등록" & "', '" & RecStat & "', '" & FTP_File & "', '" & "----------" & "', '" & t & "')"
                            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " 정기권출구" & " [미등록 차량입니다. 차량번호 :  " & CarNum & "]"
                            If (.Check1.Value = 1) Then
                                .List1.ListIndex = .List1.ListCount - 1
                            End If
                        End If
                    End If
                End If
    End Select
End If

If (.Frame3.Visible = True) Then
    .Label3.Caption = "총입차대수:" & Anal_InCnt & Space(6 - Len(Anal_InCnt)) & "완전인식:" & Anal_OkCnt & Space(6 - Len(Anal_OkCnt)) & "부분인식:" & Anal_HalfCnt & Space(6 - Len(Anal_HalfCnt)) & "인식실패:" & Anal_FailCnt & Space(6 - Len(Anal_FailCnt)) & "오인식:" & Anal_XXCnt & Space(6 - Len(Anal_XXCnt)) & "인식률 = " & Int((Anal_OkCnt / Anal_InCnt) * 100) & "%" & Space(3) & "가독률 = " & Int(((Anal_OkCnt + Anal_HalfCnt + Anal_XXCnt) / Anal_InCnt) * 100) & "%"
Put_Ini "인식률분석", "입차대수", CStr(Anal_InCnt)
Put_Ini "인식률분석", "완전인식", CStr(Anal_OkCnt)
Put_Ini "인식률분석", "부분인식", CStr(Anal_HalfCnt)
Put_Ini "인식률분석", "인식실패", CStr(Anal_FailCnt)
Put_Ini "인식률분석", "오인식", CStr(Anal_XXCnt)
End If
End With

Set JungIORs = Nothing
Set JungRs = Nothing

Exit Sub
Err_Proc:
Report_Write 0, "구분없음", "호스트", "**********", "RemoteOut_Proc 오류발생!", True
Report_Write 0, "상기내용", "호스트", Err.Number, Err.Description, True
Call Err_doc("호스트 : RemoteOut_Proc 오류발생!  Error번호 = " & Err.Number & "       Error내용 = " & Err.Description)

End Sub
