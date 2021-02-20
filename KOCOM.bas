Attribute VB_Name = "KOCOM"
Option Explicit

Public Kocom_Mode As String
Public Kocom_Data As String
Public Kocom_Cnt As String

' 코콤관련 변수
Public Type KocomHeader
    KEY(3) As Byte
    MSGTYPE(3) As Byte
    MSGLENGTH(3) As Byte
    TOWN(3) As Byte
    Dong(3) As Byte
    Ho(3) As Byte
    Reserved(3) As Byte
End Type
Public KHeader As KocomHeader

Public Type KocomBind
    HomeVersion(3) As Byte
    nKind(3) As Byte
    nVersion(15) As Byte
    SzId(39) As Byte
    SzPass(39) As Byte
End Type
Public KBind As KocomBind

Public Type KocomAlive
    HomeVersion(3) As Byte
    nKind(3) As Byte
    nVersion(15) As Byte
End Type
Public KAlive As KocomAlive

Public Type KocomInfo
    szGateId(3) As Byte
    szParkMan(3) As Byte
    szCardNo(39) As Byte
    nInOut(3) As Byte
    szDate(15) As Byte
    szCarNo(11) As Byte
End Type
Public KInfo As KocomInfo

'Public Sub Kocom(sCmd As String, sDATA As String)
'
'    Dim i                   As Integer
'    Dim bHversion(3)        As Byte
'    Dim bnKind(3)           As Byte
'    Dim bnVersion(15)       As Byte
'    Dim bSzid(39)           As Byte
'    Dim bSzPass(39)         As Byte
'    Dim sizeBody            As Long
'
'    Dim bSzGateid(3)        As Byte
'    Dim bSzParkMan(3)       As Byte
'    Dim bSzCardNo(39)       As Byte
'    Dim bnInOut(3)          As Byte
'    Dim bSzDate(15)         As Byte
'    Dim bSzCarNo(11)        As Byte
'    Dim sCarno              As String
'    Dim sDONG               As String
'    Dim sHO                 As String
'
'On Error GoTo Err
'
'    '데이터 구성
'    Select Case sCmd
'        Case "B"        'bind
'            For i = 0 To 3
'                bHversion(i) = &H0
'                bnKind(i) = &H0
'            Next
'
'            For i = 0 To 15
'                bnVersion(i) = &H0
'            Next
'
'            For i = 0 To 39
'                bSzid(i) = &H0
'                bSzPass(i) = &H0
'            Next
'
'            For i = 0 To Len(GsHouseID) - 1     '코콤에 등록된 ID
'                bSzid(i) = "&H" & Hex(Asc(Mid(GsHouseID, i + 1, 1)))
'            Next
'
'            For i = 0 To Len(GsHousePass) - 1       '코콤에 등록된 PassWord
'                bSzPass(i) = "&H" & Hex(Asc(Mid(GsHousePass, i + 1, 1)))
'            Next
'
'            KHeader.KEY(0) = &H78
'            KHeader.KEY(1) = &H56
'            KHeader.KEY(2) = &H34
'            KHeader.KEY(3) = &H12
'
'
'            KHeader.MSGTYPE(0) = &H0
'            KHeader.MSGTYPE(1) = &H0
'            KHeader.MSGTYPE(2) = &H0
'            KHeader.MSGTYPE(3) = &H11
'
'            KHeader.MSGLENGTH(0) = &H68
'            KHeader.MSGLENGTH(1) = &H0
'            KHeader.MSGLENGTH(2) = &H0
'            KHeader.MSGLENGTH(3) = &H0
'
'
'        Case "A"        'alive
'            For i = 0 To 3
'                bHversion(i) = &H0
'                bnKind(i) = &H0
'            Next
'
'            For i = 0 To 15
'                bnVersion(15) = &H0
'            Next
'
'            KHeader.MSGTYPE(0) = &H4
'            KHeader.MSGTYPE(1) = &H0
'            KHeader.MSGTYPE(2) = &H0
'            KHeader.MSGTYPE(3) = &H11
'
'            KHeader.MSGLENGTH(0) = &H18
'            KHeader.MSGLENGTH(1) = &H0
'            KHeader.MSGLENGTH(2) = &H0
'            KHeader.MSGLENGTH(3) = &H0
'
'        Case "P"        'park_info
'
'            For i = 0 To 3
'                bSzGateid(i) = &H30
'                bSzParkMan(i) = &H30
'            Next
'
'            For i = 0 To 15
'                bSzDate(i) = &H0
'            Next
'
'            For i = 0 To 39
'                bSzCardNo(i) = &H0
'            Next
'
'            For i = 1 To Len(GsSerial_New)
'                bSzCardNo(i - 1) = "&H" & Hex(Asc(Mid$(GsSerial_New, i, 1)))    'RF카드번호
'            Next
'
'            bnInOut(0) = &H1        'sDATA 의 1번째 바이트로 구분하지만 입차로 고정.
'            bnInOut(1) = &H0
'            bnInOut(2) = &H0
'            bnInOut(3) = &H0
'
'            For i = 13 To 0 Step -1
'                bSzDate(i) = "&H" & Hex(Asc(Mid$(GmsINTIME, i + 1, 1)))
'            Next
'
'            For i = 2 To Len(sDATA) '차량번호는 전국본호일시 '전국'을 붙여준다. 전국11가1234, 서울11가1234(자리수 고정)
'                sCarno = sCarno & Hex(Asc(Mid$(sDATA, i, 1)))
'            Next
'
'            bSzCarNo(0) = "&H" & Mid$(sCarno, 1, 2)
'            bSzCarNo(1) = "&H" & Mid$(sCarno, 3, 2)
'            bSzCarNo(2) = "&H" & Mid$(sCarno, 5, 2)
'            bSzCarNo(3) = "&H" & Mid$(sCarno, 7, 2)
'            bSzCarNo(4) = "&H" & Mid$(sCarno, 9, 2)
'            bSzCarNo(5) = "&H" & Mid$(sCarno, 11, 2)
'            bSzCarNo(6) = "&H" & Mid$(sCarno, 13, 2)
'            bSzCarNo(7) = "&H" & Mid$(sCarno, 15, 2)
'            bSzCarNo(8) = "&H" & Mid$(sCarno, 17, 2)
'            bSzCarNo(9) = "&H" & Mid$(sCarno, 19, 2)
'            bSzCarNo(10) = "&H" & Mid$(sCarno, 21, 2)
'            bSzCarNo(11) = "&H" & Mid$(sCarno, 23, 2)
'
'
'
'            With KHeader
'                '동
'                sDONG = Hex(GmsDONG)    '동정보 전역변수
'                If Len(sDONG) < 8 Then
'                    For i = 1 To 8 - Len(sDONG)
'                        sDONG = "0" & sDONG
'                    Next
'                End If
'
'                .Dong(0) = "&H" & Mid$(sDONG, 7, 2)
'                .Dong(1) = "&H" & Mid$(sDONG, 5, 2)
'                .Dong(2) = "&H" & Mid$(sDONG, 3, 2)
'                .Dong(3) = "&H" & Mid$(sDONG, 1, 2)
'
'                '호
'                sHO = Hex(GmsHO)    '호정보 전역변수
'                If Len(sHO) < 8 Then
'                    For i = 1 To 8 - Len(sHO)
'                        sHO = "0" & sHO
'                    Next
'                End If
'                .Ho(0) = "&H" & Mid$(sHO, 7, 2)
'                .Ho(1) = "&H" & Mid$(sHO, 5, 2)
'                .Ho(2) = "&H" & Mid$(sHO, 3, 2)
'                .Ho(3) = "&H" & Mid$(sHO, 1, 2)
'
'                .MSGTYPE(0) = &H6E
'                .MSGTYPE(1) = &H0
'                .MSGTYPE(2) = &H0
'                .MSGTYPE(3) = &H11
'
'                .MSGLENGTH(0) = &H50
'                .MSGLENGTH(1) = &H0
'                .MSGLENGTH(2) = &H0
'                .MSGLENGTH(3) = &H0
'
'            End With
'
'    End Select
'
'    '데이터 전송
'    If frmMain.Winsock1.State = 7 Then
'        With KHeader                  '헤더전송
'            frmMain.Winsock1.SendData .KEY()
'            frmMain.Winsock1.SendData .MSGTYPE()
'            frmMain.Winsock1.SendData .MSGLENGTH()
'            frmMain.Winsock1.SendData .TOWN()
'            frmMain.Winsock1.SendData .Dong()
'            frmMain.Winsock1.SendData .Ho()
'            frmMain.Winsock1.SendData .Reserved()
'        End With
'
'        Select Case sCmd                '커맨드별 데이터전송
'            Case "B"
'                frmMain.Winsock1.SendData bHversion()
'                frmMain.Winsock1.SendData bnKind()
'                frmMain.Winsock1.SendData bnVersion()
'                frmMain.Winsock1.SendData bSzid()
'                frmMain.Winsock1.SendData bSzPass()
'            Case "A"
'                frmMain.Winsock1.SendData bHversion()
'                frmMain.Winsock1.SendData bnKind()
'                frmMain.Winsock1.SendData bnVersion()
'            Case "P"
'                frmMain.Winsock1.SendData bSzGateid()
'                frmMain.Winsock1.SendData bSzParkMan()
'                frmMain.Winsock1.SendData bSzCardNo()
'                frmMain.Winsock1.SendData bnInOut()
'                frmMain.Winsock1.SendData bSzDate()
'                frmMain.Winsock1.SendData bSzCarNo()
'
'        End Select
'
'    End If
'
'    Exit Sub
'
'Err:
'    RF_LOG_ERR "kocom : " & Err.Description
'
'End Sub


'##입차데이터전송 예###
''입차데이터
'Call Kocom("P", "1" & GmsCAR)   '1=입차
'
''BIND:소켓연결완료시 전송
'Call Kocom("B", "")

Public Sub Kocom_BIND(ID As String, PW As String)
Dim i As Integer

    Kocom_Mode = "BIND"

With KBind
    For i = 0 To 3
        .HomeVersion(i) = &H0
        .nKind(i) = &H0
    Next i
    For i = 0 To 15
        .nVersion(i) = &H0
    Next i
    
    For i = 0 To 39
        .SzId(i) = &H0
        .SzPass(i) = &H0
    Next i
    
    For i = 0 To Len(ID) - 1
        .SzId(i) = "&H" & Hex(Asc(Mid(ID, i + 1, 1)))
    Next
    For i = 0 To Len(PW) - 1
        .SzPass(i) = "&H" & Hex(Asc(Mid(PW, i + 1, 1)))
    Next
End With
    
With KHeader
    .KEY(0) = &H78
    .KEY(1) = &H56
    .KEY(2) = &H34
    .KEY(3) = &H12

    .MSGTYPE(0) = &H0
    .MSGTYPE(1) = &H0
    .MSGTYPE(2) = &H0
    .MSGTYPE(3) = &H11
    
    .MSGLENGTH(0) = &H68
    .MSGLENGTH(1) = &H0
    .MSGLENGTH(2) = &H0
    .MSGLENGTH(3) = &H0
End With
    
    Call Socket_Connect_Kocom

End Sub

Public Sub Kocom_ALIVE()
Dim i As Integer

On Error GoTo Err_P

    Kocom_Mode = "ALIVE"

With KAlive
    For i = 0 To 3
        .HomeVersion(i) = &H0
        .nKind(i) = &H0
    Next i
    For i = 0 To 15
        .nVersion(i) = &H0
    Next i
End With
    
With KHeader
    .KEY(0) = &H78
    .KEY(1) = &H56
    .KEY(2) = &H34
    .KEY(3) = &H12
    
    .MSGTYPE(0) = &H4
    .MSGTYPE(1) = &H0
    .MSGTYPE(2) = &H0
    .MSGTYPE(3) = &H11
    
    .MSGLENGTH(0) = &H18
    .MSGLENGTH(1) = &H0
    .MSGLENGTH(2) = &H0
    .MSGLENGTH(3) = &H0
End With
    
    If FrmTcpServer.Winsock_Kocom.State = 7 Then
        With KHeader     '헤더전송
            FrmTcpServer.Winsock_Kocom.SendData .KEY()
            FrmTcpServer.Winsock_Kocom.SendData .MSGTYPE()
            FrmTcpServer.Winsock_Kocom.SendData .MSGLENGTH()
            FrmTcpServer.Winsock_Kocom.SendData .TOWN()
            FrmTcpServer.Winsock_Kocom.SendData .Dong()
            FrmTcpServer.Winsock_Kocom.SendData .Ho()
            FrmTcpServer.Winsock_Kocom.SendData .Reserved()
        End With
        With KAlive
            FrmTcpServer.Winsock_Kocom.SendData .HomeVersion()
            FrmTcpServer.Winsock_Kocom.SendData .nKind()
            FrmTcpServer.Winsock_Kocom.SendData .nVersion()
        End With
    Else
        Kocom_Mode = ""
        Call HomeLogger(" [Kocom_Alive] Alive Check ==> BIND Failure..!!")
        Call Kocom_BIND(Homesvr_ID, Homesvr_PW)
    End If

Exit Sub

Err_P:
     Call HomeLogger("[Kocom_ALIVE Proc] :" & Err.Description)

End Sub

Public Sub Kocom_Alarm(inout As Integer, CarNo As String, tmpDong As Integer, tmpHo As Integer)
Dim i As Integer
Dim DTime As String
Dim Car() As Byte
Dim sDong, sHo As String
Dim tmpstr As String * 12
    
    Kocom_Mode = "ALARM"

With KHeader
    .KEY(0) = &H78
    .KEY(1) = &H56
    .KEY(2) = &H34
    .KEY(3) = &H12

    .MSGTYPE(0) = &H6E
    .MSGTYPE(1) = &H0
    .MSGTYPE(2) = &H0
    .MSGTYPE(3) = &H11
    
    .MSGLENGTH(0) = &H50
    .MSGLENGTH(1) = &H0
    .MSGLENGTH(2) = &H0
    .MSGLENGTH(3) = &H0
    
    sDong = Hex(tmpDong)
    If Len(sDong) < 8 Then
        For i = 1 To 8 - Len(sDong)
            sDong = "0" & sDong
        Next
    End If
    .Dong(0) = "&H" & Mid$(sDong, 7, 2)
    .Dong(1) = "&H" & Mid$(sDong, 5, 2)
    .Dong(2) = "&H" & Mid$(sDong, 3, 2)
    .Dong(3) = "&H" & Mid$(sDong, 1, 2)

    sHo = Hex(tmpHo)
    If Len(sHo) < 8 Then
        For i = 1 To 8 - Len(sHo)
            sHo = "0" & sHo
        Next
    End If
    .Ho(0) = "&H" & Mid$(sHo, 7, 2)
    .Ho(1) = "&H" & Mid$(sHo, 5, 2)
    .Ho(2) = "&H" & Mid$(sHo, 3, 2)
    .Ho(3) = "&H" & Mid$(sHo, 1, 2)
End With

With KInfo
    For i = 0 To 3
        .szGateId(i) = &H0
        .szParkMan(i) = &H0
        .nInOut(i) = &H0
    Next i
    For i = 0 To 39
        .szCardNo(i) = &H0
    Next i

    If inout = "0" Then
        .nInOut(0) = &H1
    Else
        .nInOut(0) = &H2
    End If
    
    For i = 0 To 15
        .szDate(i) = &H0
    Next i
    DTime = ""
    DTime = Format(Now, "YYYYMMDDHHNNSS")
    For i = 0 To Len(DTime) - 1
        .szDate(i) = "&H" & Hex(Asc(Mid(DTime, i + 1, 1)))
        'Debug.Print .szDate(i)
    Next
    
    For i = 0 To 11
        .szCarNo(i) = &H0
    Next i

    tmpstr = CarNo
    Car() = StrConv(tmpstr, vbFromUnicode)
    For i = 0 To 11
        .szCarNo(i) = "&H" & Hex(Car(i))
        'Debug.Print "&H" & Hex(Car(i))
    Next
    'Debug.Print "------------ CarNo"

End With

    If FrmTcpServer.Winsock_Kocom.State = 7 Then
        With KHeader     '헤더전송
            FrmTcpServer.Winsock_Kocom.SendData .KEY()
            FrmTcpServer.Winsock_Kocom.SendData .MSGTYPE()
            FrmTcpServer.Winsock_Kocom.SendData .MSGLENGTH()
            FrmTcpServer.Winsock_Kocom.SendData .TOWN()
            FrmTcpServer.Winsock_Kocom.SendData .Dong()
            FrmTcpServer.Winsock_Kocom.SendData .Ho()
            FrmTcpServer.Winsock_Kocom.SendData .Reserved()
        End With
        With KInfo
            FrmTcpServer.Winsock_Kocom.SendData .szGateId()
            FrmTcpServer.Winsock_Kocom.SendData .szParkMan()
            FrmTcpServer.Winsock_Kocom.SendData .szCardNo()
            FrmTcpServer.Winsock_Kocom.SendData .nInOut()
            FrmTcpServer.Winsock_Kocom.SendData .szDate()
            FrmTcpServer.Winsock_Kocom.SendData .szCarNo()
        End With
        
        '성공하고 카운터 초기화
        Kocom_Cnt = 0
    Else
        Kocom_Mode = ""
        Call HomeLogger(" [Kocom_Alarm]     Alarm Fail : Discoonect.!!")
    End If

End Sub


Public Sub Socket_Connect_Kocom()
    Dim bData() As Byte
    
On Error GoTo Err_P
    
With FrmTcpServer
    If (.Winsock_Kocom.State <> sckClosed) Then
        .Winsock_Kocom.Close
    End If
    .Winsock_Kocom.Connect HomeSvr_IP, HomeSvr_Port
    Call HomeLogger(" [Socket_Kocom_Connection]" & " HomeSvr_IP : " & HomeSvr_IP & " HomeSvr_Port : " & HomeSvr_Port)
End With

Exit Sub

Err_P:
    Call HomeLogger(" [Socket_Kocom_Connect Proc] 에러내용 : " & Err.Description)

End Sub

