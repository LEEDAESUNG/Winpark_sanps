Attribute VB_Name = "GL_LED"
Option Explicit

Public Led_Show As Byte
Public Led_Speed As Byte
Public Led_StopTime As Byte
Public Led_Repeat As Byte
Public Led_up_color As Byte
Public Led_down_color As Byte

Public Nomal_Show As Byte
Public Nomal_Speed As Byte
Public Nomal_StopTime As Byte
Public Nomal_Up_color As Byte
Public Nomal_Down_color As Byte

Public Nomaltxt_Up As String
Public Nomaltxt_Down As String

Public Up_Color As Byte
Public Down_Color As Byte

Public Led_OutF_In As Byte
Public Led_OutF_Out As Byte

'Public M As String
Public C As String
Public ShowEffect As Byte
Public ShowSpeed As Byte
Public ShowTime As Byte
Public Repeat As Byte

Public In_Led_F As Boolean
Public Out_Led_F As Boolean

Public GloDisp_BData1() As Byte
Public GloDisp_BData1_Down() As Byte
Public GloDisp_BData2() As Byte
Public GloDisp_BData2_Down() As Byte
Public GloDisp_BData3() As Byte
Public GloDisp_BData3_Down() As Byte
Public GloDisp_BData4() As Byte
Public GloDisp_BData4_Down() As Byte
Public GloDisp_BData5() As Byte
Public GloDisp_BData5_Down() As Byte
Public GloDisp_BData6() As Byte
Public GloDisp_BData6_Down() As Byte
Public GlO_TcpDataGate As String

Public Glo_Emergency_F As Boolean
Public Glo_Emergency_MSG As Boolean

Public Gate_ACK(6) As Boolean
Public GateTimer_First(6) As Boolean
Public GlO_SendCnt(6) As Byte
Public GlO_GateRNum(6) As Byte


Public Sub GL_Nomal(D1 As String, D2 As String, Nomal_Show As Byte, Nomal_Speed As Byte, Nomal_StopTime As Byte, Nomal_Up_color As Byte, Nomal_Down_color As Byte, IN_OUT As Integer)
    Dim Header(16) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim Finish(1) As Byte
    Dim k1() As Byte
    Dim k2() As Byte
    Dim D() As Byte
    Dim First_Len As Integer
    Dim Second_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i As Integer
    Dim g As Integer
    Dim First_Str As String
    Dim Second_Str As String


    'If (Glo_LPRBoard = "위즈넷" Or Glo_LPRBoard = "자두이노") Then
        First_Len = LenH(D1)
        Second_Len = LenH(D2)
        If First_Len > Second_Len Then
            Bigger_Len = First_Len
        Else
            Bigger_Len = Second_Len
        End If
        
        i = Bigger_Len Mod 12
        
        If i = 0 Then
            
        Else
            Bigger_Len = Bigger_Len + (12 - i)
        End If
            
        Gap_Len = Bigger_Len - First_Len
        For g = 1 To Gap_Len
            D1 = D1 + " "
        Next g
        
        Gap_Len = Bigger_Len - Second_Len
        For g = 1 To Gap_Len
            D2 = D2 + " "
        Next g
        
        Bigger_Len = Bigger_Len - 1
        
        ReDim Color_Up(Bigger_Len) As Byte
        ReDim Color_Down(Bigger_Len) As Byte
        
        On Error GoTo Err_p
        
        Header(0) = &H10   'DLE
        Header(1) = &H2    'STX
        Header(2) = &H0    'DST
        Header(3) = &H0     'LEN
        Header(4) = ((Bigger_Len + 1) * 4 + 12)
        Header(5) = &H53    'CMD : 긴급 54 / 일반 53
        Header(6) = &H0     'Dummy
        Header(7) = &H0     'Dummy
        Header(8) = &H0     '저장매체 플래시롬
        Header(9) = &H91    ' (1001 0001) B[1:0] - 메인화면 폰트크기 16 font / B[5:4] - 화면표출 ON / B[6:7] - 문구 표출 방향 2 = 가로방향
        Header(10) = &H0    '모듈 분할 하지 않음
        Header(11) = &H0    'Dummy
        Header(12) = &H0    '분할화면 효과값 : 효과없슴
        Header(13) = Nomal_Show         '&H1    ' 메인화면 효과값 : 왼쪽이동
        Header(14) = Nomal_Speed        '&H1E   '효과 속도
        Header(15) = Nomal_StopTime     '&H0    '정지 시간 없음
        Header(16) = &H0    '세로 표출 위치 : 0 행
        Dim Up_Color As Byte
        Dim Down_Color As Byte
        
        Select Case Nomal_Up_color
            Case 0
                Up_Color = &H31 '녹
            Case 1
                Up_Color = &H32 '적
            Case 2
                Up_Color = &H33 '황
        End Select
                
        Select Case Nomal_Down_color
            Case 0
                Down_Color = &H31 '녹
            Case 1
                Down_Color = &H32 '적
            Case 2
                Down_Color = &H33 '황
        End Select
                
'        Up_Color = &H32
'        Down_Color = &H32
        For i = 0 To Bigger_Len
            Color_Up(i) = Up_Color
        Next i
        
        For i = 0 To Bigger_Len
            Color_Down(i) = Down_Color
        Next i
        Dim D_size1 As Integer
        Dim D_Size2 As Integer
        
        ReDim k1(Bigger_Len) As Byte
        ReDim k2(Bigger_Len) As Byte
        
        First_Str = D1
        D = StrConv(First_Str, vbFromUnicode)
        Bigger_Len = UBound(D)
        
        For i = 0 To (Bigger_Len)
            k1(i) = "&H" & Hex(D(i))
        Next i
        
        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        Bigger_Len = UBound(D)
        
        For i = 0 To (Bigger_Len)
            k2(i) = "&H" & Hex(D(i))
        Next i
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + UBound(Finish) + 5
        
        
        Select Case IN_OUT
        
            Case 0
                ReDim GloDisp_BData1(data_len) As Byte
                For i = 0 To UBound(Header)
                   GloDisp_BData1(i) = Header(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1(i + UBound(Header) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(Color_Down)
                    GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                Next i
                For i = 0 To UBound(k1)
                    GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                Next i
                For i = 0 To UBound(k2)
                    GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                Next i
            Case 1
                ReDim GloDisp_BData2(data_len) As Byte
                For i = 0 To UBound(Header)
                   GloDisp_BData2(i) = Header(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2(i + UBound(Header) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(Color_Down)
                    GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                Next i
                For i = 0 To UBound(k1)
                    GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                Next i
                For i = 0 To UBound(k2)
                    GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                Next i
            Case 2
                ReDim GloDisp_BData3(data_len) As Byte
                For i = 0 To UBound(Header)
                   GloDisp_BData3(i) = Header(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3(i + UBound(Header) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(Color_Down)
                    GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                Next i
                For i = 0 To UBound(k1)
                    GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                Next i
                For i = 0 To UBound(k2)
                    GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                Next i
            Case 3
                ReDim GloDisp_BData4(data_len) As Byte
                For i = 0 To UBound(Header)
                   GloDisp_BData4(i) = Header(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4(i + UBound(Header) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(Color_Down)
                    GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                Next i
                For i = 0 To UBound(k1)
                    GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                Next i
                For i = 0 To UBound(k2)
                    GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                Next i
            Case 4
                ReDim GloDisp_BData5(data_len) As Byte
                For i = 0 To UBound(Header)
                   GloDisp_BData5(i) = Header(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5(i + UBound(Header) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(Color_Down)
                    GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                Next i
                For i = 0 To UBound(k1)
                    GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                Next i
                For i = 0 To UBound(k2)
                    GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                Next i
            Case 5
                ReDim GloDisp_BData6(data_len) As Byte
                For i = 0 To UBound(Header)
                   GloDisp_BData6(i) = Header(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6(i + UBound(Header) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(Color_Down)
                    GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                Next i
                For i = 0 To UBound(k1)
                    GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                Next i
                For i = 0 To UBound(k2)
                    GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                Next i
        
        End Select
        
        
        
'    If (Glo_LPRBoard = "위즈넷") Then
    
        With FrmTcpServer
            Select Case IN_OUT
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With
'
'
'    ElseIf (Glo_LPRBoard = "자두이노") Then
'
'        Dim ip As String
'        Dim Port As Long
'
'        With FrmTcpServer
'            Select Case IN_OUT
'                Case 0
'                    ip = LANE1_DispIP:  Port = LANE1_DispPort
'                Case 1
'                    ip = LANE2_DispIP:  Port = LANE2_DispPort
'                Case 2
'                    ip = LANE3_DispIP:  Port = LANE3_DispPort
'                Case 3
'                    ip = LANE4_DispIP:  Port = LANE4_DispPort
'                Case 4
'                    ip = LANE5_DispIP:  Port = LANE5_DispPort
'                Case 5
'                    ip = LANE6_DispIP:  Port = LANE6_DispPort
'            End Select
'
'            If (.Disp1_sock(IN_OUT).State <> sckClosed) Then
'                .Disp1_sock(IN_OUT).Close
'            End If
'            .Disp1_sock(IN_OUT).Connect ip, Port
'            Call DataLogger("[DISP 접속]  시도 IP = " & ip & "    PORT = " & Port)
'
'        End With
        
        Call None_Delay_Time(0.1)
        
'    End If

Exit Sub

Err_p:



End Sub

Public Sub GL_Nomal_Horizontal(D1 As String, D2 As String, Nomal_Show As Byte, Nomal_Speed As Byte, Nomal_StopTime As Byte, Nomal_Up_color As Byte, Nomal_Down_color As Byte, IN_OUT As Integer, Normal_Shift As Byte)
    Dim Head_Up(21) As Byte
    Dim Head_Down(21) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim sHEX_Up() As Byte
    Dim sHEX_Down() As Byte
    Dim Finish(1) As Byte
    Dim D() As Byte
    Dim Up_Len As Integer
    Dim Down_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i As Integer
    Dim g As Integer
    Dim First_Str As String
    Dim Second_Str As String


'''        Up_Len = LenH(D1)
'''        Down_Len = LenH(D2)
'''        If (Up_Len > Down_Len) Then
'''            Bigger_Len = Up_Len
'''        Else
'''            Bigger_Len = Down_Len
'''        End If

'''        If Up_Len > Down_Len Then
'''            For g = 1 To (Up_Len - Down_Len)
'''                D2 = D2 + " "
'''            Next g
'''        Else
'''            For g = 1 To (Down_Len - Up_Len)
'''                D1 = D1 + " "
'''            Next g
'''        End If
'''
'        i = Bigger_Len Mod 12
'        If i = 0 Then
'        Else
'            Bigger_Len = Bigger_Len + (12 - i + 1)
'        End If

'        Up_Len = LenH(D1)
'        Down_Len = LenH(D2)
'        If (Up_Len > Down_Len) Then
'            Bigger_Len = Up_Len
'        Else
'            Bigger_Len = Down_Len
'        End If
        

        If (Normal_Shift = enumDISP_NML_SHIFT.eSTOP) Then
            D1 = LeftH(D1, Glo_DISP_COL * 2)
            D2 = LeftH(D2, Glo_DISP_COL * 2)
            Up_Len = LenH(D1)     '정지: 12문자만 출력
            Down_Len = LenH(D2)   '정지: 12문자만 출력
            
            If (Up_Len > Down_Len) Then
                Bigger_Len = Up_Len
            Else
                Bigger_Len = Down_Len
            End If
            
        ElseIf (Normal_Shift = enumDISP_NML_SHIFT.eSHIFT) Then
            Up_Len = LenH(D1)
            Down_Len = LenH(D2)
            If (Up_Len > Down_Len) Then
                Bigger_Len = Up_Len
            Else
                Bigger_Len = Down_Len
            End If
        
            If ((Bigger_Len Mod 12) = 0) Then
                Bigger_Len = Bigger_Len + 1
            Else
                Bigger_Len = Bigger_Len + 12 - (Bigger_Len Mod 12) + 1
            End If
        End If

        
        
        If Bigger_Len > Up_Len Then
            For g = 1 To (Bigger_Len - Up_Len)
                D1 = D1 + " "
            Next g
        End If
        If Bigger_Len > Down_Len Then
            For g = 1 To (Bigger_Len - Down_Len)
                D2 = D2 + " "
            Next g
        End If
        
        
        
        
        On Error GoTo Err_p
        
        Head_Up(0) = &H10    'DLE
        Head_Up(1) = &H2     'STX
        Head_Up(2) = &H0     'DST
        Head_Up(3) = &H0     'LEN
        Head_Up(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Up 속성 시작
        Head_Up(5) = &H94    '고정
        Head_Up(6) = &H1     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Up(7) = &H0     '※섹션번호
        Head_Up(8) = &H63    '표시제어(무한반복)
        Head_Up(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시, 00:현재 표시문구 종료 후 표시
        Head_Up(10) = &H0    '모듈 분할 하지 않음
        Head_Up(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
'        Head_Up(12) = &H6    '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
'        Head_Up(13) = &H6    '퇴장효과
        Head_Up(12) = Normal_Shift     '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
        Head_Up(13) = Normal_Shift     '퇴장효과
        
        Head_Up(14) = &H0    '보조효과:&H0, 사용하지 않음
        'Head_Up(15) = &H14               '효과속도:일반적으로 H14(20)으로 설정함
        'Head_Up(15) = &H1E                '효과속도:일반적인 속도보다 조금 느림 H1E(30)으로 설정, '임시주석
        Head_Up(15) = Nomal_Speed                '효과속도:일반적인 속도보다 조금 느림 H1E(30)으로 설정
        
        Head_Up(16) = &H0                '※유지시간:4초( 8 x 0.5초), ※섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함
        Head_Up(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(18) = &H0                'Y축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(19) = &H18               'X축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(20) = &H4                'Y축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Up 속성 끝
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Down 속성 시작
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력
        
        Head_Down(5) = &H94    '고정
        Head_Down(6) = &H1     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Down(7) = &H1     '섹션번호(1)
        Head_Down(8) = &H63    '표시제어(무한반복)
        Head_Down(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시
        Head_Down(10) = &H0    '모듈 분할 하지 않음
        Head_Down(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        'Head_Down(12) = &H6    '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
        'Head_Down(13) = &H6    '퇴장효과
        Head_Down(12) = Normal_Shift     '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
        Head_Down(13) = Normal_Shift     '퇴장효과
        Head_Down(14) = &H0    '보조효과:&H0, 사용하지 않음
        'Head_Down(15) = &H14              '효과속도:일반적으로 H14(20)으로 설정함
        'Head_Down(15) = &H1E               '효과속도:일반적인 속도보다 조금 느림 H1E(30)으로 설정,
        Head_Down(15) = Nomal_Speed                '효과속도:일반적인 속도보다 조금 느림 H1E(30)으로 설정
        
        'Head_Down(16) = &H4                '유지시간:4초( 8 x 0.5초)
        Head_Down(16) = &H0                '유지시간:4초( 8 x 0.5초), 긴 문장의 경우 0으로 설정함
        Head_Down(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(18) = &H4                'Y축 시작점:0픽셀(섹션분리할 경우 사용함) : 16픽셀
        Head_Down(19) = &H18                'X축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(20) = &H8                'Y축 종료점:0픽셀(섹션분리할 경우 사용함) : 32픽셀
        Head_Down(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Down 속성 끝
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
                
        ReDim Color_Up(Bigger_Len - 1) As Byte
        ReDim Color_Down(Bigger_Len - 1) As Byte
        
        For i = 0 To UBound(Color_Up)
            Color_Up(i) = Nomal_Up_color
        Next i
        
        For i = 0 To UBound(Color_Down)
            Color_Down(i) = Nomal_Down_color
        Next i
        
        ReDim sHEX_Up(Bigger_Len - 1) As Byte
        ReDim sHEX_Down(Bigger_Len - 1) As Byte
        
        First_Str = D1
        D = StrConv(First_Str, vbFromUnicode)
        '윗줄(가로)
        For i = 0 To UBound(D)
            sHEX_Up(i) = "&H" & Hex(D(i))
        Next i
        
        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        '아랫줄(가로)
        For i = 0 To UBound(D)
            sHEX_Down(i) = "&H" & Hex(D(i))
        Next i
        
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1)
        Head_Up(4) = "&H" & Hex(data_len)   '데이터 길이
        Head_Down(4) = "&H" & Hex(data_len)
        
        '임시테스트
'        Dim strHex As String
'        strHex = ByteArrayToHex(Head_Up)
'        Debug.Print "Head_Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Up)
'        Debug.Print "Color Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Down)
'        Debug.Print "Color Dn:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Up)
'        Debug.Print "Data Up:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Down)
'        Debug.Print "Data Dn:" & strHex

    
        
        Select Case IN_OUT
        
            Case 0
                ReDim GloDisp_BData1(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData1_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData1(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i
                'Debug.Print "일반UP:" & ByteArrayToHex(GloDisp_BData1)
                ''''''''''
                For i = 0 To UBound(Head_Down)
                   GloDisp_BData1_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
                'Debug.Print "일반DN:" & ByteArrayToHex(GloDisp_BData1_Down)
            Case 1
                ReDim GloDisp_BData2(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData2_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData2(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData2_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i

            Case 2
                ReDim GloDisp_BData3(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData3_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData3(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData3_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 3
                ReDim GloDisp_BData4(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData4_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData4(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData4_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 4
                ReDim GloDisp_BData5(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData5_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData5(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData5_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 5
                ReDim GloDisp_BData6(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData6_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData6(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData6_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
        
        End Select
        
        'Dim strHex As String
        'strHex = ByteArrayToHex(GloDisp_BData1)
        'Debug.Print "GloDisp_BData1:" & strHex

        'strHex = ByteArrayToHex(GloDisp_BData1_Down)
        'Debug.Print "GloDisp_BData1_Down:" & strHex

        
        With FrmTcpServer
            Select Case IN_OUT
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)

Exit Sub

Err_p:



End Sub


'전광판 세로출력
Public Sub GL_Nomal_Vertical(D1 As String, D2 As String, Nomal_Show As Byte, Nomal_Speed As Byte, Nomal_StopTime As Byte, Nomal_Up_color As Byte, Nomal_Down_color As Byte, IN_OUT As Integer, Normal_Shift As Byte)
    Dim Head_Up(21) As Byte
    Dim Head_Down(21) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim sHEX_Up() As Byte
    Dim sHEX_Down() As Byte
    Dim Finish(1) As Byte
    Dim D() As Byte
    Dim Up_Len As Integer
    Dim Down_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i, j As Integer
    Dim g As Integer
    Dim First_Str As String
    Dim Second_Str As String
    Dim iAscCount As Integer
    Dim Up_Unicode(256) As Byte
    Dim DOWN_Unicode(256) As Byte
    Dim iUniIDX As Integer
    Dim iUp_Unicode_Len As Integer
    Dim iDown_Unicode_Len As Integer

'        '윗줄, 아랫줄 문자열 길이중에서 가장 긴 길이 찾기
'        Up_Len = Len(D1)
'        Down_Len = Len(D2)
'        If (Up_Len > Down_Len) Then
'            Bigger_Len = Up_Len * 2     '세로출력시 한 문자당 2BYTE 처리하므로 x2 로 계산
'        Else
'            Bigger_Len = Down_Len * 2   '세로출력시 한 문자당 2BYTE 처리하므로 x2 로 계산
'        End If
        
        
        If (Normal_Shift = enumDISP_NML_SHIFT.eSTOP) Then
            D1 = Left(D1, Glo_DISP_COL)
            D2 = Left(D2, Glo_DISP_COL)
            Up_Len = Len(D1)     '정지: 6문자만 출력
            Down_Len = Len(D2)   '정지: 6문자만 출력
            
            If (Up_Len > Down_Len) Then
                Bigger_Len = Up_Len * 2    '세로출력시 한 문자당 2BYTE 처리하므로 x2 로 계산
            Else
                Bigger_Len = Down_Len * 2  '세로출력시 한 문자당 2BYTE 처리하므로 x2 로 계산
            End If
        
        
        ElseIf (Normal_Shift = enumDISP_NML_SHIFT.eSHIFT) Then
            Up_Len = Len(D1)
            Down_Len = Len(D2)
            
            If (Up_Len > Down_Len) Then
                Bigger_Len = Up_Len * 2     '세로출력시 한 문자당 2BYTE 처리하므로 x2 로 계산
            Else
                Bigger_Len = Down_Len * 2   '세로출력시 한 문자당 2BYTE 처리하므로 x2 로 계산
            End If
            
            If ((Bigger_Len Mod Glo_DISP_COL) = 0) Then
                Bigger_Len = Bigger_Len + 1
            Else
                Bigger_Len = Bigger_Len + Glo_DISP_COL - (Bigger_Len Mod Glo_DISP_COL) + 1
            End If
        End If



        '윗줄, 아랫줄 문자열 길이 같게 만듬
        If Up_Len > Down_Len Then
            For g = 1 To (Up_Len - Down_Len)
                D2 = D2 + " "
            Next g
        Else
            For g = 1 To (Down_Len - Up_Len)
                D1 = D1 + " "
            Next g
        End If

'''        iAscCount = 0
'''        D = StrConv(D1, vbFromUnicode)
'''        For i = 0 To UBound(D)
'''            If (D(i) >= 32 And D(i) <= 126) Then '1byte 아스키문자
'''                iAscCount = iAscCount + 1
'''            End If
'''        Next
        
        
        '아스키문자 수 계산(유니코드로 변경할때 아스키문자 수 만큼 더 만들어야 함)
'        iAscCount = 0
'        If Up_Len > Down_Len Then
'            D = StrConv(D1, vbFromUnicode)
'        Else
'            D = StrConv(D2, vbFromUnicode)
'        End If
'        For i = 0 To UBound(D)
'            If (D(i) >= 32 And D(i) <= 126) Then '1byte 아스키문자
'                iAscCount = iAscCount + 1
'            End If
'        Next

        
        On Error GoTo Err_p
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Up 속성 시작
        Head_Up(0) = &H10    'DLE
        Head_Up(1) = &H2     'STX
        Head_Up(2) = &H0     'DST
        Head_Up(3) = &H0     'LEN
        Head_Up(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력
        
        'Up 속성 시작
        Head_Up(5) = &H94    '고정
        Head_Up(6) = &H1     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Up(7) = &H0     '※섹션번호
        Head_Up(8) = &H63    '표시제어(무한반복)
        Head_Up(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시, 00:현재 표시문구 종료 후 표시
        Head_Up(10) = &H0    '모듈 분할 하지 않음
        Head_Up(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        'Head_Up(12) = &H6    '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
        'Head_Up(13) = &H6    '퇴장효과
        Head_Up(12) = Normal_Shift
        Head_Up(13) = Normal_Shift

        
        Head_Up(14) = &H0    '보조효과:&H0, 사용하지 않음
        'Head_Up(15) = &H1E   '효과속도:일반적으로 H1E(30)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        'Head_Up(15) = &H14   '효과속도:일반적으로 H14(20)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        Head_Up(15) = Nomal_Speed  '효과속도:일반적으로 H14(20)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        
        Head_Up(16) = &H0                '※유지시간:4초( 8 x 0.5초), ※섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함(페이지메세지에서는 의미없는 듯함)
        Head_Up(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(18) = &H0                'Y축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(19) = &H18               'X축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(20) = &H4                'Y축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Up 속성 끝
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Down 속성 시작
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력
        
        Head_Down(5) = &H94    '고정
        Head_Down(6) = &H1     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Down(7) = &H1     '섹션번호(1)
        Head_Down(8) = &H63    '표시제어(무한반복)
        Head_Down(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시
        Head_Down(10) = &H0    '모듈 분할 하지 않음
        Head_Down(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        
'        Head_Down(12) = &H6    '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
'        Head_Down(13) = &H6    '퇴장효과
        Head_Down(12) = Normal_Shift
        Head_Down(13) = Normal_Shift
        
        Head_Down(14) = &H0    '보조효과:&H0, 사용하지 않음
        'Head_Down(15) = &H1E   '효과속도:일반적으로 H1E(30)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        'Head_Down(15) = &H14   '효과속도:일반적으로 H14(20)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        Head_Down(15) = Nomal_Speed   '효과속도:일반적으로 H14(20)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        
        Head_Down(16) = &H0                '유지시간:4초( 8 x 0.5초), ※섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함(페이지메세지에서는 의미없는 듯함)
        Head_Down(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(18) = &H4                'Y축 시작점:0픽셀(섹션분리할 경우 사용함) : 16픽셀
        Head_Down(19) = &H18                'X축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(20) = &H8                'Y축 종료점:0픽셀(섹션분리할 경우 사용함) : 32픽셀
        Head_Down(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Down 속성 끝
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
'''        ReDim Color_Up(Bigger_Len - 1 + iAscCount) As Byte
'''        ReDim Color_Down(Bigger_Len - 1 + iAscCount) As Byte
        ReDim Color_Up(Bigger_Len - 1) As Byte
        ReDim Color_Down(Bigger_Len - 1) As Byte
        
        For i = 0 To UBound(Color_Up)
            'Debug.Print i
            Color_Up(i) = Nomal_Up_color + 8 '세로출력위해 +8
            
        Next i
        
        For i = 0 To UBound(Color_Down)
            Color_Down(i) = Nomal_Down_color + 8 '세로출력위해 +8
        Next i
        
'''        ReDim sHEX_Up(Bigger_Len - 1 + iAscCount) As Byte
'''        ReDim sHEX_Down(Bigger_Len - 1 + iAscCount) As Byte
        ReDim sHEX_Up(Bigger_Len - 1) As Byte
        ReDim sHEX_Down(Bigger_Len - 1) As Byte
        
        First_Str = D1
        D = StrConv(First_Str, vbFromUnicode)
        j = 0
        For i = 0 To UBound(D)
            If (i = 20) Then
                Debug.Print i
            End If
            If (D(i) >= 32 And D(i) <= 126) Then
                sHEX_Up(j) = "&HE0"
                j = j + 1
                sHEX_Up(j) = "&H" & Hex(D(i))
            Else
                sHEX_Up(j) = "&H" & Hex(D(i))
            End If
            j = j + 1
        Next i

        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        j = 0
        For i = 0 To UBound(D)
            If (D(i) >= 32 And D(i) <= 126) Then
                sHEX_Down(j) = "&HE0"
                j = j + 1
                sHEX_Down(j) = "&H" & Hex(D(i))
            Else
                sHEX_Down(j) = "&H" & Hex(D(i))
            End If
            j = j + 1
        Next i

        
        
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1) '-5 : 헤더에서
        Head_Up(4) = "&H" & Hex(data_len)   '데이터 길이
        Head_Down(4) = "&H" & Hex(data_len)
        
'        Dim strHex As String
'        strHex = ByteArrayToHex(Head_Up)
'        Debug.Print "Head_Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Up)
'        Debug.Print "Color Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Down)
'        Debug.Print "Color Dn:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Up)
'        Debug.Print "Data Up:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Down)
'        Debug.Print "Data Dn:" & strHex

    
        
        Select Case IN_OUT
        
            Case 0
                ReDim GloDisp_BData1(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData1_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData1(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i
                Debug.Print "UP:" & ByteArrayToHex(GloDisp_BData1) '임시테스트
                ''''''''''
                For i = 0 To UBound(Head_Down)
                   GloDisp_BData1_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
                Debug.Print "DN:" & ByteArrayToHex(GloDisp_BData1_Down) '임시테스트
            Case 1
                ReDim GloDisp_BData2(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData2_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData2(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData2_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i

            Case 2
                ReDim GloDisp_BData3(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData3_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData3(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData3_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 3
                ReDim GloDisp_BData4(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData4_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData4(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData4_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 4
                ReDim GloDisp_BData5(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData5_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData5(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData5_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 5
                ReDim GloDisp_BData6(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData6_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData6(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData6_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
        
        End Select
        
        
        
        With FrmTcpServer
            Select Case IN_OUT
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP TCP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP TCP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP TCP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP TCP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP TCP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP TCP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub


'긴급문구
'Led_StopTime:출력시간(10:5초, 20:10초..)
Public Sub GL_Emergency(D1 As String, D2 As String, Led_Show As Byte, Led_Speed As Byte, Led_StopTime As Byte, Led_Repeat As Byte, Led_up_color As Byte, Led_down_color As Byte, IN_OUT As Integer)
    
    If (Glo_Display = "전광판" Or Glo_Display = "전광판(풀컬러)") Then
            
        Dim Header(16) As Byte
        Dim Color_Up() As Byte
        Dim Color_Down() As Byte
        Dim Finish(1) As Byte
        Dim k1() As Byte
        Dim k2() As Byte
        Dim D() As Byte
        Dim First_Len As Integer
        Dim Second_Len As Integer
        Dim Bigger_Len As Integer
        Dim Gap_Len As Integer
        Dim i, g As Integer
        Dim First_Str As String
        Dim Second_Str As String
        Dim Up_Color As Byte
        Dim Down_Color As Byte
        Dim D_size1 As Integer
        Dim D_Size2 As Integer
    
        Dim carnum As String
        
        
        On Error GoTo Err_p
        
        
        
                    carnum = D1
        
                    First_Len = LenH(carnum)
                    Second_Len = LenH(D2)
        
                    If First_Len > Second_Len Then
                        Bigger_Len = First_Len
                    Else
                        Bigger_Len = Second_Len
                    End If
        
                    i = Bigger_Len Mod 12
                    If i = 0 Then
                    Else
                        Bigger_Len = Bigger_Len + (12 - i)
                    End If
                    Gap_Len = Bigger_Len - First_Len
                    For g = 1 To Gap_Len
                        carnum = carnum + " "
                    Next g
                    Gap_Len = Bigger_Len - Second_Len
                    For g = 1 To Gap_Len
                        D2 = D2 + " "
                    Next g
                    Bigger_Len = Bigger_Len - 1
        
                    ReDim Color_Up(Bigger_Len) As Byte
                    ReDim Color_Down(Bigger_Len) As Byte
        
                    Header(0) = &H10                'DLE
                    Header(1) = &H2                 'STX
                    Header(2) = &H0                  'DST
                    Header(3) = &H0                 'LEN
                    Header(4) = ((Bigger_Len + 1) * 4 + 12)
                    Header(5) = &H54                'CMD : 긴급 54 / 일반 53
                    Header(6) = &H0                 'Dummy
                    Header(7) = &H1                 '기존 메세지 삭제 후 표출
                    Header(8) = Led_Repeat          '반복 횟수
                    Header(9) = &H91                ' (1001 0001) B[1:0] - 메인화면 폰트크기 16 font / B[5:4] - 화면표출 ON / B[6:7] - 문구 표출 방향 2 = 가로방향
                    Header(10) = &H0                '모듈 분할 하지 않음
                    Header(11) = &H0                'Dummy
                    Header(12) = &H0                '분할화면 효과값 : 효과없슴
                    Header(13) = Led_Show           '&H1    ' 메인화면 효과값 : 왼쪽이동
                    Header(14) = Led_Speed          '&H1E   '효과 속도
                    Header(15) = Led_StopTime       '&H0    '정지 시간 없음
                    Header(16) = &H0                '세로 표출 위치 : 0 행
        
        
                    Select Case Led_up_color
                        Case 0
                            Up_Color = &H31
                        Case 1
                            Up_Color = &H32
                        Case 2
                            Up_Color = &H33
                        Case Else
                            Up_Color = &H32
                    End Select
                    Select Case Led_down_color
                        Case 0
                            Down_Color = &H31
                        Case 1
                            Down_Color = &H32
                        Case 2
                            Down_Color = &H33
                        Case Else
                            Down_Color = &H32
                    End Select
                    For i = 0 To Bigger_Len
                        Color_Up(i) = Up_Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색
                    Next i
                    For i = 0 To Bigger_Len
                        Color_Down(i) = Down_Color    '&H31 : 적색 / 32 : 녹색 / 33 : 노란색
                    Next i
        
                    ReDim k1(Bigger_Len) As Byte
                    ReDim k2(Bigger_Len) As Byte
        
                    First_Str = carnum
                    D = StrConv(First_Str, vbFromUnicode)
                    Bigger_Len = UBound(D)
                    For i = 0 To (Bigger_Len)
                        k1(i) = "&H" & Hex(D(i))
                    Next i
                    Second_Str = D2
                    D = StrConv(Second_Str, vbFromUnicode)
                    Bigger_Len = UBound(D)
                    For i = 0 To (Bigger_Len)
                        k2(i) = "&H" & Hex(D(i))
                    Next i
                    Finish(0) = &H10
                    Finish(1) = &H3
        
                    Dim data_len  As Integer
                    data_len = UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + UBound(Finish) + 5
        
        
        
        
                    Select Case IN_OUT
                        Case 0
                            ReDim GloDisp_BData1(data_len) As Byte
                            For i = 0 To UBound(Header)
                               GloDisp_BData1(i) = Header(i)
                            Next i
                            For i = 0 To UBound(Color_Up)
                                GloDisp_BData1(i + UBound(Header) + 1) = Color_Up(i)
                            Next i
                            For i = 0 To UBound(Color_Down)
                                GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                            Next i
                            For i = 0 To UBound(k1)
                                GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                            Next i
                            For i = 0 To UBound(k2)
                                GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                            Next i
                            For i = 0 To UBound(Finish)
                                GloDisp_BData1(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                            Next i
        
                        Case 1
                            ReDim GloDisp_BData2(data_len) As Byte
                            For i = 0 To UBound(Header)
                               GloDisp_BData2(i) = Header(i)
                            Next i
                            For i = 0 To UBound(Color_Up)
                                GloDisp_BData2(i + UBound(Header) + 1) = Color_Up(i)
                            Next i
                            For i = 0 To UBound(Color_Down)
                                GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                            Next i
                            For i = 0 To UBound(k1)
                                GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                            Next i
                            For i = 0 To UBound(k2)
                                GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                            Next i
                            For i = 0 To UBound(Finish)
                                GloDisp_BData2(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                            Next i
                        Case 2
                            ReDim GloDisp_BData3(data_len) As Byte
                            For i = 0 To UBound(Header)
                               GloDisp_BData3(i) = Header(i)
                            Next i
                            For i = 0 To UBound(Color_Up)
                                GloDisp_BData3(i + UBound(Header) + 1) = Color_Up(i)
                            Next i
                            For i = 0 To UBound(Color_Down)
                                GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                            Next i
                            For i = 0 To UBound(k1)
                                GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                            Next i
                            For i = 0 To UBound(k2)
                                GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                            Next i
                            For i = 0 To UBound(Finish)
                                GloDisp_BData3(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                            Next i
                        Case 3
                            ReDim GloDisp_BData4(data_len) As Byte
                            For i = 0 To UBound(Header)
                               GloDisp_BData4(i) = Header(i)
                            Next i
                            For i = 0 To UBound(Color_Up)
                                GloDisp_BData4(i + UBound(Header) + 1) = Color_Up(i)
                            Next i
                            For i = 0 To UBound(Color_Down)
                                GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                            Next i
                            For i = 0 To UBound(k1)
                                GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                            Next i
                            For i = 0 To UBound(k2)
                                GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                            Next i
                            For i = 0 To UBound(Finish)
                                GloDisp_BData4(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                            Next i
                        Case 4
                            ReDim GloDisp_BData5(data_len) As Byte
                            For i = 0 To UBound(Header)
                               GloDisp_BData5(i) = Header(i)
                            Next i
                            For i = 0 To UBound(Color_Up)
                                GloDisp_BData5(i + UBound(Header) + 1) = Color_Up(i)
                            Next i
                            For i = 0 To UBound(Color_Down)
                                GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                            Next i
                            For i = 0 To UBound(k1)
                                GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                            Next i
                            For i = 0 To UBound(k2)
                                GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                            Next i
                            For i = 0 To UBound(Finish)
                                GloDisp_BData5(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                            Next i
                        Case 5
                            ReDim GloDisp_BData6(data_len) As Byte
                            For i = 0 To UBound(Header)
                               GloDisp_BData6(i) = Header(i)
                            Next i
                            For i = 0 To UBound(Color_Up)
                                GloDisp_BData6(i + UBound(Header) + 1) = Color_Up(i)
                            Next i
                            For i = 0 To UBound(Color_Down)
                                GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + 2) = Color_Down(i)
                            Next i
                            For i = 0 To UBound(k1)
                                GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + 3) = k1(i)
                            Next i
                            For i = 0 To UBound(k2)
                                GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + 4) = k2(i)
                            Next i
                            For i = 0 To UBound(Finish)
                                GloDisp_BData6(i + UBound(Header) + UBound(Color_Up) + UBound(Color_Down) + UBound(k1) + UBound(k2) + 5) = Finish(i)
                            Next i
                    End Select
        
        
        
                With FrmTcpServer
                    Select Case IN_OUT
                        Case 0
                            Select Case LANE1_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(0).State <> sckClosed) Then
                                                .Disp1_sock(0).Close
                                            End If
                                            .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(0).SendData GloDisp_BData1
                                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                            End Select
        
                        Case 1
                            Select Case LANE2_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(1).State <> sckClosed) Then
                                                .Disp1_sock(1).Close
                                            End If
                                            .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(1).SendData GloDisp_BData2
                                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
        
                            End Select
        
                        Case 2
                             Select Case LANE3_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(2).State <> sckClosed) Then
                                                .Disp1_sock(2).Close
                                            End If
                                            .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(2).SendData GloDisp_BData3
                                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
        
                            End Select
        
                        Case 3
                            Select Case LANE4_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(3).State <> sckClosed) Then
                                                .Disp1_sock(3).Close
                                            End If
                                            .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(3).SendData GloDisp_BData4
                                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
        
                            End Select
        
                        Case 4
                            Select Case LANE5_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(4).State <> sckClosed) Then
                                                .Disp1_sock(4).Close
                                            End If
                                            .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(4).SendData GloDisp_BData5
                                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
        
                            End Select
        
                        Case 5
                            Select Case LANE6_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(5).State <> sckClosed) Then
                                                .Disp1_sock(5).Close
                                            End If
                                            .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(5).SendData GloDisp_BData6
                                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
        
                            End Select
                    End Select
                End With
        
        Exit Sub
Err_p:
    
        

    ElseIf (Glo_Display = "전광판(풀컬러)_FW7") Then
            If (Led_up_color = enumDIS_COLORs.eBLUE) Then
                Led_up_color = enumDIS_COLOR2s.eBLUE
            ElseIf (Led_up_color = enumDIS_COLORs.eGreen) Then
                Led_up_color = enumDIS_COLOR2s.eGreen
            ElseIf (Led_up_color = enumDIS_COLORs.eRED) Then
                Led_up_color = enumDIS_COLOR2s.eRED
            ElseIf (Led_up_color = enumDIS_COLORs.eSKY) Then
                Led_up_color = enumDIS_COLOR2s.eSKY
            ElseIf (Led_up_color = enumDIS_COLORs.eWHITE) Then
                Led_up_color = enumDIS_COLOR2s.eWHITE
            ElseIf (Led_up_color = enumDIS_COLORs.eWINE) Then
                Led_up_color = enumDIS_COLOR2s.eWINE
            ElseIf (Led_up_color = enumDIS_COLORs.eYellow) Then
                Led_up_color = enumDIS_COLOR2s.eYellow
            End If
            
            If (Glo_Display_Direct = "가로") Then
                DoEvents
                Call GL_Emergency_Horizontal(D1, D2, Led_up_color, Led_down_color, IN_OUT)  'New 전광판 가로 제어
            Else
                DoEvents
                Call GL_Emergency_Vertical_First(D1, D2, Led_up_color, Led_down_color, IN_OUT) 'New 전광판 세로 제어
            End If
    End If

End Sub

Public Sub Relay_Alive(RNum As Integer, gateNo As Integer)
'Dim PauseTime As Single
'Dim start  As Single
'PauseTime = 0.2
'
''RNum 0 : Gate Relay, 1: Capture Test
''On Error GoTo Err_P
'
'With FrmTcpServer
'    If (RNum = 0) Then
'        GlO_TcpDataGate = Chr$(2) & "2" & Chr$(3) & Chr$(13) & Chr$(2) & "2" & Chr$(3) '무의미한 패킷(위즈넷 깨우기 위함)
'    Else
'        GlO_TcpDataGate = Chr$(2) & "1" & Chr$(3) & Chr$(13) & Chr$(2) & "1" & Chr$(3) '무의미한 패킷(위즈넷 깨우기 위함)
'    End If
'
'    Select Case gateNo
'        Case 0
'            Select Case LANE1_DeviceMode
'                   Case "0" 'Tcp Ip
'
'                            If (.Gate1_sock(0).State <> sckClosed) Then
'                                .Gate1_sock(0).Close
'                            End If
'
'                            .Gate1_sock(0).Connect LANE1_DeviceIP, LANE1_RelayPort
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                            End If
'
'                            'Call None_Delay_Time(0.3)
''''                            GlO_GateRNum(0) = RNum
''''                            Gate_ACK(0) = False
''''                            GateTimer_First(0) = False
''''                            GlO_SendCnt(0) = 0
''''                            .GateTimer(0).Enabled = True
'
'
'
'                   Case "1" 'UDP
'                            .Gate1_sock(0).SendData (GlO_TcpDataGate)
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                            End If
'                            Call None_Delay_Time(0.3)
'            End Select
'
'        Case 1
'            Select Case LANE2_DeviceMode
'                   Case "0" 'Tcp Ip
'                            If (.Gate1_sock(1).State <> sckClosed) Then
'                                .Gate1_sock(1).Close
'                            End If
'                            .Gate1_sock(1).Connect LANE2_DeviceIP, LANE2_RelayPort
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                            End If
'                            'Call None_Delay_Time(0.1)
''''                            GlO_GateRNum(1) = RNum
''''                            Gate_ACK(1) = False
''''                            GateTimer_First(1) = False
''''                            GlO_SendCnt(1) = 0
''''                            .GateTimer(1).Enabled = True
'
'
'                   Case "1" 'UDP
'                            .Gate1_sock(1).SendData (GlO_TcpDataGate)
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                            End If
'                            Call None_Delay_Time(0.3)
'            End Select
'
'        Case 2
'            Select Case LANE3_DeviceMode
'                   Case "0" 'Tcp Ip
'                            If (.Gate1_sock(2).State <> sckClosed) Then
'                                .Gate1_sock(2).Close
'                            End If
'                            .Gate1_sock(2).Connect LANE3_DeviceIP, LANE3_RelayPort
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                            End If
'                            'Call None_Delay_Time(0.1)
''''                            GlO_GateRNum(2) = RNum
''''                            Gate_ACK(2) = False
''''                            GateTimer_First(2) = False
''''                            GlO_SendCnt(2) = 0
''''                            .GateTimer(2).Enabled = True
'
'
'                   Case "1" 'UDP
'                            .Gate1_sock(2).SendData (GlO_TcpDataGate)
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                            End If
'                            Call None_Delay_Time(0.3)
'
'            End Select
'
'        Case 3
'            Select Case LANE4_DeviceMode
'                   Case "0" 'Tcp Ip
'                            If (.Gate1_sock(3).State <> sckClosed) Then
'                                .Gate1_sock(3).Close
'                            End If
'                            .Gate1_sock(3).Connect LANE4_DeviceIP, LANE4_RelayPort
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                            End If
'                            'Call None_Delay_Time(0.1)
''''                            GlO_GateRNum(3) = RNum
''''                            Gate_ACK(3) = False
''''                            GateTimer_First(3) = False
''''                            GlO_SendCnt(3) = 0
''''                            .GateTimer(3).Enabled = True
'
'
'                   Case "1" 'UDP
'                            .Gate1_sock(3).SendData (GlO_TcpDataGate)
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                            End If
'                            Call None_Delay_Time(0.3)
'
'            End Select
'
'        Case 4
'            Select Case LANE5_DeviceMode
'                   Case "0" 'Tcp Ip
'                            If (.Gate1_sock(4).State <> sckClosed) Then
'                                .Gate1_sock(4).Close
'                            End If
'                            .Gate1_sock(4).Connect LANE5_DeviceIP, LANE5_RelayPort
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE TCP/IP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                            End If
'                            'Call None_Delay_Time(0.1)
''''                            GlO_GateRNum(4) = RNum
''''                            Gate_ACK(4) = False
''''                            GateTimer_First(4) = False
''''                            GlO_SendCnt(4) = 0
''''                            .GateTimer(4).Enabled = True
'
'                   Case "1" 'UDP
'                            .Gate1_sock(4).SendData (GlO_TcpDataGate)
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                            End If
'                            Call None_Delay_Time(0.3)
'            End Select
'
'        Case 5
'            Select Case LANE6_DeviceMode
'                   Case "0" 'Tcp Ip
'                            If (.Gate1_sock(5).State <> sckClosed) Then
'                                .Gate1_sock(5).Close
'                            End If
'                            .Gate1_sock(5).Connect LANE6_DeviceIP, LANE6_RelayPort
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE TCP/IP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                            End If
'                            'Call None_Delay_Time(0.1)
''''                            GlO_GateRNum(5) = RNum
''''                            Gate_ACK(5) = False
''''                            GateTimer_First(5) = False
''''                            GlO_SendCnt(5) = 0
''''                            .GateTimer(5).Enabled = True
'
'                   Case "1" 'UDP
'                            .Gate1_sock(5).SendData (GlO_TcpDataGate)
'                            If (RNum = 0) Then
'                                Call DataLogger("[GATE AWAKE UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                            End If
'                            Call None_Delay_Time(0.3)
'            End Select
'
'    End Select
'
'
'End With
'
'Exit Sub
'
'Err_P:

End Sub

Public Sub Relay_Out(RNum As Integer, gateNo As Integer)
Dim PauseTime As Single
Dim start  As Single
PauseTime = 0.2

'RNum 0 : Gate Relay, 1: Capture Test
'On Error GoTo Err_P

With FrmTcpServer

    
        If (Glo_LPRBoard = "위즈넷") Then
            If (LANE1_DeviceMode = "0") Then 'TCP
                If (RNum = 0) Then
                    GlO_TcpDataGate = Chr$(2) & "R2" & Chr$(3)
                Else
                    GlO_TcpDataGate = Chr$(2) & "R1" & Chr$(3)
                End If
                
            ElseIf (LANE1_DeviceMode = "1") Then 'UDP
                If (RNum = 0) Then
                    GlO_TcpDataGate = Chr$(2) & "R2" & Chr$(3) & Chr$(13) & Chr$(2) & "R2" & Chr$(3)
                Else
                    GlO_TcpDataGate = Chr$(2) & "R1" & Chr$(3) & Chr$(13) & Chr$(2) & "R1" & Chr$(3)
                End If
            End If
            
            
        ElseIf (Glo_LPRBoard = "자두이노") Then
                If (RNum = 0) Then
                    GlO_TcpDataGate = Chr$(2) & "GATE UP" & Chr$(3) '차단기오픈(차단기 컨트롤러 프로토콜:FPtech와 다인전자 동일함)(Test:Debug용)
                Else
                    GlO_TcpDataGate = "GETFRAME" '캡쳐(자두이노에서 구현해야 함)
                End If
        End If
            
            
            Select Case gateNo
                Case 0
                    Select Case LANE1_DeviceMode
                           Case "0" 'Tcp Ip
                                    
                                    Glo_Gate_ReconnCnt(0) = 0 '재접속 카운트 초기화

                                    If (.Gate1_sock(0).State <> sckConnected) Then
                                        .Gate1_sock(0).Close
                                        
                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        .Gate1_sock(0).Connect LANE1_DeviceIP, LANE1_RelayPort
                                    
                                    ElseIf (.Gate1_sock(0).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP 전송]  준비 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        Dim bData() As Byte
                                        ReDim bData(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(0).State = sckConnected) Then
                                            .Gate1_sock(0).SendData bData
                                        Else
                                            Call DataLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(0).State)
                                            Call DebugLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(0).State)
                                        End If

                                    Else
                                        Call DataLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(0).State)
                                        Call DebugLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(0).State)
                                    End If
                                    
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                                    End If
                                    
                                    Call None_Delay_Time(0.1)
                                    
                                    GlO_GateRNum(0) = RNum
                                    Gate_ACK(0) = False
                                    GateTimer_First(0) = False
                                    GlO_SendCnt(0) = 0
                                    '.GateTimer(0).Enabled = True
                                    .GateTimer(0).Enabled = False
                                    
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(0).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
'                                    If (RNum = 0) Then '차단기오픈
'                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                                        .Gate1_sock(0).RemoteHost = LANE1_DeviceIP
'                                        .Gate1_sock(0).RemotePort = LANE1_RelayPort
'                                    Else    '캡쳐
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_RelayPort)
'                                        .Gate1_sock(0).RemoteHost = LANE1_DispIP '위즈넷보드 IP:차단기,
'                                        .Gate1_sock(0).RemotePort = LANE1_RelayPort
'                                    End If
'                                    .Gate1_sock(0).SendData (GlO_TcpDataGate)
'                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 1
                    Select Case LANE2_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(1) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(1).State <> sckConnected) Then
                                        .Gate1_sock(1).Close
                                        
                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        .Gate1_sock(1).Connect LANE2_DeviceIP, LANE2_RelayPort
                                    
                                    ElseIf (.Gate1_sock(1).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP 전송]  준비 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        Dim bData1() As Byte
                                        ReDim bData1(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData1 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(1).State = sckConnected) Then
                                            .Gate1_sock(1).SendData bData1
                                        Else
                                            Call DataLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", 소켓상태 = " & .Gate1_sock(1).State)
                                            Call DebugLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", 소켓상태 = " & .Gate1_sock(1).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(1).State)
                                        Call DebugLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(1).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                                    End If
                                    
                                    'Call None_Delay_Time(0.1)
                                    
                                    GlO_GateRNum(1) = RNum
                                    Gate_ACK(1) = False
                                    GateTimer_First(1) = False
                                    GlO_SendCnt(1) = 0
                                    '.GateTimer(1).Enabled = True
                                    .GateTimer(1).Enabled = False
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(1).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 2
                    Select Case LANE3_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(2) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(2).State <> sckConnected) Then
                                        .Gate1_sock(2).Close
                                        
                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        .Gate1_sock(2).Connect LANE3_DeviceIP, LANE3_RelayPort
                                    
                                    ElseIf (.Gate1_sock(2).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP 전송]  준비 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        Dim bData2() As Byte
                                        ReDim bData2(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData2 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(2).State = sckConnected) Then
                                            .Gate1_sock(2).SendData bData2
                                        Else
                                            Call DataLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(2).State)
                                            Call DebugLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(2).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(2).State)
                                        Call DebugLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(2).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                                    End If
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(2) = RNum
                                    Gate_ACK(2) = False
                                    GateTimer_First(2) = False
                                    GlO_SendCnt(2) = 0
                                    '.GateTimer(2).Enabled = True
                                    .GateTimer(2).Enabled = False
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(2).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                
                Case 3
                    Select Case LANE4_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(3) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(3).State <> sckConnected) Then
                                        .Gate1_sock(3).Close
                                        
                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        .Gate1_sock(3).Connect LANE4_DeviceIP, LANE4_RelayPort
                                    
                                    ElseIf (.Gate1_sock(3).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP 전송]  준비 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        Dim bData3() As Byte
                                        ReDim bData3(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData3 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(3).State = sckConnected) Then
                                            .Gate1_sock(3).SendData bData3
                                        Else
                                            Call DataLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", 소켓상태 = " & .Gate1_sock(3).State)
                                            Call DebugLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", 소켓상태 = " & .Gate1_sock(3).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(3).State)
                                        Call DebugLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(3).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                                    End If
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(3) = RNum
                                    Gate_ACK(3) = False
                                    GateTimer_First(3) = False
                                    GlO_SendCnt(3) = 0
                                    '.GateTimer(3).Enabled = True
                                    .GateTimer(3).Enabled = False
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(3).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                    
                Case 4
                    Select Case LANE5_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(4) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(4).State <> sckConnected) Then
                                        .Gate1_sock(4).Close
                                    
                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        .Gate1_sock(4).Connect LANE5_DeviceIP, LANE5_RelayPort
                                    
                                    ElseIf (.Gate1_sock(4).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP 전송]  준비 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        Dim bData4() As Byte
                                        ReDim bData4(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData4 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(4).State = sckConnected) Then
                                            .Gate1_sock(4).SendData bData4
                                        Else
                                            Call DataLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", 소켓상태 = " & .Gate1_sock(4).State)
                                            Call DebugLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", 소켓상태 = " & .Gate1_sock(4).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(4).State)
                                        Call DebugLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(4).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                                    End If
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(4) = RNum
                                    Gate_ACK(4) = False
                                    GateTimer_First(4) = False
                                    GlO_SendCnt(4) = 0
                                    '.GateTimer(4).Enabled = True
                                    .GateTimer(4).Enabled = False
                           
                           Case "1" 'UDP
                                    .Gate1_sock(4).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 5
                    Select Case LANE6_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(5) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(5).State <> sckConnected) Then
                                        .Gate1_sock(5).Close
                                    
                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        .Gate1_sock(5).Connect LANE6_DeviceIP, LANE6_RelayPort
                                    
                                    ElseIf (.Gate1_sock(5).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP 전송]  준비 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        Dim bData5() As Byte
                                        ReDim bData5(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData5 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(5).State = sckConnected) Then
                                            .Gate1_sock(5).SendData bData5
                                        Else
                                            Call DataLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", 소켓상태 = " & .Gate1_sock(5).State)
                                            Call DebugLogger("[GATE TCP/IP 전송]  실패 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", 소켓상태 = " & .Gate1_sock(5).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(5).State)
                                        Call DebugLogger("[GATE TCP/IP 상태 예외]  : " & .Gate1_sock(5).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                                    End If
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(5) = RNum
                                    Gate_ACK(5) = False
                                    GateTimer_First(5) = False
                                    GlO_SendCnt(5) = 0
                                    '.GateTimer(5).Enabled = True
                                    .GateTimer(5).Enabled = False
                           
                           Case "1" 'UDP
                                    .Gate1_sock(5).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
            
            End Select
    
    
End With

Exit Sub

Err_p:


End Sub

Public Sub Relay_Close(RNum As Integer, gateNo As Integer)
Dim PauseTime As Single
Dim start  As Single
PauseTime = 0.2

'RNum 0 : Gate Relay, 1: Capture Test
'On Error GoTo Err_P

With FrmTcpServer

    
        If (Glo_LPRBoard = "위즈넷") Then
'            If (LANE1_DeviceMode = "0") Then 'TCP
'                If (RNum = 0) Then
'                    GlO_TcpDataGate = Chr$(2) & "R2" & Chr$(3)
'                Else
'                    GlO_TcpDataGate = Chr$(2) & "R1" & Chr$(3)
'                End If
'
'            ElseIf (LANE1_DeviceMode = "1") Then 'UDP
'                If (RNum = 0) Then
'                    GlO_TcpDataGate = Chr$(2) & "R2" & Chr$(3) & Chr$(13) & Chr$(2) & "R2" & Chr$(3)
'                Else
'                    GlO_TcpDataGate = Chr$(2) & "R1" & Chr$(3) & Chr$(13) & Chr$(2) & "R1" & Chr$(3)
'                End If
'            End If
            
            
        ElseIf (Glo_LPRBoard = "자두이노") Then
                If (RNum = 0) Then
                    GlO_TcpDataGate = Chr$(2) & "GATE DOWN" & Chr$(3) '차단기오픈(Test:debug용)
                Else
                    GlO_TcpDataGate = "GETFRAME" '캡쳐
                End If
        End If
            
            
            Select Case gateNo
                Case 0
                    Select Case LANE1_DeviceMode
                           Case "0" 'Tcp Ip
                                    
                                    Glo_Gate_ReconnCnt(0) = 0 '재접속 카운트 초기화

                                    If (.Gate1_sock(0).State <> sckConnected) Then
                                        .Gate1_sock(0).Close
                                        
                                        Call DataLogger("[GATE DOWN TCP/IP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        .Gate1_sock(0).Connect LANE1_DeviceIP, LANE1_RelayPort
                                    
                                    ElseIf (.Gate1_sock(0).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP 전송]  준비 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        Dim bData() As Byte
                                        ReDim bData(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(0).State = sckConnected) Then
                                            .Gate1_sock(0).SendData bData
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(0).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(0).State)
                                        End If

                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(0).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(0).State)
                                    End If
                                    
                                    
                                    Call None_Delay_Time(0.1)
                                    
                                    GlO_GateRNum(0) = RNum
                                    Gate_ACK(0) = False
                                    GateTimer_First(0) = False
                                    GlO_SendCnt(0) = 0
                                    '.GateTimer(0).Enabled = True
                                    .GateTimer(0).Enabled = False
                                    
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(0).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE DOWN UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 1
                    Select Case LANE2_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(1) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(1).State <> sckConnected) Then
                                        .Gate1_sock(1).Close
                                        
                                        Call DataLogger("[GATE UP TCP/IP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        .Gate1_sock(1).Connect LANE2_DeviceIP, LANE2_RelayPort
                                    
                                    ElseIf (.Gate1_sock(1).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP 전송]  준비 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        Dim bData1() As Byte
                                        ReDim bData1(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData1 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(1).State = sckConnected) Then
                                            .Gate1_sock(1).SendData bData1
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", 소켓상태 = " & .Gate1_sock(1).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", 소켓상태 = " & .Gate1_sock(1).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(1).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(1).State)
                                    End If
                                    
                                    
                                    'Call None_Delay_Time(0.1)
                                    
                                    GlO_GateRNum(1) = RNum
                                    Gate_ACK(1) = False
                                    GateTimer_First(1) = False
                                    GlO_SendCnt(1) = 0
                                    '.GateTimer(1).Enabled = True
                                    .GateTimer(1).Enabled = False
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(1).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE DOWN UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 2
                    Select Case LANE3_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(2) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(2).State <> sckConnected) Then
                                        .Gate1_sock(2).Close
                                        
                                        Call DataLogger("[GATE DOWN TCP/IP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        .Gate1_sock(2).Connect LANE3_DeviceIP, LANE3_RelayPort
                                    
                                    ElseIf (.Gate1_sock(2).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP 전송]  준비 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        Dim bData2() As Byte
                                        ReDim bData2(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData2 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(2).State = sckConnected) Then
                                            .Gate1_sock(2).SendData bData2
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(2).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", 소켓상태 = " & .Gate1_sock(2).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(2).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(2).State)
                                    End If
                                    
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(2) = RNum
                                    Gate_ACK(2) = False
                                    GateTimer_First(2) = False
                                    GlO_SendCnt(2) = 0
                                    '.GateTimer(2).Enabled = True
                                    .GateTimer(2).Enabled = False
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(2).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE DOWN UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                
                Case 3
                    Select Case LANE4_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(3) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(3).State <> sckConnected) Then
                                        .Gate1_sock(3).Close
                                        
                                        Call DataLogger("[GATE DOWN TCP/IP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        .Gate1_sock(3).Connect LANE4_DeviceIP, LANE4_RelayPort
                                    
                                    ElseIf (.Gate1_sock(3).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP 전송]  준비 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        Dim bData3() As Byte
                                        ReDim bData3(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData3 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(3).State = sckConnected) Then
                                            .Gate1_sock(3).SendData bData3
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", 소켓상태 = " & .Gate1_sock(3).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", 소켓상태 = " & .Gate1_sock(3).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(3).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(3).State)
                                    End If
                                    
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(3) = RNum
                                    Gate_ACK(3) = False
                                    GateTimer_First(3) = False
                                    GlO_SendCnt(3) = 0
                                    '.GateTimer(3).Enabled = True
                                    .GateTimer(3).Enabled = False
                                    
                           
                           Case "1" 'UDP
                                    .Gate1_sock(3).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE DOWN UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                    
                Case 4
                    Select Case LANE5_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(4) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(4).State <> sckConnected) Then
                                        .Gate1_sock(4).Close
                                    
                                        Call DataLogger("[GATE DOWN TCP/IP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        .Gate1_sock(4).Connect LANE5_DeviceIP, LANE5_RelayPort
                                    
                                    ElseIf (.Gate1_sock(4).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP 전송]  준비 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        Dim bData4() As Byte
                                        ReDim bData4(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData4 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(4).State = sckConnected) Then
                                            .Gate1_sock(4).SendData bData4
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", 소켓상태 = " & .Gate1_sock(4).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", 소켓상태 = " & .Gate1_sock(4).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(4).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(4).State)
                                    End If
                                    
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(4) = RNum
                                    Gate_ACK(4) = False
                                    GateTimer_First(4) = False
                                    GlO_SendCnt(4) = 0
                                    '.GateTimer(4).Enabled = True
                                    .GateTimer(4).Enabled = False
                           
                           Case "1" 'UDP
                                    .Gate1_sock(4).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE DOWN UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 5
                    Select Case LANE6_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(5) = 0 '재접속 카운트 초기화
                                    
                                    If (.Gate1_sock(5).State <> sckConnected) Then
                                        .Gate1_sock(5).Close
                                    
                                        Call DataLogger("[GATE DOWN TCP/IP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        .Gate1_sock(5).Connect LANE6_DeviceIP, LANE6_RelayPort
                                    
                                    ElseIf (.Gate1_sock(5).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP 전송]  준비 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        Dim bData5() As Byte
                                        ReDim bData5(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData5 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(5).State = sckConnected) Then
                                            .Gate1_sock(5).SendData bData5
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", 소켓상태 = " & .Gate1_sock(5).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP 전송]  실패 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", 소켓상태 = " & .Gate1_sock(5).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(5).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP 상태 예외]  : " & .Gate1_sock(5).State)
                                    End If
                                    
                                    'Call None_Delay_Time(0.1)
                                    GlO_GateRNum(5) = RNum
                                    Gate_ACK(5) = False
                                    GateTimer_First(5) = False
                                    GlO_SendCnt(5) = 0
                                    '.GateTimer(5).Enabled = True
                                    .GateTimer(5).Enabled = False
                           
                           Case "1" 'UDP
                                    .Gate1_sock(5).SendData (GlO_TcpDataGate)
                                    If (RNum = 0) Then
                                        Call DataLogger("[GATE DOWN UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
            
            End Select
    
    
End With

Exit Sub

Err_p:


End Sub

Public Sub FND_Display(ByVal FourNum As String, ByVal gateNo As Integer)
Dim tmpCarNo As String
Dim a() As Byte
ReDim a(12)
Dim i As Integer

'Call DFee("0123")

With FrmTcpServer
    tmpCarNo = CStr(FourNum)
    If (Len(tmpCarNo) > 6) Then
        tmpCarNo = "------"
    Else
        tmpCarNo = Space(6 - Len(tmpCarNo)) & tmpCarNo
        '.FNDTimer(GateNo).Enabled = False
        '.FNDTimer(GateNo).Enabled = True
    End If

    Dim GlO_TcpDataDisp As String
    GlO_TcpDataDisp = Chr$(2) & "WD" & tmpCarNo & Chr$(3)
    
    
    Select Case gateNo
        Case 0
            Select Case LANE1_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(0).State <> sckClosed) Then
                                .Disp1_sock(0).Close
                            End If
                            .Disp1_sock(0).Connect LANE1_DeviceIP, LANE1_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(0).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(0).PortOpen = True) Then
'                                .MSCommDisp(0).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE1_RelayComPort)
'                            End If
            End Select
        
        Case 1
            Select Case LANE2_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(1).State <> sckClosed) Then
                                .Disp1_sock(1).Close
                            End If
                            .Disp1_sock(1).Connect LANE2_DeviceIP, LANE2_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(1).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(1).PortOpen = True) Then
'                                .MSCommDisp(1).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE2_RelayComPort)
'                            End If
            End Select
    
        Case 2
            Select Case LANE3_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(2).State <> sckClosed) Then
                                .Disp1_sock(2).Close
                            End If
                            .Disp1_sock(2).Connect LANE3_DeviceIP, LANE3_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(2).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(2).PortOpen = True) Then
'                                .MSCommDisp(2).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE3_RelayComPort)
'                            End If
            End Select
        
        
        Case 3
            Select Case LANE4_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(3).State <> sckClosed) Then
                                .Disp1_sock(3).Close
                            End If
                            .Disp1_sock(3).Connect LANE4_DeviceIP, LANE4_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(3).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(3).PortOpen = True) Then
'                                .MSCommDisp(3).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE4_RelayComPort)
'                            End If
            End Select
            
        Case 4
            Select Case LANE5_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(4).State <> sckClosed) Then
                                .Disp1_sock(4).Close
                            End If
                            .Disp1_sock(4).Connect LANE5_DeviceIP, LANE5_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(4).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(4).PortOpen = True) Then
'                                .MSCommDisp(4).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE5_RelayComPort)
'                            End If
            End Select
            
        Case 5
            Select Case LANE6_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(5).State <> sckClosed) Then
                                .Disp1_sock(5).Close
                            End If
                            .Disp1_sock(5).Connect LANE6_DeviceIP, LANE6_DispPort
                            Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(5).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(5).PortOpen = True) Then
'                                .MSCommDisp(5).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c 전송 PORT = " & LANE6_RelayComPort)
'                            End If
            End Select
            
    End Select
    
    
End With
End Sub


Public Sub GL_Nomal_ParkFullLight(D1 As String, Color As Byte)
    Dim Header(21) As Byte
    Dim ColorArr(4) As Byte
    Dim i As Integer
    Dim AsciiLen As Integer
    Dim AsciiStr() As Byte
    Dim Data() As Byte

        On Error GoTo Err_p

        Header(0) = &H10
        Header(1) = &H2
        Header(2) = &H0
        Header(3) = &H0     '데이터길이
        Header(4) = &H11    '데이터길이
        Header(5) = &H94    '커맨드
        Header(6) = &H1     '페이지
        Header(7) = &H0
        Header(8) = &H63    '저장매체 플래시롬
        Header(9) = &H0
        Header(10) = &H0
        Header(11) = &H3
        Header(12) = &H1
        Header(13) = &H0
        Header(14) = &H0
        Header(15) = &H0    '효과속도
        Header(16) = &H8    '표시시간
        Header(17) = &H0
        Header(18) = &H0
        Header(19) = &H0
        Header(20) = &H4
        Header(21) = &H0


        ColorArr(0) = Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색//  1:적
        ColorArr(1) = Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색
        ColorArr(2) = Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색
        ColorArr(3) = Color    ' &H31 : 적색 / 32 : 녹색 / 33 : 노란색

        ReDim Data(4) As Byte

        AsciiStr = StrConv(D1, vbFromUnicode)
        AsciiLen = UBound(AsciiStr)
        
        For i = 3 To 0 Step -1 '1단2열 전광판
            If (AsciiLen >= 0) Then
                Data(i) = "&H" & Hex(AsciiStr(AsciiLen))
                AsciiLen = AsciiLen - 1
            Else
                Data(i) = "&H" & Hex(32)
            End If
        Next i


        ReDim GlO_ParkFullLight_BData(29 + 2) As Byte
        For i = 0 To UBound(Header)
           GlO_ParkFullLight_BData(i) = Header(i)
        Next i
        For i = 1 To UBound(ColorArr)
            GlO_ParkFullLight_BData(UBound(Header) + i) = ColorArr(i - 1)
        Next i
        For i = 1 To UBound(Data)
            GlO_ParkFullLight_BData(UBound(Header) + UBound(ColorArr) + i) = Data(i - 1)
        Next i


        GlO_ParkFullLight_BData(4) = GlO_ParkFullLight_BData(4) + UBound(ColorArr) + UBound(Data)
        GlO_ParkFullLight_BData(30) = "&H10"
        GlO_ParkFullLight_BData(31) = "&H3"

'        With FrmTcpServer
'
'            Glo_ParkFullLIGHT_IP = "192.168.0.130"
'            Glo_ParkFullLIGHT_PORT = 5000
'
'            If (.ParkFullLightS_sock.State = sckClosed) Then
'                Call DataLogger("[만차등 TCP/IP 접속] 시도 IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
'                .ParkFullLightS_sock.Connect Glo_ParkFullLIGHT_IP, Glo_ParkFullLIGHT_PORT
'            Else
'                Call DataLogger("[만차등 TCP/IP 전송] 준비 IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
'                .ParkFullLightS_sock.SendData GlO_ParkFullLight_BData
'            End If
'
'
'            Call None_Delay_Time(0.1)
'        End With

        With FrmTcpServer

            Select Case Glo_ParkFullLight_DispMode
                   Case "0" 'Tcp Ip
                            If (.ParkFullLightS_sock.State <> sckClosed) Then
                                .ParkFullLightS_sock.Close
                            End If
                            .ParkFullLightS_sock.Connect Glo_ParkFullLIGHT_IP, Glo_ParkFullLIGHT_PORT
                            Call DataLogger("[만차등 DISP TCP/IP 접속]  시도 IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
                   
                   Case "1" 'UDP
                            .ParkFullLightS_sock.SendData GlO_ParkFullLight_BData
                            Call DataLogger("[만차등 DISP UDP 전송]  IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
            End Select
            
            Call None_Delay_Time(0.1)
        End With


Exit Sub

Err_p:

'Debug.Print Err.Description
Call DebugLogger("[ParkFullLight Err] " & Err.Description)


End Sub

'긴급문구:가로
Public Sub GL_Emergency_Horizontal(D1 As String, D2 As String, Nomal_Up_color As Byte, Nomal_Down_color As Byte, IN_OUT As Integer)
    Dim Head_Up(21) As Byte
    Dim Head_Down(21) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim sHEX_Up() As Byte
    Dim sHEX_Down() As Byte
    Dim Finish(1) As Byte
    Dim D() As Byte
    Dim Up_Len As Integer
    Dim Down_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i As Integer
    Dim g As Integer
    Dim First_Str As String
    Dim Second_Str As String


        Up_Len = LenH(LeftH(D1, 12))
        Down_Len = LenH(LeftH(D2, 12))
        If (Up_Len > Down_Len) Then
            Bigger_Len = Up_Len
        Else
            Bigger_Len = Down_Len
        End If
        
        If Up_Len > Down_Len Then
            For g = 1 To (Up_Len - Down_Len)
                D2 = D2 + " "
            Next g
        Else
            For g = 1 To (Down_Len - Up_Len)
                D1 = D1 + " "
            Next g
        End If
        
        
        
        
        On Error GoTo Err_p
        
        Head_Up(0) = &H10    'DLE
        Head_Up(1) = &H2     'STX
        Head_Up(2) = &H0     'DST
        Head_Up(3) = &H0     'LEN
        Head_Up(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력
        
        'Up 속성 시작
        Head_Up(5) = &H94    '고정
        Head_Up(6) = &H0     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Up(7) = &H0     '※섹션번호(0)
        Head_Up(8) = &H1    '표시제어(H63:무한반복)
        Head_Up(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시, 00:현재 표시문구 종료 후 표시
        Head_Up(10) = &H0    '모듈 분할 하지 않음
        Head_Up(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        Head_Up(12) = &H1    '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
        Head_Up(13) = &H1    '퇴장효과
        Head_Up(14) = &H0    '보조효과:&H0, 사용하지 않음
        'Head_Up(15) = &H14               '효과속도:일반적으로 H14(20)으로 설정함
        Head_Up(15) = &H0                '효과속도:일반적인 속도보다 조금 느림 H1E(30)으로 설정,
        
        'Head_Up(16) = &H4                '※유지시간:4초( 8 x 0.5초), ※섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함
        'Head_Up(16) = "&H" & Hex(enumDISP_EMG_TIME.e10sec)
        Head_Up(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime * 2)
        
        Head_Up(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(18) = &H0                'Y축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(19) = &H18               'X축 종료점:0픽셀(섹션분리할 경우 사용함), H18:96픽셀
        Head_Up(20) = &H8                'Y축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Up 속성 끝
        
        
        'Down 속성 시작
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력
        
        'Down 속성 시작
        Head_Down(5) = &H94    '고정
        Head_Down(6) = &H0     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Down(7) = &H1     '※섹션번호(0)
        Head_Down(8) = &H1    '표시제어(H63:무한반복)
        Head_Down(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시, 00:현재 표시문구 종료 후 표시
        Head_Down(10) = &H0    '모듈 분할 하지 않음
        Head_Down(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        Head_Down(12) = &H1    '입장효과 => 이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)
        Head_Down(13) = &H1    '퇴장효과
        Head_Down(14) = &H0    '보조효과:&H0, 사용하지 않음
        'Head_Down(15) = &H14               '효과속도:일반적으로 H14(20)으로 설정함
        Head_Down(15) = &H0                '효과속도:긴급문구에서는 무의미함(0 또는 FF으로 설정함)
        
        'Head_Down(16) = &H8                '※유지시간:4초( 8 x 0.5초), ※섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함
        'Head_Down(16) = "&H" & Hex(enumDISP_EMG_TIME.e10sec)
        Head_Down(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime * 2)
        
        Head_Down(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(18) = &H4                'Y축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(19) = &H18               'X축 종료점:0픽셀(섹션분리할 경우 사용함), H18:96픽셀
        Head_Down(20) = &H8                'Y축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Down 속성 끝
        
        
        ReDim Color_Up(Bigger_Len - 1) As Byte
        ReDim Color_Down(Bigger_Len - 1) As Byte
        
        For i = 0 To UBound(Color_Up)
            'Color_Up(i) = Nomal_Up_color
            Color_Up(i) = Nomal_Up_color
        Next i
        
        For i = 0 To UBound(Color_Down)
            'Color_Down(i) = Nomal_Down_color
            Color_Down(i) = Nomal_Down_color
        Next i
        
        ReDim sHEX_Up(Bigger_Len - 1) As Byte
        ReDim sHEX_Down(Bigger_Len - 1) As Byte
        
        First_Str = D1
        D = StrConv(First_Str, vbFromUnicode)
        '윗줄(가로)
        For i = 0 To UBound(D)
            sHEX_Up(i) = "&H" & Hex(D(i))
        Next i
        
        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        '아랫줄(가로)
        For i = 0 To UBound(D)
            sHEX_Down(i) = "&H" & Hex(D(i))
        Next i
        
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1)
        Head_Up(4) = "&H" & Hex(data_len)   '데이터 길이
        Head_Down(4) = "&H" & Hex(data_len)
        
'        Dim strHex As String
'        strHex = ByteArrayToHex(Head_Up)
'        Debug.Print "Head_Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Up)
'        Debug.Print "Color Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Down)
'        Debug.Print "Color Dn:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Up)
'        Debug.Print "Data Up:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Down)
'        Debug.Print "Data Dn:" & strHex

    
        
        Select Case IN_OUT
        
            Case 0
                ReDim GloDisp_BData1(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData1_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData1(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i
                'Debug.Print "UP:" & ByteArrayToHex(GloDisp_BData1)
                ''''''''''
                For i = 0 To UBound(Head_Down)
                   GloDisp_BData1_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
                'Debug.Print "DN:" & ByteArrayToHex(GloDisp_BData1_Down)
            Case 1
                ReDim GloDisp_BData2(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData2_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData2(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData2_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i

            Case 2
                ReDim GloDisp_BData3(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData3_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData3(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData3_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 3
                ReDim GloDisp_BData4(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData4_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData4(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData4_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 4
                ReDim GloDisp_BData5(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData5_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData5(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData5_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 5
                ReDim GloDisp_BData6(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData6_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData6(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData6_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
        
        End Select
        
        
        
        With FrmTcpServer
            Select Case IN_OUT
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)

Exit Sub

Err_p:



End Sub


Public Sub GL_Emergency_Vertical(D1 As String, D2 As String, Nomal_Up_color As Byte, Nomal_Down_color As Byte, IN_OUT As Integer)
    Dim Head_Up(21) As Byte
    Dim Head_Down(21) As Byte
    Dim Color_Up() As Byte
    Dim Color_Down() As Byte
    Dim sHEX_Up() As Byte
    Dim sHEX_Down() As Byte
    Dim Finish(1) As Byte
    Dim D() As Byte
    Dim Up_Len As Integer
    Dim Down_Len As Integer
    Dim Bigger_Len As Integer
    Dim Gap_Len As Integer
    Dim i, j As Integer
    Dim g As Integer
    Dim First_Str As String
    Dim Second_Str As String
    Dim iAscCount As Integer
    Dim Up_Unicode(256) As Byte
    Dim DOWN_Unicode(256) As Byte
    Dim iUniIDX As Integer
    Dim iUp_Unicode_Len As Integer
    Dim iDown_Unicode_Len As Integer

        '윗줄, 아랫줄 문자열 길이중에서 가장 긴 길이 찾기
'        Up_Len = LenH(D1)
'        Down_Len = LenH(D2)
'        If (Up_Len > Down_Len) Then
'            Bigger_Len = Up_Len
'        Else
'            Bigger_Len = Down_Len
'        End If
        
        Up_Len = Len(D1)
        Down_Len = Len(D2)
        If (Up_Len > Down_Len) Then
            Bigger_Len = Up_Len
        Else
            Bigger_Len = Down_Len
        End If
        
'        '윗줄, 아랫줄 문자열 길이 같게 만듬
'''        If Up_Len > Down_Len Then
'''            For g = 1 To (Up_Len - Down_Len)
'''                D2 = D2 + " "
'''            Next g
'''        Else
'''            For g = 1 To (Down_Len - Up_Len)
'''                D1 = D1 + " "
'''            Next g
'''        End If
        
        '아스키문자 수 계산(유니코드로 변경할때 아스키문자 수 만큼 더 만들어야 함)
'        iAscCount = 0
'        D = StrConv(D1, vbFromUnicode)
'        For i = 0 To UBound(D)
'            If (D(i) >= 32 And D(i) <= 126) Then '1byte 아스키문자
'                iAscCount = iAscCount + 1
'            End If
'        Next
        
        On Error GoTo Err_p
        
        Head_Up(0) = &H10    'DLE
        Head_Up(1) = &H2     'STX
        Head_Up(2) = &H0     'DST
        Head_Up(3) = &H0     'LEN
        Head_Up(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력

        'Up 속성 시작
        Head_Up(5) = &H94    '고정
        Head_Up(6) = &H0     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Up(7) = &H0     '※섹션번호
        Head_Up(8) = &H2    '표시제어(무한반복)
        Head_Up(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시, 00:현재 표시문구 종료 후 표시
        Head_Up(10) = &H0    '모듈 분할 하지 않음
        Head_Up(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        Head_Up(12) = &H1    '입장효과 => {정지:&H1}, {이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)}
        Head_Up(13) = &H1    '퇴장효과 => {정지:&H1}, {이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)}
        Head_Up(14) = &H0    '보조효과:&H0, 사용하지 않음
        Head_Up(15) = &H14   '효과속도:일반적으로 H14(20)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        'Head_Up(15) = &H0   '효과속도:일반적으로 H1E(30)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)

        'Head_Up(16) = &H4                '※유지시간: 섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함(페이지메세지에서는 의미없는 듯함)
        'Head_Up(16) = "&H" & Hex(enumDISP_EMG_TIME.e3sec)        '유지시간
        Head_Up(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime) '유지시간
        
        Head_Up(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(18) = &H0                'Y축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(19) = &H18               'X축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(20) = &H4                'Y축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Up(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Up 속성 끝


        'Down 속성 시작
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '데이터전체길이(속성길이+문자길이+컬러길이) - 아래쪽에서 재계산 후 입력

        Head_Down(5) = &H94    '고정
        Head_Down(6) = &H0     '문구형식: 긴급(실시간메세지) 00 / 일반(페이지메세지) 01
        Head_Down(7) = &H1     '섹션번호(1)
        Head_Down(8) = &H2    '표시제어(무한반복)
        Head_Down(9) = &H1     '01:현재 표시문구 삭제 후 즉시표시
        Head_Down(10) = &H0    '모듈 분할 하지 않음
        Head_Down(11) = &H3    '폰트크기:16x16픽셀(단, 영문/숫자는 8X16)
        Head_Down(12) = &H1    '입장효과 => {정지:&H1}, {이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)}
        Head_Down(13) = &H1    '퇴장효과 => {정지:&H1}, {이동하기: 왼쪽(&H6), 오른쪽(&H7), 위쪽(&H8), 아래쪽(&H9)}
        Head_Down(14) = &H0    '보조효과:&H0, 사용하지 않음
        Head_Down(15) = &H14   '효과속도:일반적으로 H14(20)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)
        'Head_Down(15) = &H0  '효과속도:일반적으로 H1E(30)으로 설정함(10,20,30,..90:숫자가 작을수록 빨라짐)

        'Head_Down(16) = &H4                '유지시간:4초( 8 x 0.5초), ※섹션분리할 경우 상단섹션은 0, 하단섹션에서 설정함, 긴 문장의 경우 0으로 설정함(페이지메세지에서는 의미없는 듯함)
        'Head_Down(16) = "&H" & Hex(enumDISP_EMG_TIME.e3sec)        '유지시간
        Head_Down(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime) '유지시간

        Head_Down(17) = &H0                'X축 시작점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(18) = &H4                'Y축 시작점:0픽셀(섹션분리할 경우 사용함) : 16픽셀
        Head_Down(19) = &H18                'X축 종료점:0픽셀(섹션분리할 경우 사용함)
        Head_Down(20) = &H8                'Y축 종료점:0픽셀(섹션분리할 경우 사용함) : 32픽셀
        Head_Down(21) = &H0                '배경이미지 삽입:0(사용안함)
        'Down 속성 끝
        
                
        'ReDim Color_Up(Bigger_Len - 1 + iAscCount) As Byte
        'ReDim Color_Down(Bigger_Len - 1 + iAscCount) As Byte
        ReDim Color_Up(Bigger_Len * 2 - 1) As Byte
        ReDim Color_Down(Bigger_Len * 2 - 1) As Byte
        
        For i = 0 To UBound(Color_Up)
            'Debug.Print i
            Color_Up(i) = Nomal_Up_color + 8 '세로출력위해 +8
        Next i
        
        For i = 0 To UBound(Color_Down)
            Color_Down(i) = Nomal_Down_color + 8 '세로출력위해 +8
        Next i
        
        'ReDim sHEX_Up(Bigger_Len - 1 + iAscCount) As Byte
        'ReDim sHEX_Down(Bigger_Len - 1 + iAscCount) As Byte
        ReDim sHEX_Up(Bigger_Len * 2 - 1) As Byte
        ReDim sHEX_Down(Bigger_Len * 2 - 1) As Byte
        
        First_Str = D1
        D = StrConv(First_Str, vbFromUnicode)
        j = 0
        For i = 0 To UBound(D)
            If (D(i) >= 32 And D(i) <= 126) Then
                sHEX_Up(j) = "&HE0"
                j = j + 1
                sHEX_Up(j) = "&H" & Hex(D(i))
            Else
                sHEX_Up(j) = "&H" & Hex(D(i))
            End If
            j = j + 1
        Next i
        
        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        j = 0
        For i = 0 To UBound(D)
            If (D(i) >= 32 And D(i) <= 126) Then
                sHEX_Down(j) = "&HE0"
                j = j + 1
                sHEX_Down(j) = "&H" & Hex(D(i))
            Else
                sHEX_Down(j) = "&H" & Hex(D(i))
            End If
            j = j + 1
        Next i
        
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1)
        Head_Up(4) = "&H" & Hex(data_len)   '데이터 길이
        Head_Down(4) = "&H" & Hex(data_len)
        
'        Dim strHex As String
'        strHex = ByteArrayToHex(Head_Up)
'        Debug.Print "Head_Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Up)
'        Debug.Print "Color Up:" & strHex
'
'        strHex = ByteArrayToHex(Color_Down)
'        Debug.Print "Color Dn:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Up)
'        Debug.Print "Data Up:" & strHex
'
'        strHex = ByteArrayToHex(sHEX_Down)
'        Debug.Print "Data Dn:" & strHex

    
        
        Select Case IN_OUT
        
            Case 0
                ReDim GloDisp_BData1(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData1_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData1(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i
                'Debug.Print "긴급UP:" & ByteArrayToHex(GloDisp_BData1)
                ''''''''''
                For i = 0 To UBound(Head_Down)
                   GloDisp_BData1_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData1_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
                'Debug.Print "긴급DN:" & ByteArrayToHex(GloDisp_BData1_Down)
            Case 1
                ReDim GloDisp_BData2(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData2_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData2(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData2_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData2_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i

            Case 2
                ReDim GloDisp_BData3(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData3_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData3(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData3_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData3_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 3
                ReDim GloDisp_BData4(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData4_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData4(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData4_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData4_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 4
                ReDim GloDisp_BData5(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData5_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData5(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData5_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData5_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
            Case 5
                ReDim GloDisp_BData6(data_len - 1 + 7) As Byte
                ReDim GloDisp_BData6_Down(data_len - 1 + 7) As Byte

                For i = 0 To UBound(Head_Up)
                   GloDisp_BData6(i) = Head_Up(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + 1) = Color_Up(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + 2) = sHEX_Up(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6(i + UBound(Head_Up) + UBound(Color_Up) + UBound(sHEX_Up) + 3) = Finish(i)
                Next i

                For i = 0 To UBound(Head_Down)
                   GloDisp_BData6_Down(i) = Head_Down(i)
                Next i
                For i = 0 To UBound(Color_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + 1) = Color_Down(i)
                Next i
                For i = 0 To UBound(sHEX_Up)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + 2) = sHEX_Down(i)
                Next i
                For i = 0 To UBound(Finish)
                    GloDisp_BData6_Down(i + UBound(Head_Down) + UBound(Color_Down) + UBound(sHEX_Down) + 3) = Finish(i)
                Next i
        
        End Select
        
        
        
        With FrmTcpServer
            Select Case IN_OUT
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub


'차량번호:뒷4자리 숫자 분리 출력
Public Sub GL_Emergency_Vertical_First(D1 As String, D2 As String, Nomal_Up_color As Byte, Nomal_Down_color As Byte, gateNo As Integer)
    
    'D1:차량번호인지 아닌지 확인
    Dim i As Integer
    Dim sCarNo As String
    Dim sCarNo1 As String
    Dim sCarNo2 As String
    Dim sCarStat As String
    Dim sCarStat1 As String
    Dim sCarStat2 As String
    Dim iCarNoLen As Integer
    Dim iToggleCount As Integer
    iToggleCount = 2 '차량번호 + 처리결과 2회 출력
    
    iCarNoLen = LenH(D1)
    If ((IsNumeric(Right(D1, 4)) = True) And (iCarNoLen = 8 Or iCarNoLen = 9 Or iCarNoLen = 11 Or iCarNoLen = 12)) Then
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '방법1:타이머만 이용해서 출력할 경우 첫출력까지 약 1000ms 기다려야함
'''                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = False
'''                '차량번호
'''                sCarNo = D1
'''                sCarNo1 = LeftH(D1, iCarNoLen - 4)
'''                sCarNo2 = Right(D1, 4)
'''                sCarStat = D2
'''                sCarStat1 = Left(D2, Int(Len(D2) / 2) + Len(D2) Mod 2) '문자열/2 앞부분 + Mod 2 나머지 문자열
'''                sCarStat2 = Mid(D2, (Int(Len(D2) / 2) + Len(D2) Mod 2) + 1)
'''
'''
'''                Glo_Emerg_Vertical(gateNo).CarNo1 = sCarNo1
'''                Glo_Emerg_Vertical(gateNo).CarNo2 = sCarNo2
'''                'Glo_Emerg_Vertical(gateNo).CarNoColor1 = enumDIS_COLORs.eGreen
'''                'Glo_Emerg_Vertical(gateNo).CarNoColor2 = enumDIS_COLORs.eYellow
'''                Glo_Emerg_Vertical(gateNo).CarNoColor1 = Nomal_Up_color
'''                Glo_Emerg_Vertical(gateNo).CarNoColor2 = Nomal_Up_color
'''                Glo_Emerg_Vertical(gateNo).CarNoCount = Glo_Emerg_Vertical_ToggleCount
'''
'''                Glo_Emerg_Vertical(gateNo).CarStat1 = sCarStat1
'''                Glo_Emerg_Vertical(gateNo).CarStat2 = sCarStat2
'''                Glo_Emerg_Vertical(gateNo).CarStatColor1 = Nomal_Down_color
'''                Glo_Emerg_Vertical(gateNo).CarStatColor2 = Nomal_Down_color
'''                Glo_Emerg_Vertical(gateNo).CarStatCount = Glo_Emerg_Vertical_ToggleCount
'''
'''                Glo_Emerg_Vertical(gateNo).ToggleSelect = EnumEmergToggleOrder.enumCarNo  '처음 출력한 내용: 차량번호
'''
'''
'''                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Interval = 100 '처음 출력할 내용은 즉시 출력해야함
'''                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = True
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '방법2:첫출력은 직접 처리하고, 이후부터는 타이머 이용해서 출력
                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = False
                '차량번호
                sCarNo = D1
                sCarNo1 = LeftH(D1, iCarNoLen - 4)
                sCarNo2 = Right(D1, 4)
                sCarStat = D2
                sCarStat1 = Left(D2, Int(Len(D2) / 2) + Len(D2) Mod 2) '문자열/2 앞부분 + Mod 2 나머지 문자열
                sCarStat2 = Mid(D2, (Int(Len(D2) / 2) + Len(D2) Mod 2) + 1)


                Glo_Emerg_Vertical(gateNo).CarNo1 = sCarNo1
                Glo_Emerg_Vertical(gateNo).CarNo2 = sCarNo2
                Glo_Emerg_Vertical(gateNo).CarNoColor1 = Nomal_Up_color
                Glo_Emerg_Vertical(gateNo).CarNoColor2 = Nomal_Up_color
                Glo_Emerg_Vertical(gateNo).CarNoCount = Glo_Emerg_Vertical_ToggleCount

                Glo_Emerg_Vertical(gateNo).CarStat1 = sCarStat2
                Glo_Emerg_Vertical(gateNo).CarStat2 = sCarStat1
                Glo_Emerg_Vertical(gateNo).CarStatColor1 = Nomal_Down_color
                Glo_Emerg_Vertical(gateNo).CarStatColor2 = Nomal_Down_color
                Glo_Emerg_Vertical(gateNo).CarStatCount = Glo_Emerg_Vertical_ToggleCount

                
                'DoEvents
                Call GL_Emergency_Vertical(sCarNo1, sCarNo2, Glo_Emerg_Vertical(gateNo).CarNoColor1, Glo_Emerg_Vertical(gateNo).CarNoColor2, gateNo) '차량번호 즉시 출력
                Glo_Emerg_Vertical(gateNo).CarNoCount = Glo_Emerg_Vertical_ToggleCount - 1 '차번 즉시 출력했으므로 카운트 -1 처리함
                Glo_Emerg_Vertical(gateNo).ToggleSelect = EnumEmergToggleOrder.enumCarStat  '다음에 출력할 내용(처리결과) 지정

                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Interval = Glo_Emerg_Vertical_ToggleTime * 1000 '단위 ms
                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = True
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Else
    
        '차량번호외 긴급문구
        DoEvents
        Call GL_Emergency_Vertical(D1, D2, Nomal_Up_color, Nomal_Down_color, gateNo)
    End If
        
        
Exit Sub

Err_p:

End Sub




Public Sub GL_Display_PowerOFF(gateNo As Integer)
    Dim Head(8) As Byte
    Dim i As Integer
    
        On Error GoTo Err_p

        Head(0) = &H10    'DLE
        Head(1) = &H2     'STX
        Head(2) = &H0     'DST
        Head(3) = &H0     'LEN
        Head(4) = &H2     '데이터전체길이

        'Up 속성 시작
        Head(5) = &H41
        Head(6) = &H0
        'Up 속성 끝
        Head(7) = &H10
        Head(8) = &H3
        
        Select Case gateNo
        
            Case 0
                ReDim GloDisp_BData1(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData1(i) = Head(i)
                Next i
                Debug.Print "POWER OFF:" & ByteArrayToHex(GloDisp_BData1) '임시테스트
                ''''''''''
            Case 1
                ReDim GloDisp_BData2(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData2(i) = Head(i)
                Next i

            Case 2
                ReDim GloDisp_BData3(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData3(i) = Head(i)
                Next i
                
            Case 3
                ReDim GloDisp_BData4(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData4(i) = Head(i)
                Next i
                
            Case 4
                ReDim GloDisp_BData5(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData5(i) = Head(i)
                Next i
                
            Case 5
                ReDim GloDisp_BData6(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData6(i) = Head(i)
                Next i
        
        End Select
        
        
        
        With FrmTcpServer
            Select Case gateNo
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub



Public Sub GL_Display_PowerON(gateNo As Integer)
    Dim Head(8) As Byte
    Dim i As Integer
    
        On Error GoTo Err_p

        Head(0) = &H10    'DLE
        Head(1) = &H2     'STX
        Head(2) = &H0     'DST
        Head(3) = &H0     'LEN
        Head(4) = &H2     '데이터전체길이

        'Up 속성 시작
        Head(5) = &H41
        Head(6) = &H1
        'Up 속성 끝
        Head(7) = &H10
        Head(8) = &H3
        
        Select Case gateNo
        
            Case 0
                ReDim GloDisp_BData1(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData1(i) = Head(i)
                Next i
                Debug.Print "POWER ON:" & ByteArrayToHex(GloDisp_BData1) '임시테스트
                ''''''''''
            Case 1
                ReDim GloDisp_BData2(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData2(i) = Head(i)
                Next i

            Case 2
                ReDim GloDisp_BData3(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData3(i) = Head(i)
                Next i
                
            Case 3
                ReDim GloDisp_BData4(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData4(i) = Head(i)
                Next i
                
            Case 4
                ReDim GloDisp_BData5(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData5(i) = Head(i)
                Next i
                
            Case 5
                ReDim GloDisp_BData6(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData6(i) = Head(i)
                Next i
        
        End Select
        
        
        
        With FrmTcpServer
            Select Case gateNo
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP 접속]  시도 IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP 전송]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub

