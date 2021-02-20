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


    'If (Glo_LPRBoard = "�����" Or Glo_LPRBoard = "�ڵ��̳�") Then
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
        Header(5) = &H53    'CMD : ��� 54 / �Ϲ� 53
        Header(6) = &H0     'Dummy
        Header(7) = &H0     'Dummy
        Header(8) = &H0     '�����ü �÷��÷�
        Header(9) = &H91    ' (1001 0001) B[1:0] - ����ȭ�� ��Ʈũ�� 16 font / B[5:4] - ȭ��ǥ�� ON / B[6:7] - ���� ǥ�� ���� 2 = ���ι���
        Header(10) = &H0    '��� ���� ���� ����
        Header(11) = &H0    'Dummy
        Header(12) = &H0    '����ȭ�� ȿ���� : ȿ������
        Header(13) = Nomal_Show         '&H1    ' ����ȭ�� ȿ���� : �����̵�
        Header(14) = Nomal_Speed        '&H1E   'ȿ�� �ӵ�
        Header(15) = Nomal_StopTime     '&H0    '���� �ð� ����
        Header(16) = &H0    '���� ǥ�� ��ġ : 0 ��
        Dim Up_Color As Byte
        Dim Down_Color As Byte
        
        Select Case Nomal_Up_color
            Case 0
                Up_Color = &H31 '��
            Case 1
                Up_Color = &H32 '��
            Case 2
                Up_Color = &H33 'Ȳ
        End Select
                
        Select Case Nomal_Down_color
            Case 0
                Down_Color = &H31 '��
            Case 1
                Down_Color = &H32 '��
            Case 2
                Down_Color = &H33 'Ȳ
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
        
        
        
'    If (Glo_LPRBoard = "�����") Then
    
        With FrmTcpServer
            Select Case IN_OUT
                Case 0
                    Select Case LANE1_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(0).State <> sckClosed) Then
                                        .Disp1_sock(0).Close
                                    End If
                                    .Disp1_sock(0).Connect LANE1_DispIP, LANE1_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With
'
'
'    ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
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
'            Call DataLogger("[DISP ����]  �õ� IP = " & ip & "    PORT = " & Port)
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
            Up_Len = LenH(D1)     '����: 12���ڸ� ���
            Down_Len = LenH(D2)   '����: 12���ڸ� ���
            
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
        Head_Up(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Up �Ӽ� ����
        Head_Up(5) = &H94    '����
        Head_Up(6) = &H1     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Up(7) = &H0     '�ؼ��ǹ�ȣ
        Head_Up(8) = &H63    'ǥ������(���ѹݺ�)
        Head_Up(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��, 00:���� ǥ�ù��� ���� �� ǥ��
        Head_Up(10) = &H0    '��� ���� ���� ����
        Head_Up(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
'        Head_Up(12) = &H6    '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
'        Head_Up(13) = &H6    '����ȿ��
        Head_Up(12) = Normal_Shift     '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
        Head_Up(13) = Normal_Shift     '����ȿ��
        
        Head_Up(14) = &H0    '����ȿ��:&H0, ������� ����
        'Head_Up(15) = &H14               'ȿ���ӵ�:�Ϲ������� H14(20)���� ������
        'Head_Up(15) = &H1E                'ȿ���ӵ�:�Ϲ����� �ӵ����� ���� ���� H1E(30)���� ����, '�ӽ��ּ�
        Head_Up(15) = Nomal_Speed                'ȿ���ӵ�:�Ϲ����� �ӵ����� ���� ���� H1E(30)���� ����
        
        Head_Up(16) = &H0                '�������ð�:4��( 8 x 0.5��), �ؼ��Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������
        Head_Up(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(18) = &H0                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(19) = &H18               'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(20) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(21) = &H0                '����̹��� ����:0(������)
        'Up �Ӽ� ��
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Down �Ӽ� ����
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�
        
        Head_Down(5) = &H94    '����
        Head_Down(6) = &H1     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Down(7) = &H1     '���ǹ�ȣ(1)
        Head_Down(8) = &H63    'ǥ������(���ѹݺ�)
        Head_Down(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��
        Head_Down(10) = &H0    '��� ���� ���� ����
        Head_Down(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        'Head_Down(12) = &H6    '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
        'Head_Down(13) = &H6    '����ȿ��
        Head_Down(12) = Normal_Shift     '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
        Head_Down(13) = Normal_Shift     '����ȿ��
        Head_Down(14) = &H0    '����ȿ��:&H0, ������� ����
        'Head_Down(15) = &H14              'ȿ���ӵ�:�Ϲ������� H14(20)���� ������
        'Head_Down(15) = &H1E               'ȿ���ӵ�:�Ϲ����� �ӵ����� ���� ���� H1E(30)���� ����,
        Head_Down(15) = Nomal_Speed                'ȿ���ӵ�:�Ϲ����� �ӵ����� ���� ���� H1E(30)���� ����
        
        'Head_Down(16) = &H4                '�����ð�:4��( 8 x 0.5��)
        Head_Down(16) = &H0                '�����ð�:4��( 8 x 0.5��), �� ������ ��� 0���� ������
        Head_Down(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(18) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����) : 16�ȼ�
        Head_Down(19) = &H18                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(20) = &H8                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����) : 32�ȼ�
        Head_Down(21) = &H0                '����̹��� ����:0(������)
        'Down �Ӽ� ��
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
        '����(����)
        For i = 0 To UBound(D)
            sHEX_Up(i) = "&H" & Hex(D(i))
        Next i
        
        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        '�Ʒ���(����)
        For i = 0 To UBound(D)
            sHEX_Down(i) = "&H" & Hex(D(i))
        Next i
        
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1)
        Head_Up(4) = "&H" & Hex(data_len)   '������ ����
        Head_Down(4) = "&H" & Hex(data_len)
        
        '�ӽ��׽�Ʈ
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
                'Debug.Print "�Ϲ�UP:" & ByteArrayToHex(GloDisp_BData1)
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
                'Debug.Print "�Ϲ�DN:" & ByteArrayToHex(GloDisp_BData1_Down)
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
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)

Exit Sub

Err_p:



End Sub


'������ �������
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

'        '����, �Ʒ��� ���ڿ� �����߿��� ���� �� ���� ã��
'        Up_Len = Len(D1)
'        Down_Len = Len(D2)
'        If (Up_Len > Down_Len) Then
'            Bigger_Len = Up_Len * 2     '������½� �� ���ڴ� 2BYTE ó���ϹǷ� x2 �� ���
'        Else
'            Bigger_Len = Down_Len * 2   '������½� �� ���ڴ� 2BYTE ó���ϹǷ� x2 �� ���
'        End If
        
        
        If (Normal_Shift = enumDISP_NML_SHIFT.eSTOP) Then
            D1 = Left(D1, Glo_DISP_COL)
            D2 = Left(D2, Glo_DISP_COL)
            Up_Len = Len(D1)     '����: 6���ڸ� ���
            Down_Len = Len(D2)   '����: 6���ڸ� ���
            
            If (Up_Len > Down_Len) Then
                Bigger_Len = Up_Len * 2    '������½� �� ���ڴ� 2BYTE ó���ϹǷ� x2 �� ���
            Else
                Bigger_Len = Down_Len * 2  '������½� �� ���ڴ� 2BYTE ó���ϹǷ� x2 �� ���
            End If
        
        
        ElseIf (Normal_Shift = enumDISP_NML_SHIFT.eSHIFT) Then
            Up_Len = Len(D1)
            Down_Len = Len(D2)
            
            If (Up_Len > Down_Len) Then
                Bigger_Len = Up_Len * 2     '������½� �� ���ڴ� 2BYTE ó���ϹǷ� x2 �� ���
            Else
                Bigger_Len = Down_Len * 2   '������½� �� ���ڴ� 2BYTE ó���ϹǷ� x2 �� ���
            End If
            
            If ((Bigger_Len Mod Glo_DISP_COL) = 0) Then
                Bigger_Len = Bigger_Len + 1
            Else
                Bigger_Len = Bigger_Len + Glo_DISP_COL - (Bigger_Len Mod Glo_DISP_COL) + 1
            End If
        End If



        '����, �Ʒ��� ���ڿ� ���� ���� ����
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
'''            If (D(i) >= 32 And D(i) <= 126) Then '1byte �ƽ�Ű����
'''                iAscCount = iAscCount + 1
'''            End If
'''        Next
        
        
        '�ƽ�Ű���� �� ���(�����ڵ�� �����Ҷ� �ƽ�Ű���� �� ��ŭ �� ������ ��)
'        iAscCount = 0
'        If Up_Len > Down_Len Then
'            D = StrConv(D1, vbFromUnicode)
'        Else
'            D = StrConv(D2, vbFromUnicode)
'        End If
'        For i = 0 To UBound(D)
'            If (D(i) >= 32 And D(i) <= 126) Then '1byte �ƽ�Ű����
'                iAscCount = iAscCount + 1
'            End If
'        Next

        
        On Error GoTo Err_p
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Up �Ӽ� ����
        Head_Up(0) = &H10    'DLE
        Head_Up(1) = &H2     'STX
        Head_Up(2) = &H0     'DST
        Head_Up(3) = &H0     'LEN
        Head_Up(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�
        
        'Up �Ӽ� ����
        Head_Up(5) = &H94    '����
        Head_Up(6) = &H1     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Up(7) = &H0     '�ؼ��ǹ�ȣ
        Head_Up(8) = &H63    'ǥ������(���ѹݺ�)
        Head_Up(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��, 00:���� ǥ�ù��� ���� �� ǥ��
        Head_Up(10) = &H0    '��� ���� ���� ����
        Head_Up(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        'Head_Up(12) = &H6    '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
        'Head_Up(13) = &H6    '����ȿ��
        Head_Up(12) = Normal_Shift
        Head_Up(13) = Normal_Shift

        
        Head_Up(14) = &H0    '����ȿ��:&H0, ������� ����
        'Head_Up(15) = &H1E   'ȿ���ӵ�:�Ϲ������� H1E(30)���� ������(10,20,30,..90:���ڰ� �������� ������)
        'Head_Up(15) = &H14   'ȿ���ӵ�:�Ϲ������� H14(20)���� ������(10,20,30,..90:���ڰ� �������� ������)
        Head_Up(15) = Nomal_Speed  'ȿ���ӵ�:�Ϲ������� H14(20)���� ������(10,20,30,..90:���ڰ� �������� ������)
        
        Head_Up(16) = &H0                '�������ð�:4��( 8 x 0.5��), �ؼ��Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������(�������޼��������� �ǹ̾��� ����)
        Head_Up(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(18) = &H0                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(19) = &H18               'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(20) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(21) = &H0                '����̹��� ����:0(������)
        'Up �Ӽ� ��
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Down �Ӽ� ����
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�
        
        Head_Down(5) = &H94    '����
        Head_Down(6) = &H1     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Down(7) = &H1     '���ǹ�ȣ(1)
        Head_Down(8) = &H63    'ǥ������(���ѹݺ�)
        Head_Down(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��
        Head_Down(10) = &H0    '��� ���� ���� ����
        Head_Down(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        
'        Head_Down(12) = &H6    '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
'        Head_Down(13) = &H6    '����ȿ��
        Head_Down(12) = Normal_Shift
        Head_Down(13) = Normal_Shift
        
        Head_Down(14) = &H0    '����ȿ��:&H0, ������� ����
        'Head_Down(15) = &H1E   'ȿ���ӵ�:�Ϲ������� H1E(30)���� ������(10,20,30,..90:���ڰ� �������� ������)
        'Head_Down(15) = &H14   'ȿ���ӵ�:�Ϲ������� H14(20)���� ������(10,20,30,..90:���ڰ� �������� ������)
        Head_Down(15) = Nomal_Speed   'ȿ���ӵ�:�Ϲ������� H14(20)���� ������(10,20,30,..90:���ڰ� �������� ������)
        
        Head_Down(16) = &H0                '�����ð�:4��( 8 x 0.5��), �ؼ��Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������(�������޼��������� �ǹ̾��� ����)
        Head_Down(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(18) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����) : 16�ȼ�
        Head_Down(19) = &H18                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(20) = &H8                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����) : 32�ȼ�
        Head_Down(21) = &H0                '����̹��� ����:0(������)
        'Down �Ӽ� ��
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
'''        ReDim Color_Up(Bigger_Len - 1 + iAscCount) As Byte
'''        ReDim Color_Down(Bigger_Len - 1 + iAscCount) As Byte
        ReDim Color_Up(Bigger_Len - 1) As Byte
        ReDim Color_Down(Bigger_Len - 1) As Byte
        
        For i = 0 To UBound(Color_Up)
            'Debug.Print i
            Color_Up(i) = Nomal_Up_color + 8 '����������� +8
            
        Next i
        
        For i = 0 To UBound(Color_Down)
            Color_Down(i) = Nomal_Down_color + 8 '����������� +8
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
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1) '-5 : �������
        Head_Up(4) = "&H" & Hex(data_len)   '������ ����
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
                Debug.Print "UP:" & ByteArrayToHex(GloDisp_BData1) '�ӽ��׽�Ʈ
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
                Debug.Print "DN:" & ByteArrayToHex(GloDisp_BData1_Down) '�ӽ��׽�Ʈ
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
                                    Call DataLogger("[DISP TCP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP TCP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP TCP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP TCP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP TCP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP TCP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub


'��޹���
'Led_StopTime:��½ð�(10:5��, 20:10��..)
Public Sub GL_Emergency(D1 As String, D2 As String, Led_Show As Byte, Led_Speed As Byte, Led_StopTime As Byte, Led_Repeat As Byte, Led_up_color As Byte, Led_down_color As Byte, IN_OUT As Integer)
    
    If (Glo_Display = "������" Or Glo_Display = "������(Ǯ�÷�)") Then
            
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
                    Header(5) = &H54                'CMD : ��� 54 / �Ϲ� 53
                    Header(6) = &H0                 'Dummy
                    Header(7) = &H1                 '���� �޼��� ���� �� ǥ��
                    Header(8) = Led_Repeat          '�ݺ� Ƚ��
                    Header(9) = &H91                ' (1001 0001) B[1:0] - ����ȭ�� ��Ʈũ�� 16 font / B[5:4] - ȭ��ǥ�� ON / B[6:7] - ���� ǥ�� ���� 2 = ���ι���
                    Header(10) = &H0                '��� ���� ���� ����
                    Header(11) = &H0                'Dummy
                    Header(12) = &H0                '����ȭ�� ȿ���� : ȿ������
                    Header(13) = Led_Show           '&H1    ' ����ȭ�� ȿ���� : �����̵�
                    Header(14) = Led_Speed          '&H1E   'ȿ�� �ӵ�
                    Header(15) = Led_StopTime       '&H0    '���� �ð� ����
                    Header(16) = &H0                '���� ǥ�� ��ġ : 0 ��
        
        
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
                        Color_Up(i) = Up_Color    ' &H31 : ���� / 32 : ��� / 33 : �����
                    Next i
                    For i = 0 To Bigger_Len
                        Color_Down(i) = Down_Color    '&H31 : ���� / 32 : ��� / 33 : �����
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
                                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(0).SendData GloDisp_BData1
                                            Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                            End Select
        
                        Case 1
                            Select Case LANE2_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(1).State <> sckClosed) Then
                                                .Disp1_sock(1).Close
                                            End If
                                            .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(1).SendData GloDisp_BData2
                                            Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
        
                            End Select
        
                        Case 2
                             Select Case LANE3_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(2).State <> sckClosed) Then
                                                .Disp1_sock(2).Close
                                            End If
                                            .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(2).SendData GloDisp_BData3
                                            Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
        
                            End Select
        
                        Case 3
                            Select Case LANE4_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(3).State <> sckClosed) Then
                                                .Disp1_sock(3).Close
                                            End If
                                            .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(3).SendData GloDisp_BData4
                                            Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
        
                            End Select
        
                        Case 4
                            Select Case LANE5_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(4).State <> sckClosed) Then
                                                .Disp1_sock(4).Close
                                            End If
                                            .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(4).SendData GloDisp_BData5
                                            Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
        
                            End Select
        
                        Case 5
                            Select Case LANE6_DisplayMode
                                   Case "0" 'Tcp Ip
                                            If (.Disp1_sock(5).State <> sckClosed) Then
                                                .Disp1_sock(5).Close
                                            End If
                                            .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
        
                                   Case "1" 'UDP
                                            .Disp1_sock(5).SendData GloDisp_BData6
                                            Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
        
                            End Select
                    End Select
                End With
        
        Exit Sub
Err_p:
    
        

    ElseIf (Glo_Display = "������(Ǯ�÷�)_FW7") Then
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
            
            If (Glo_Display_Direct = "����") Then
                DoEvents
                Call GL_Emergency_Horizontal(D1, D2, Led_up_color, Led_down_color, IN_OUT)  'New ������ ���� ����
            Else
                DoEvents
                Call GL_Emergency_Vertical_First(D1, D2, Led_up_color, Led_down_color, IN_OUT) 'New ������ ���� ����
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
'        GlO_TcpDataGate = Chr$(2) & "2" & Chr$(3) & Chr$(13) & Chr$(2) & "2" & Chr$(3) '���ǹ��� ��Ŷ(����� ����� ����)
'    Else
'        GlO_TcpDataGate = Chr$(2) & "1" & Chr$(3) & Chr$(13) & Chr$(2) & "1" & Chr$(3) '���ǹ��� ��Ŷ(����� ����� ����)
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
'                                Call DataLogger("[GATE AWAKE TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
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
'                                Call DataLogger("[GATE AWAKE UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                            Else
'                                Call DataLogger("[Get Frame AWAKE UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
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

    
        If (Glo_LPRBoard = "�����") Then
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
            
            
        ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
                If (RNum = 0) Then
                    GlO_TcpDataGate = Chr$(2) & "GATE UP" & Chr$(3) '���ܱ����(���ܱ� ��Ʈ�ѷ� ��������:FPtech�� �������� ������)(Test:Debug��)
                Else
                    GlO_TcpDataGate = "GETFRAME" 'ĸ��(�ڵ��̳뿡�� �����ؾ� ��)
                End If
        End If
            
            
            Select Case gateNo
                Case 0
                    Select Case LANE1_DeviceMode
                           Case "0" 'Tcp Ip
                                    
                                    Glo_Gate_ReconnCnt(0) = 0 '������ ī��Ʈ �ʱ�ȭ

                                    If (.Gate1_sock(0).State <> sckConnected) Then
                                        .Gate1_sock(0).Close
                                        
                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        .Gate1_sock(0).Connect LANE1_DeviceIP, LANE1_RelayPort
                                    
                                    ElseIf (.Gate1_sock(0).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP ����]  �غ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        Dim bData() As Byte
                                        ReDim bData(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(0).State = sckConnected) Then
                                            .Gate1_sock(0).SendData bData
                                        Else
                                            Call DataLogger("[GATE TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(0).State)
                                            Call DebugLogger("[GATE TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(0).State)
                                        End If

                                    Else
                                        Call DataLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(0).State)
                                        Call DebugLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(0).State)
                                    End If
                                    
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
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
                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
'                                    If (RNum = 0) Then '���ܱ����
'                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
'                                        .Gate1_sock(0).RemoteHost = LANE1_DeviceIP
'                                        .Gate1_sock(0).RemotePort = LANE1_RelayPort
'                                    Else    'ĸ��
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_RelayPort)
'                                        .Gate1_sock(0).RemoteHost = LANE1_DispIP '����ݺ��� IP:���ܱ�,
'                                        .Gate1_sock(0).RemotePort = LANE1_RelayPort
'                                    End If
'                                    .Gate1_sock(0).SendData (GlO_TcpDataGate)
'                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 1
                    Select Case LANE2_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(1) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(1).State <> sckConnected) Then
                                        .Gate1_sock(1).Close
                                        
                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        .Gate1_sock(1).Connect LANE2_DeviceIP, LANE2_RelayPort
                                    
                                    ElseIf (.Gate1_sock(1).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP ����]  �غ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        Dim bData1() As Byte
                                        ReDim bData1(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData1 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(1).State = sckConnected) Then
                                            .Gate1_sock(1).SendData bData1
                                        Else
                                            Call DataLogger("[GATE TCP/IP ����]  ���� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", ���ϻ��� = " & .Gate1_sock(1).State)
                                            Call DebugLogger("[GATE TCP/IP ����]  ���� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", ���ϻ��� = " & .Gate1_sock(1).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(1).State)
                                        Call DebugLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(1).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
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
                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 2
                    Select Case LANE3_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(2) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(2).State <> sckConnected) Then
                                        .Gate1_sock(2).Close
                                        
                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        .Gate1_sock(2).Connect LANE3_DeviceIP, LANE3_RelayPort
                                    
                                    ElseIf (.Gate1_sock(2).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP ����]  �غ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        Dim bData2() As Byte
                                        ReDim bData2(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData2 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(2).State = sckConnected) Then
                                            .Gate1_sock(2).SendData bData2
                                        Else
                                            Call DataLogger("[GATE TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(2).State)
                                            Call DebugLogger("[GATE TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(2).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(2).State)
                                        Call DebugLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(2).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
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
                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                
                Case 3
                    Select Case LANE4_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(3) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(3).State <> sckConnected) Then
                                        .Gate1_sock(3).Close
                                        
                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        .Gate1_sock(3).Connect LANE4_DeviceIP, LANE4_RelayPort
                                    
                                    ElseIf (.Gate1_sock(3).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP ����]  �غ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        Dim bData3() As Byte
                                        ReDim bData3(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData3 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(3).State = sckConnected) Then
                                            .Gate1_sock(3).SendData bData3
                                        Else
                                            Call DataLogger("[GATE TCP/IP ����]  ���� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", ���ϻ��� = " & .Gate1_sock(3).State)
                                            Call DebugLogger("[GATE TCP/IP ����]  ���� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", ���ϻ��� = " & .Gate1_sock(3).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(3).State)
                                        Call DebugLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(3).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
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
                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                    
                Case 4
                    Select Case LANE5_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(4) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(4).State <> sckConnected) Then
                                        .Gate1_sock(4).Close
                                    
                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        .Gate1_sock(4).Connect LANE5_DeviceIP, LANE5_RelayPort
                                    
                                    ElseIf (.Gate1_sock(4).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP ����]  �غ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        Dim bData4() As Byte
                                        ReDim bData4(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData4 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(4).State = sckConnected) Then
                                            .Gate1_sock(4).SendData bData4
                                        Else
                                            Call DataLogger("[GATE TCP/IP ����]  ���� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", ���ϻ��� = " & .Gate1_sock(4).State)
                                            Call DebugLogger("[GATE TCP/IP ����]  ���� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", ���ϻ��� = " & .Gate1_sock(4).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(4).State)
                                        Call DebugLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(4).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
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
                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 5
                    Select Case LANE6_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(5) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(5).State <> sckConnected) Then
                                        .Gate1_sock(5).Close
                                    
                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        .Gate1_sock(5).Connect LANE6_DeviceIP, LANE6_RelayPort
                                    
                                    ElseIf (.Gate1_sock(5).State = sckConnected) Then
                                        Call DataLogger("[GATE TCP/IP ����]  �غ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        Dim bData5() As Byte
                                        ReDim bData5(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData5 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(5).State = sckConnected) Then
                                            .Gate1_sock(5).SendData bData5
                                        Else
                                            Call DataLogger("[GATE TCP/IP ����]  ���� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", ���ϻ��� = " & .Gate1_sock(5).State)
                                            Call DebugLogger("[GATE TCP/IP ����]  ���� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", ���ϻ��� = " & .Gate1_sock(5).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(5).State)
                                        Call DebugLogger("[GATE TCP/IP ���� ����]  : " & .Gate1_sock(5).State)
                                    End If
                                    
'                                    If (RNum = 0) Then
'                                        Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
'                                    Else
'                                        Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
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
                                        Call DataLogger("[GATE UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                    Else
                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
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

    
        If (Glo_LPRBoard = "�����") Then
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
            
            
        ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
                If (RNum = 0) Then
                    GlO_TcpDataGate = Chr$(2) & "GATE DOWN" & Chr$(3) '���ܱ����(Test:debug��)
                Else
                    GlO_TcpDataGate = "GETFRAME" 'ĸ��
                End If
        End If
            
            
            Select Case gateNo
                Case 0
                    Select Case LANE1_DeviceMode
                           Case "0" 'Tcp Ip
                                    
                                    Glo_Gate_ReconnCnt(0) = 0 '������ ī��Ʈ �ʱ�ȭ

                                    If (.Gate1_sock(0).State <> sckConnected) Then
                                        .Gate1_sock(0).Close
                                        
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        .Gate1_sock(0).Connect LANE1_DeviceIP, LANE1_RelayPort
                                    
                                    ElseIf (.Gate1_sock(0).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �غ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                        Dim bData() As Byte
                                        ReDim bData(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(0).State = sckConnected) Then
                                            .Gate1_sock(0).SendData bData
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(0).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(0).State)
                                        End If

                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(0).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(0).State)
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
                                        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 1
                    Select Case LANE2_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(1) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(1).State <> sckConnected) Then
                                        .Gate1_sock(1).Close
                                        
                                        Call DataLogger("[GATE UP TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        .Gate1_sock(1).Connect LANE2_DeviceIP, LANE2_RelayPort
                                    
                                    ElseIf (.Gate1_sock(1).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �غ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                        Dim bData1() As Byte
                                        ReDim bData1(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData1 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(1).State = sckConnected) Then
                                            .Gate1_sock(1).SendData bData1
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", ���ϻ��� = " & .Gate1_sock(1).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort & ", ���ϻ��� = " & .Gate1_sock(1).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(1).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(1).State)
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
                                        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 2
                    Select Case LANE3_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(2) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(2).State <> sckConnected) Then
                                        .Gate1_sock(2).Close
                                        
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        .Gate1_sock(2).Connect LANE3_DeviceIP, LANE3_RelayPort
                                    
                                    ElseIf (.Gate1_sock(2).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �غ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                        Dim bData2() As Byte
                                        ReDim bData2(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData2 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(2).State = sckConnected) Then
                                            .Gate1_sock(2).SendData bData2
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(2).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort & ", ���ϻ��� = " & .Gate1_sock(2).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(2).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(2).State)
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
                                        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                
                Case 3
                    Select Case LANE4_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(3) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(3).State <> sckConnected) Then
                                        .Gate1_sock(3).Close
                                        
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        .Gate1_sock(3).Connect LANE4_DeviceIP, LANE4_RelayPort
                                    
                                    ElseIf (.Gate1_sock(3).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �غ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                        Dim bData3() As Byte
                                        ReDim bData3(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData3 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(3).State = sckConnected) Then
                                            .Gate1_sock(3).SendData bData3
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", ���ϻ��� = " & .Gate1_sock(3).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort & ", ���ϻ��� = " & .Gate1_sock(3).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(3).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(3).State)
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
                                        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
        
                    End Select
                    
                Case 4
                    Select Case LANE5_DeviceMode
                           Case "0" 'Tcp Ip
                                    Glo_Gate_ReconnCnt(4) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(4).State <> sckConnected) Then
                                        .Gate1_sock(4).Close
                                    
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        .Gate1_sock(4).Connect LANE5_DeviceIP, LANE5_RelayPort
                                    
                                    ElseIf (.Gate1_sock(4).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �غ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                        Dim bData4() As Byte
                                        ReDim bData4(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData4 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(4).State = sckConnected) Then
                                            .Gate1_sock(4).SendData bData4
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", ���ϻ��� = " & .Gate1_sock(4).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort & ", ���ϻ��� = " & .Gate1_sock(4).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(4).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(4).State)
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
                                        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                                    End If
                                    Call None_Delay_Time(0.1)
                    End Select
                
                Case 5
                    Select Case LANE6_DeviceMode
                           Case "0" 'Tcp Ip

                                    Glo_Gate_ReconnCnt(5) = 0 '������ ī��Ʈ �ʱ�ȭ
                                    
                                    If (.Gate1_sock(5).State <> sckConnected) Then
                                        .Gate1_sock(5).Close
                                    
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        .Gate1_sock(5).Connect LANE6_DeviceIP, LANE6_RelayPort
                                    
                                    ElseIf (.Gate1_sock(5).State = sckConnected) Then
                                        Call DataLogger("[GATE DOWN TCP/IP ����]  �غ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                        Dim bData5() As Byte
                                        ReDim bData5(Len(GlO_TcpDataGate) - 1) As Byte
                                        bData5 = StrConv(GlO_TcpDataGate, vbFromUnicode)

                                        If (.Gate1_sock(5).State = sckConnected) Then
                                            .Gate1_sock(5).SendData bData5
                                        Else
                                            Call DataLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", ���ϻ��� = " & .Gate1_sock(5).State)
                                            Call DebugLogger("[GATE DOWN TCP/IP ����]  ���� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort & ", ���ϻ��� = " & .Gate1_sock(5).State)
                                        End If
                                    
                                    Else
                                        Call DataLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(5).State)
                                        Call DebugLogger("[GATE DOWN TCP/IP ���� ����]  : " & .Gate1_sock(5).State)
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
                                        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                                    Else
'                                        Call DataLogger("[Get Frame UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
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
                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(0).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(0).PortOpen = True) Then
'                                .MSCommDisp(0).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c ���� PORT = " & LANE1_RelayComPort)
'                            End If
            End Select
        
        Case 1
            Select Case LANE2_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(1).State <> sckClosed) Then
                                .Disp1_sock(1).Close
                            End If
                            .Disp1_sock(1).Connect LANE2_DeviceIP, LANE2_DispPort
                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(1).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(1).PortOpen = True) Then
'                                .MSCommDisp(1).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c ���� PORT = " & LANE2_RelayComPort)
'                            End If
            End Select
    
        Case 2
            Select Case LANE3_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(2).State <> sckClosed) Then
                                .Disp1_sock(2).Close
                            End If
                            .Disp1_sock(2).Connect LANE3_DeviceIP, LANE3_DispPort
                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(2).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(2).PortOpen = True) Then
'                                .MSCommDisp(2).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c ���� PORT = " & LANE3_RelayComPort)
'                            End If
            End Select
        
        
        Case 3
            Select Case LANE4_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(3).State <> sckClosed) Then
                                .Disp1_sock(3).Close
                            End If
                            .Disp1_sock(3).Connect LANE4_DeviceIP, LANE4_DispPort
                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(3).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(3).PortOpen = True) Then
'                                .MSCommDisp(3).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c ���� PORT = " & LANE4_RelayComPort)
'                            End If
            End Select
            
        Case 4
            Select Case LANE5_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(4).State <> sckClosed) Then
                                .Disp1_sock(4).Close
                            End If
                            .Disp1_sock(4).Connect LANE5_DeviceIP, LANE5_DispPort
                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(4).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(4).PortOpen = True) Then
'                                .MSCommDisp(4).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c ���� PORT = " & LANE5_RelayComPort)
'                            End If
            End Select
            
        Case 5
            Select Case LANE6_DeviceMode
                   Case "0" 'Tcp Ip
                            If (.Disp1_sock(5).State <> sckClosed) Then
                                .Disp1_sock(5).Close
                            End If
                            .Disp1_sock(5).Connect LANE6_DeviceIP, LANE6_DispPort
                            Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_DispPort)
                            Call Delay_Time(0.1)
                   
                   Case "1" 'UDP
                            .Disp1_sock(5).SendData GlO_TcpDataDisp
                            Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_DispPort)
                   
'                   Case "2" 'Serial Rs-232c
'                            If (.MSCommDisp(5).PortOpen = True) Then
'                                .MSCommDisp(5).Output = GlO_TcpDataDisp
'                                Call DataLogger("[DISP ] RS-232c ���� PORT = " & LANE6_RelayComPort)
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
        Header(3) = &H0     '�����ͱ���
        Header(4) = &H11    '�����ͱ���
        Header(5) = &H94    'Ŀ�ǵ�
        Header(6) = &H1     '������
        Header(7) = &H0
        Header(8) = &H63    '�����ü �÷��÷�
        Header(9) = &H0
        Header(10) = &H0
        Header(11) = &H3
        Header(12) = &H1
        Header(13) = &H0
        Header(14) = &H0
        Header(15) = &H0    'ȿ���ӵ�
        Header(16) = &H8    'ǥ�ýð�
        Header(17) = &H0
        Header(18) = &H0
        Header(19) = &H0
        Header(20) = &H4
        Header(21) = &H0


        ColorArr(0) = Color    ' &H31 : ���� / 32 : ��� / 33 : �����//  1:��
        ColorArr(1) = Color    ' &H31 : ���� / 32 : ��� / 33 : �����
        ColorArr(2) = Color    ' &H31 : ���� / 32 : ��� / 33 : �����
        ColorArr(3) = Color    ' &H31 : ���� / 32 : ��� / 33 : �����

        ReDim Data(4) As Byte

        AsciiStr = StrConv(D1, vbFromUnicode)
        AsciiLen = UBound(AsciiStr)
        
        For i = 3 To 0 Step -1 '1��2�� ������
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
'                Call DataLogger("[������ TCP/IP ����] �õ� IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
'                .ParkFullLightS_sock.Connect Glo_ParkFullLIGHT_IP, Glo_ParkFullLIGHT_PORT
'            Else
'                Call DataLogger("[������ TCP/IP ����] �غ� IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
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
                            Call DataLogger("[������ DISP TCP/IP ����]  �õ� IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
                   
                   Case "1" 'UDP
                            .ParkFullLightS_sock.SendData GlO_ParkFullLight_BData
                            Call DataLogger("[������ DISP UDP ����]  IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
            End Select
            
            Call None_Delay_Time(0.1)
        End With


Exit Sub

Err_p:

'Debug.Print Err.Description
Call DebugLogger("[ParkFullLight Err] " & Err.Description)


End Sub

'��޹���:����
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
        Head_Up(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�
        
        'Up �Ӽ� ����
        Head_Up(5) = &H94    '����
        Head_Up(6) = &H0     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Up(7) = &H0     '�ؼ��ǹ�ȣ(0)
        Head_Up(8) = &H1    'ǥ������(H63:���ѹݺ�)
        Head_Up(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��, 00:���� ǥ�ù��� ���� �� ǥ��
        Head_Up(10) = &H0    '��� ���� ���� ����
        Head_Up(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        Head_Up(12) = &H1    '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
        Head_Up(13) = &H1    '����ȿ��
        Head_Up(14) = &H0    '����ȿ��:&H0, ������� ����
        'Head_Up(15) = &H14               'ȿ���ӵ�:�Ϲ������� H14(20)���� ������
        Head_Up(15) = &H0                'ȿ���ӵ�:�Ϲ����� �ӵ����� ���� ���� H1E(30)���� ����,
        
        'Head_Up(16) = &H4                '�������ð�:4��( 8 x 0.5��), �ؼ��Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������
        'Head_Up(16) = "&H" & Hex(enumDISP_EMG_TIME.e10sec)
        Head_Up(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime * 2)
        
        Head_Up(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(18) = &H0                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(19) = &H18               'X�� ������:0�ȼ�(���Ǻи��� ��� �����), H18:96�ȼ�
        Head_Up(20) = &H8                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(21) = &H0                '����̹��� ����:0(������)
        'Up �Ӽ� ��
        
        
        'Down �Ӽ� ����
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�
        
        'Down �Ӽ� ����
        Head_Down(5) = &H94    '����
        Head_Down(6) = &H0     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Down(7) = &H1     '�ؼ��ǹ�ȣ(0)
        Head_Down(8) = &H1    'ǥ������(H63:���ѹݺ�)
        Head_Down(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��, 00:���� ǥ�ù��� ���� �� ǥ��
        Head_Down(10) = &H0    '��� ���� ���� ����
        Head_Down(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        Head_Down(12) = &H1    '����ȿ�� => �̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)
        Head_Down(13) = &H1    '����ȿ��
        Head_Down(14) = &H0    '����ȿ��:&H0, ������� ����
        'Head_Down(15) = &H14               'ȿ���ӵ�:�Ϲ������� H14(20)���� ������
        Head_Down(15) = &H0                'ȿ���ӵ�:��޹��������� ���ǹ���(0 �Ǵ� FF���� ������)
        
        'Head_Down(16) = &H8                '�������ð�:4��( 8 x 0.5��), �ؼ��Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������
        'Head_Down(16) = "&H" & Hex(enumDISP_EMG_TIME.e10sec)
        Head_Down(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime * 2)
        
        Head_Down(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(18) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(19) = &H18               'X�� ������:0�ȼ�(���Ǻи��� ��� �����), H18:96�ȼ�
        Head_Down(20) = &H8                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(21) = &H0                '����̹��� ����:0(������)
        'Down �Ӽ� ��
        
        
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
        '����(����)
        For i = 0 To UBound(D)
            sHEX_Up(i) = "&H" & Hex(D(i))
        Next i
        
        Second_Str = D2
        D = StrConv(Second_Str, vbFromUnicode)
        '�Ʒ���(����)
        For i = 0 To UBound(D)
            sHEX_Down(i) = "&H" & Hex(D(i))
        Next i
        
        Finish(0) = &H10
        Finish(1) = &H3
        
        Dim data_len  As Integer
        data_len = (UBound(Head_Up) + 1 - 5) + (UBound(Color_Up) + 1) + (UBound(sHEX_Up) + 1)
        Head_Up(4) = "&H" & Hex(data_len)   '������ ����
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
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
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

        '����, �Ʒ��� ���ڿ� �����߿��� ���� �� ���� ã��
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
        
'        '����, �Ʒ��� ���ڿ� ���� ���� ����
'''        If Up_Len > Down_Len Then
'''            For g = 1 To (Up_Len - Down_Len)
'''                D2 = D2 + " "
'''            Next g
'''        Else
'''            For g = 1 To (Down_Len - Up_Len)
'''                D1 = D1 + " "
'''            Next g
'''        End If
        
        '�ƽ�Ű���� �� ���(�����ڵ�� �����Ҷ� �ƽ�Ű���� �� ��ŭ �� ������ ��)
'        iAscCount = 0
'        D = StrConv(D1, vbFromUnicode)
'        For i = 0 To UBound(D)
'            If (D(i) >= 32 And D(i) <= 126) Then '1byte �ƽ�Ű����
'                iAscCount = iAscCount + 1
'            End If
'        Next
        
        On Error GoTo Err_p
        
        Head_Up(0) = &H10    'DLE
        Head_Up(1) = &H2     'STX
        Head_Up(2) = &H0     'DST
        Head_Up(3) = &H0     'LEN
        Head_Up(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�

        'Up �Ӽ� ����
        Head_Up(5) = &H94    '����
        Head_Up(6) = &H0     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Up(7) = &H0     '�ؼ��ǹ�ȣ
        Head_Up(8) = &H2    'ǥ������(���ѹݺ�)
        Head_Up(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��, 00:���� ǥ�ù��� ���� �� ǥ��
        Head_Up(10) = &H0    '��� ���� ���� ����
        Head_Up(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        Head_Up(12) = &H1    '����ȿ�� => {����:&H1}, {�̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)}
        Head_Up(13) = &H1    '����ȿ�� => {����:&H1}, {�̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)}
        Head_Up(14) = &H0    '����ȿ��:&H0, ������� ����
        Head_Up(15) = &H14   'ȿ���ӵ�:�Ϲ������� H14(20)���� ������(10,20,30,..90:���ڰ� �������� ������)
        'Head_Up(15) = &H0   'ȿ���ӵ�:�Ϲ������� H1E(30)���� ������(10,20,30,..90:���ڰ� �������� ������)

        'Head_Up(16) = &H4                '�������ð�: ���Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������(�������޼��������� �ǹ̾��� ����)
        'Head_Up(16) = "&H" & Hex(enumDISP_EMG_TIME.e3sec)        '�����ð�
        Head_Up(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime) '�����ð�
        
        Head_Up(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(18) = &H0                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(19) = &H18               'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(20) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Up(21) = &H0                '����̹��� ����:0(������)
        'Up �Ӽ� ��


        'Down �Ӽ� ����
        Head_Down(0) = &H10    'DLE
        Head_Down(1) = &H2     'STX
        Head_Down(2) = &H0     'DST
        Head_Down(3) = &H0     'LEN
        Head_Down(4) = 0       '��������ü����(�Ӽ�����+���ڱ���+�÷�����) - �Ʒ��ʿ��� ���� �� �Է�

        Head_Down(5) = &H94    '����
        Head_Down(6) = &H0     '��������: ���(�ǽð��޼���) 00 / �Ϲ�(�������޼���) 01
        Head_Down(7) = &H1     '���ǹ�ȣ(1)
        Head_Down(8) = &H2    'ǥ������(���ѹݺ�)
        Head_Down(9) = &H1     '01:���� ǥ�ù��� ���� �� ���ǥ��
        Head_Down(10) = &H0    '��� ���� ���� ����
        Head_Down(11) = &H3    '��Ʈũ��:16x16�ȼ�(��, ����/���ڴ� 8X16)
        Head_Down(12) = &H1    '����ȿ�� => {����:&H1}, {�̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)}
        Head_Down(13) = &H1    '����ȿ�� => {����:&H1}, {�̵��ϱ�: ����(&H6), ������(&H7), ����(&H8), �Ʒ���(&H9)}
        Head_Down(14) = &H0    '����ȿ��:&H0, ������� ����
        Head_Down(15) = &H14   'ȿ���ӵ�:�Ϲ������� H14(20)���� ������(10,20,30,..90:���ڰ� �������� ������)
        'Head_Down(15) = &H0  'ȿ���ӵ�:�Ϲ������� H1E(30)���� ������(10,20,30,..90:���ڰ� �������� ������)

        'Head_Down(16) = &H4                '�����ð�:4��( 8 x 0.5��), �ؼ��Ǻи��� ��� ��ܼ����� 0, �ϴܼ��ǿ��� ������, �� ������ ��� 0���� ������(�������޼��������� �ǹ̾��� ����)
        'Head_Down(16) = "&H" & Hex(enumDISP_EMG_TIME.e3sec)        '�����ð�
        Head_Down(16) = "&H" & Hex(Glo_Emerg_Vertical_ToggleTime) '�����ð�

        Head_Down(17) = &H0                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(18) = &H4                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����) : 16�ȼ�
        Head_Down(19) = &H18                'X�� ������:0�ȼ�(���Ǻи��� ��� �����)
        Head_Down(20) = &H8                'Y�� ������:0�ȼ�(���Ǻи��� ��� �����) : 32�ȼ�
        Head_Down(21) = &H0                '����̹��� ����:0(������)
        'Down �Ӽ� ��
        
                
        'ReDim Color_Up(Bigger_Len - 1 + iAscCount) As Byte
        'ReDim Color_Down(Bigger_Len - 1 + iAscCount) As Byte
        ReDim Color_Up(Bigger_Len * 2 - 1) As Byte
        ReDim Color_Down(Bigger_Len * 2 - 1) As Byte
        
        For i = 0 To UBound(Color_Up)
            'Debug.Print i
            Color_Up(i) = Nomal_Up_color + 8 '����������� +8
        Next i
        
        For i = 0 To UBound(Color_Down)
            Color_Down(i) = Nomal_Down_color + 8 '����������� +8
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
        Head_Up(4) = "&H" & Hex(data_len)   '������ ����
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
                'Debug.Print "���UP:" & ByteArrayToHex(GloDisp_BData1)
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
                'Debug.Print "���DN:" & ByteArrayToHex(GloDisp_BData1_Down)
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
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub


'������ȣ:��4�ڸ� ���� �и� ���
Public Sub GL_Emergency_Vertical_First(D1 As String, D2 As String, Nomal_Up_color As Byte, Nomal_Down_color As Byte, gateNo As Integer)
    
    'D1:������ȣ���� �ƴ��� Ȯ��
    Dim i As Integer
    Dim sCarNo As String
    Dim sCarNo1 As String
    Dim sCarNo2 As String
    Dim sCarStat As String
    Dim sCarStat1 As String
    Dim sCarStat2 As String
    Dim iCarNoLen As Integer
    Dim iToggleCount As Integer
    iToggleCount = 2 '������ȣ + ó����� 2ȸ ���
    
    iCarNoLen = LenH(D1)
    If ((IsNumeric(Right(D1, 4)) = True) And (iCarNoLen = 8 Or iCarNoLen = 9 Or iCarNoLen = 11 Or iCarNoLen = 12)) Then
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '���1:Ÿ�̸Ӹ� �̿��ؼ� ����� ��� ù��±��� �� 1000ms ��ٷ�����
'''                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = False
'''                '������ȣ
'''                sCarNo = D1
'''                sCarNo1 = LeftH(D1, iCarNoLen - 4)
'''                sCarNo2 = Right(D1, 4)
'''                sCarStat = D2
'''                sCarStat1 = Left(D2, Int(Len(D2) / 2) + Len(D2) Mod 2) '���ڿ�/2 �պκ� + Mod 2 ������ ���ڿ�
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
'''                Glo_Emerg_Vertical(gateNo).ToggleSelect = EnumEmergToggleOrder.enumCarNo  'ó�� ����� ����: ������ȣ
'''
'''
'''                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Interval = 100 'ó�� ����� ������ ��� ����ؾ���
'''                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = True
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '���2:ù����� ���� ó���ϰ�, ���ĺ��ʹ� Ÿ�̸� �̿��ؼ� ���
                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = False
                '������ȣ
                sCarNo = D1
                sCarNo1 = LeftH(D1, iCarNoLen - 4)
                sCarNo2 = Right(D1, 4)
                sCarStat = D2
                sCarStat1 = Left(D2, Int(Len(D2) / 2) + Len(D2) Mod 2) '���ڿ�/2 �պκ� + Mod 2 ������ ���ڿ�
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
                Call GL_Emergency_Vertical(sCarNo1, sCarNo2, Glo_Emerg_Vertical(gateNo).CarNoColor1, Glo_Emerg_Vertical(gateNo).CarNoColor2, gateNo) '������ȣ ��� ���
                Glo_Emerg_Vertical(gateNo).CarNoCount = Glo_Emerg_Vertical_ToggleCount - 1 '���� ��� ��������Ƿ� ī��Ʈ -1 ó����
                Glo_Emerg_Vertical(gateNo).ToggleSelect = EnumEmergToggleOrder.enumCarStat  '������ ����� ����(ó�����) ����

                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Interval = Glo_Emerg_Vertical_ToggleTime * 1000 '���� ms
                FrmTcpServer.Timer_Emerg_Vertical(gateNo).Enabled = True
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Else
    
        '������ȣ�� ��޹���
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
        Head(4) = &H2     '��������ü����

        'Up �Ӽ� ����
        Head(5) = &H41
        Head(6) = &H0
        'Up �Ӽ� ��
        Head(7) = &H10
        Head(8) = &H3
        
        Select Case gateNo
        
            Case 0
                ReDim GloDisp_BData1(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData1(i) = Head(i)
                Next i
                Debug.Print "POWER OFF:" & ByteArrayToHex(GloDisp_BData1) '�ӽ��׽�Ʈ
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
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
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
        Head(4) = &H2     '��������ü����

        'Up �Ӽ� ����
        Head(5) = &H41
        Head(6) = &H1
        'Up �Ӽ� ��
        Head(7) = &H10
        Head(8) = &H3
        
        Select Case gateNo
        
            Case 0
                ReDim GloDisp_BData1(UBound(Head)) As Byte

                For i = 0 To UBound(Head)
                   GloDisp_BData1(i) = Head(i)
                Next i
                Debug.Print "POWER ON:" & ByteArrayToHex(GloDisp_BData1) '�ӽ��׽�Ʈ
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
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(0).SendData GloDisp_BData1
                                    .Disp1_sock(0).SendData GloDisp_BData1_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE1_DispIP & "    PORT = " & LANE1_DispPort)
                    End Select
                
                Case 1
                    Select Case LANE2_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(1).State <> sckClosed) Then
                                        .Disp1_sock(1).Close
                                    End If
                                    .Disp1_sock(1).Connect LANE2_DispIP, LANE2_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(1).SendData GloDisp_BData2
                                    .Disp1_sock(1).SendData GloDisp_BData2_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE2_DispIP & "    PORT = " & LANE2_DispPort)
                           
                    End Select
            
                Case 2
                     Select Case LANE3_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(2).State <> sckClosed) Then
                                        .Disp1_sock(2).Close
                                    End If
                                    .Disp1_sock(2).Connect LANE3_DispIP, LANE3_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(2).SendData GloDisp_BData3
                                    .Disp1_sock(2).SendData GloDisp_BData3_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE3_DispIP & "    PORT = " & LANE3_DispPort)
                           
                    End Select
                
                Case 3
                    Select Case LANE4_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(3).State <> sckClosed) Then
                                        .Disp1_sock(3).Close
                                    End If
                                    .Disp1_sock(3).Connect LANE4_DispIP, LANE4_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(3).SendData GloDisp_BData4
                                    .Disp1_sock(3).SendData GloDisp_BData4_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE4_DispIP & "    PORT = " & LANE4_DispPort)
                           
                    End Select
                
                Case 4
                    Select Case LANE5_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(4).State <> sckClosed) Then
                                        .Disp1_sock(4).Close
                                    End If
                                    .Disp1_sock(4).Connect LANE5_DispIP, LANE5_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(4).SendData GloDisp_BData5
                                    .Disp1_sock(4).SendData GloDisp_BData5_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE5_DispIP & "    PORT = " & LANE5_DispPort)
                           
                    End Select
                    
                Case 5
                    Select Case LANE6_DisplayMode
                           Case "0" 'Tcp Ip
                                    If (.Disp1_sock(5).State <> sckClosed) Then
                                        .Disp1_sock(5).Close
                                    End If
                                    .Disp1_sock(5).Connect LANE6_DispIP, LANE6_DispPort
                                    Call DataLogger("[DISP ����]  �õ� IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                           Case "1" 'UDP
                                    .Disp1_sock(5).SendData GloDisp_BData6
                                    .Disp1_sock(5).SendData GloDisp_BData6_Down
                                    Call DataLogger("[DISP UDP ����]  IP = " & LANE6_DispIP & "    PORT = " & LANE6_DispPort)
                           
                    End Select
            End Select
        End With

        
        Call None_Delay_Time(0.1)
        

Exit Sub

Err_p:



End Sub

