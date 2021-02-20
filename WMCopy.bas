Attribute VB_Name = "WMCopy"
Option Explicit

Public Const WM_LANE1_HANDLE = "01"
Public Const WM_LANE2_HANDLE = "02"
Public Const WM_LANE3_HANDLE = "03"
Public Const WM_LANE4_HANDLE = "04"
Public Const WM_LANE5_HANDLE = "05"
Public Const WM_LANE6_HANDLE = "06"
Public Const WM_LANE7_HANDLE = "07"
Public Const WM_LANE8_HANDLE = "08"

Public Const WM_LANE1_CARNUM = "11"
Public Const WM_LANE2_CARNUM = "12"
Public Const WM_LANE3_CARNUM = "13"
Public Const WM_LANE4_CARNUM = "14"
Public Const WM_LANE5_CARNUM = "15"
Public Const WM_LANE6_CARNUM = "16"
Public Const WM_LANE7_CARNUM = "17"
Public Const WM_LANE8_CARNUM = "18"


Public Const WM_LANE1_WATCHDOG_ACK = "21"
Public Const WM_LANE2_WATCHDOG_ACK = "22"
Public Const WM_LANE3_WATCHDOG_ACK = "23"
Public Const WM_LANE4_WATCHDOG_ACK = "24"
Public Const WM_LANE5_WATCHDOG_ACK = "25"
Public Const WM_LANE6_WATCHDOG_ACK = "26"

Public Const WM_LANE1_LOADING = "31"
Public Const WM_LANE2_LOADING = "32"
Public Const WM_LANE3_LOADING = "33"
Public Const WM_LANE4_LOADING = "34"
Public Const WM_LANE5_LOADING = "35"
Public Const WM_LANE6_LOADING = "36"

Public Const WM_LANE1_CAMERA_ERR = "41"
Public Const WM_LANE2_CAMERA_ERR = "42"
Public Const WM_LANE3_CAMERA_ERR = "43"
Public Const WM_LANE4_CAMERA_ERR = "44"
Public Const WM_LANE5_CAMERA_ERR = "45"
Public Const WM_LANE6_CAMERA_ERR = "46"

Public Const WM_HOST_HANDLE = "51"
Public Const WM_FEE1_HANDLE = "52"
Public Const WM_FEE2_HANDLE = "53"

Public Const WM_WATCHDOG_POLL = "99"

Public LANE1_Handle As Long
Public LANE2_Handle As Long
Public LANE3_Handle As Long
Public LANE4_Handle As Long
Public LANE5_Handle As Long
Public LANE6_Handle As Long

Global gHW As Long


'Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
   ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
   ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
   ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMsg Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long



 Public Type COPYDATASTRUCT
              dwData As Long
              cbData As Long
              lpData As Long
 End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Global lpPrevWndProc As Long


Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As _
         Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)

Private Type CRITICAL_SECTION
    Reserved1 As Long
    Reserved2 As Long
    Reserved3 As Long
    Reserved4 As Long
    Reserved5 As Long
    Reserved6 As Long
End Type
Public Glo_CS As CRITICAL_SECTION


Public Sub Hook()
    If (Glo_TestMode = "Y") Then
        Exit Sub
    End If
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    Dim Temp As Long
    If (Glo_TestMode = "Y") Then
        Exit Sub
    End If
    Temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If uMsg = WM_COPYDATA Then
        Call mySub(lParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)

End Function

Sub mySub(lParam As Long)
    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 1024) As Byte
    Dim a As String
    Call CopyMemory(cds, ByVal lParam, Len(cds))

    Select Case cds.dwData
     Case 1
     Case 2
     Case 3
        Call CopyMemory(buf(1), ByVal cds.lpData, cds.cbData)
        a$ = StrConv(buf, vbUnicode)
        a$ = Left$(a$, InStr(1, a$, Chr$(0)) - 1)
        AddLog a$
    End Select
End Sub

Public Function SendMess(ByVal Mess As String, TrHwnd As Long)
    If TrHwnd = 0 Then Exit Function
    
    Dim cds As COPYDATASTRUCT
    Dim ThWnd As Long, Sownd As Long
    Dim buf(1 To 1024) As Byte
    Dim i As Long
    
    Dim strsz As Integer
    Sownd = TrHwnd
    ThWnd = TrHwnd
    strsz = LenB(StrConv(Mess, vbFromUnicode)) '' 한글 2바이트 영문은 1바이트
    Call CopyMemory(buf(1), ByVal Mess, strsz)
    cds.dwData = 3
    cds.cbData = strsz + 1
    cds.lpData = VarPtr(buf(1))
    'i = SendMessage(ThWnd, WM_COPYDATA, Sownd, cds)
    i = SendMsg(ThWnd, WM_COPYDATA, Sownd, cds)
End Function


Public Sub AddLog(str As String)
Dim cmd As String
Dim car_num As String
Dim tmp_str As String
    
    
    

With Jung

If (Len(str) <= 20) Then
    cmd = Mid(str, 1, 2)
    Select Case cmd
        Case WM_LANE1_HANDLE
             LANE1_Handle = Mid(str, 3, Len(str) - 2)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE1 Start] " & LANE1_Handle, 0
            Call DataLogger("LANE1_Handle recieved " & LANE1_Handle & "    " & str)
        Case WM_LANE2_HANDLE
             LANE2_Handle = Mid(str, 3, Len(str) - 2)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE2 Start] " & LANE2_Handle, 0
             Call DataLogger("LANE2_Handle recieved " & LANE2_Handle & "    " & str)
        Case WM_LANE3_HANDLE
             LANE3_Handle = Mid(str, 3, Len(str) - 2)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE3 Start] " & LANE3_Handle, 0
             Call DataLogger("LANE3_Handle recieved " & LANE3_Handle & "    " & str)
        Case WM_LANE4_HANDLE
             LANE4_Handle = Mid(str, 3, Len(str) - 2)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE4 Start] " & LANE4_Handle, 0
             Call DataLogger("LANE4_Handle recieved " & LANE4_Handle & "    " & str)
    End Select
Else
    '데이터 복호화
    tmp_str = DecodeNDE01(str, "www.jawootek.com")
    cmd = Mid(tmp_str, 1, 2)
    Select Case cmd
        Case WM_LANE1_CARNUM
             tmp_str = Mid(tmp_str, 3, LenH(tmp_str) - 3)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE1 DATA] " & LANE1_Handle, 0
            Call DataLogger("LANE1_Handle recieved " & LANE1_Handle & "    " & tmp_str)
            Call UDP_Proc(tmp_str)
            SendMess "ACK", LANE1_Handle
        Case WM_LANE2_CARNUM
             tmp_str = Mid(tmp_str, 3, LenH(tmp_str) - 3)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE2 DATA] " & LANE2_Handle, 0
            Call DataLogger("LANE2_Handle recieved " & LANE2_Handle & "    " & tmp_str)
            Call UDP_Proc(tmp_str)
            SendMess "ACK", LANE2_Handle
        Case WM_LANE3_CARNUM
             tmp_str = Mid(tmp_str, 3, LenH(tmp_str) - 3)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE3 DATA] " & LANE3_Handle, 0
            Call DataLogger("LANE3_Handle recieved " & LANE3_Handle & "    " & tmp_str)
            Call UDP_Proc(tmp_str)
            SendMess "ACK", LANE3_Handle
        Case WM_LANE4_CARNUM
             tmp_str = Mid(tmp_str, 3, LenH(tmp_str) - 3)
             .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LANE4 DATA] " & LANE4_Handle, 0
            Call DataLogger("LANE4_Handle recieved " & LANE4_Handle & "    " & tmp_str)
            Call UDP_Proc(tmp_str)
            SendMess "ACK", LANE4_Handle
    End Select

End If



    
End With
    
End Sub


Public Sub UDP_Proc(sdata As String)
    'Dim sdata As String
    Dim Tmp_Path As String
    Dim i, gateNo As Integer
    Dim carnum As String
    Dim s As String
    
    Dim sLprIP As String
    Dim sLaneInout As String
    Dim sFreePass As String
    Dim sBlackList As String
    Dim sNoRecOpen As String
    Dim sPassDate As String
    Dim sGateOpen As String
    Dim sGateStat As String
    Dim sTaxiPass As String
    Dim sGuestLane As String
    
    Dim stLPR As structLPR
    Dim stGate As structGate
    Dim stSound As structSound
    Dim stEmerg As structEmerg
    Dim iViewGateNo As Integer
    Dim sNoWork As String
    
    Dim sStrLine() As String
    

On Error GoTo Err_P

    

'    If (InStr(sdata, "_") > 0) Then
'    Else
'        Call DataLogger("UDP_Proc  알수없는 데이터 수신 : " & sdata)
'        Exit Sub
'    End If
    
    sStrLine() = Split(sdata, "_")
    If (UBound(sStrLine) <> 3) Then
        Call DebugLogger("[DataCheck] 알수없는 데이터 수신#1 : " & sdata)
        Exit Sub
    End If
    
    
    Call EnterCriticalSection(Glo_CS)
    
    
    i = InStr(1, sdata, "_", 1)
    gateNo = Val(Left(sdata, (i - 1)))
    
    iViewGateNo = gateNo - Glo_GateNo_StartNo
    If (iViewGateNo < 0) Then
        Select Case Glo_Screen_No
            Case 6
                FrmG6_23.LblDBInfo = "LPR GateNo 시작번호 오류(시작번호는 GateNo 보다 작아야합니다)"
            Case 4
                FrmG4Mini.LblDBInfo = "LPR GateNo 시작번호 오류(시작번호는 GateNo 보다 작아야합니다)"
            Case 2
                Jung.LblDBInfo = "LPR GateNo 시작번호 오류(시작번호는 GateNo 보다 작아야합니다)"
            Case 1
                FrmG1.LblDBInfo = "LPR GateNo 시작번호 오류(시작번호는 GateNo 보다 작아야합니다)"
        End Select
        Exit Sub
    End If
    
    'Glo_GateNo = GateNo
    'Select Case GateNo
    Glo_GateNo = iViewGateNo
    Select Case iViewGateNo
        Case 0
            sLprIP = LANE1_LPRIP
            sLaneInout = LANE1_Inout
            sFreePass = Glo_FreePassLane1_YN
            sBlackList = Glo_BlackList1_YN
            sNoRecOpen = Glo_NoRecOpen1_YN
            sTaxiPass = Glo_TAXI1_YN
            sGuestLane = Glo_GUEST_LANE1_YN
            stSound.sSnd_YN = Glo_SOUND_YN
            stSound.sSndReg_YN = Glo_SND_Lane1_Reg_YN
            stSound.sSndGuest_YN = Glo_SND_Lane1_Guest_YN
            stSound.sSndNoRec_YN = Glo_SND_Lane1_NoRec_YN
            stSound.sSndBlackList_YN = Glo_SND_Lane1_BlackList_YN
            stSound.sSndTaxi_YN = Glo_SND_Lane1_Taxi_YN
            stSound.sSndDay_YN = Glo_SND_Lane1_Day_YN
            stSound.sSndRegExpDate_YN = Glo_SND_Lane1_RegExpDate_YN
            sNoWork = Glo_Lane1_NoWork
        Case 1
            sLprIP = LANE2_LPRIP
            sLaneInout = LANE2_Inout
            sFreePass = Glo_FreePassLane2_YN
            sBlackList = Glo_BlackList2_YN
            sNoRecOpen = Glo_NoRecOpen2_YN
            sTaxiPass = Glo_TAXI2_YN
            sGuestLane = Glo_GUEST_LANE2_YN
            stSound.sSnd_YN = Glo_SOUND_YN
            stSound.sSndReg_YN = Glo_SND_Lane2_Reg_YN
            stSound.sSndGuest_YN = Glo_SND_Lane2_Guest_YN
            stSound.sSndNoRec_YN = Glo_SND_Lane2_NoRec_YN
            stSound.sSndBlackList_YN = Glo_SND_Lane2_BlackList_YN
            stSound.sSndTaxi_YN = Glo_SND_Lane2_Taxi_YN
            stSound.sSndDay_YN = Glo_SND_Lane2_Day_YN
            stSound.sSndRegExpDate_YN = Glo_SND_Lane2_RegExpDate_YN
            sNoWork = Glo_Lane2_NoWork
        Case 2
            sLprIP = LANE3_LPRIP
            sLaneInout = LANE3_Inout
            sFreePass = Glo_FreePassLane3_YN
            sBlackList = Glo_BlackList3_YN
            sNoRecOpen = Glo_NoRecOpen3_YN
            sTaxiPass = Glo_TAXI3_YN
            sGuestLane = Glo_GUEST_LANE3_YN
            stSound.sSnd_YN = Glo_SOUND_YN
            stSound.sSndReg_YN = Glo_SND_Lane3_Reg_YN
            stSound.sSndGuest_YN = Glo_SND_Lane3_Guest_YN
            stSound.sSndNoRec_YN = Glo_SND_Lane3_NoRec_YN
            stSound.sSndBlackList_YN = Glo_SND_Lane3_BlackList_YN
            stSound.sSndTaxi_YN = Glo_SND_Lane3_Taxi_YN
            stSound.sSndDay_YN = Glo_SND_Lane3_Day_YN
            stSound.sSndRegExpDate_YN = Glo_SND_Lane3_RegExpDate_YN
            sNoWork = Glo_Lane3_NoWork
        Case 3
            sLprIP = LANE4_LPRIP
            sLaneInout = LANE4_Inout
            sFreePass = Glo_FreePassLane4_YN
            sBlackList = Glo_BlackList4_YN
            sNoRecOpen = Glo_NoRecOpen4_YN
            sTaxiPass = Glo_TAXI4_YN
            sGuestLane = Glo_GUEST_LANE4_YN
            stSound.sSnd_YN = Glo_SOUND_YN
            stSound.sSndReg_YN = Glo_SND_Lane4_Reg_YN
            stSound.sSndGuest_YN = Glo_SND_Lane5_Guest_YN
            stSound.sSndNoRec_YN = Glo_SND_Lane4_NoRec_YN
            stSound.sSndBlackList_YN = Glo_SND_Lane4_BlackList_YN
            stSound.sSndTaxi_YN = Glo_SND_Lane4_Taxi_YN
            stSound.sSndDay_YN = Glo_SND_Lane4_Day_YN
            stSound.sSndRegExpDate_YN = Glo_SND_Lane4_RegExpDate_YN
            sNoWork = Glo_Lane4_NoWork
        Case 4
            sLprIP = LANE5_LPRIP
            sLaneInout = LANE5_Inout
            sFreePass = Glo_FreePassLane5_YN
            sBlackList = Glo_BlackList5_YN
            sNoRecOpen = Glo_NoRecOpen5_YN
            sTaxiPass = Glo_TAXI5_YN
            sGuestLane = Glo_GUEST_LANE5_YN
            stSound.sSnd_YN = Glo_SOUND_YN
            stSound.sSndReg_YN = Glo_SND_Lane5_Reg_YN
            stSound.sSndGuest_YN = Glo_SND_Lane5_Guest_YN
            stSound.sSndNoRec_YN = Glo_SND_Lane5_NoRec_YN
            stSound.sSndBlackList_YN = Glo_SND_Lane5_BlackList_YN
            stSound.sSndTaxi_YN = Glo_SND_Lane5_Taxi_YN
            stSound.sSndDay_YN = Glo_SND_Lane5_Day_YN
            stSound.sSndRegExpDate_YN = Glo_SND_Lane5_RegExpDate_YN
            sNoWork = Glo_Lane5_NoWork
        Case 5
            sLprIP = LANE6_LPRIP
            sLaneInout = LANE6_Inout
            sFreePass = Glo_FreePassLane6_YN
            sBlackList = Glo_BlackList6_YN
            sNoRecOpen = Glo_NoRecOpen6_YN
            sTaxiPass = Glo_TAXI6_YN
            sGuestLane = Glo_GUEST_LANE6_YN
            stSound.sSnd_YN = Glo_SOUND_YN
            stSound.sSndReg_YN = Glo_SND_Lane6_Reg_YN
            stSound.sSndGuest_YN = Glo_SND_Lane6_Guest_YN
            stSound.sSndNoRec_YN = Glo_SND_Lane6_NoRec_YN
            stSound.sSndBlackList_YN = Glo_SND_Lane6_BlackList_YN
            stSound.sSndTaxi_YN = Glo_SND_Lane6_Taxi_YN
            stSound.sSndDay_YN = Glo_SND_Lane6_Day_YN
            stSound.sSndRegExpDate_YN = Glo_SND_Lane6_RegExpDate_YN
            sNoWork = Glo_Lane6_NoWork
    End Select

    If (iViewGateNo >= 0 And iViewGateNo <= 5) Then
        stSound.sSndFName_Reg = Glo_SNDFILE_Reg(iViewGateNo)
        stSound.sSndFName_Guest = Glo_SNDFILE_Guest(iViewGateNo)
        stSound.sSndFName_NoRec = Glo_SNDFILE_NoRec(iViewGateNo)
        stSound.sSndFName_BlackList = Glo_SNDFILE_BlackList(iViewGateNo)
        stSound.sSndFName_Taxi = Glo_SNDFILE_Taxi(iViewGateNo)
        stSound.sSndFName_Day = Glo_SNDFILE_Day(iViewGateNo)
        stSound.sSndFName_RegExpDate = Glo_SNDFILE_RegExpDate(iViewGateNo)
        
        '긴급문구
        stEmerg.sEmergReg = Glo_Str_Reg(iViewGateNo)
        stEmerg.sEmergGuest = Glo_Str_Guest(iViewGateNo)
        stEmerg.sEmergNoRec = Glo_Str_NoRec(iViewGateNo)
        stEmerg.sEmergBlackList = Glo_Str_BlackList(iViewGateNo)
        stEmerg.sEmergTaxi = Glo_Str_Taxi(iViewGateNo)
        stEmerg.sEmergDay = Glo_Str_Day(iViewGateNo)
        stEmerg.sEmergRegExpDate = Glo_Str_RegExpDate(iViewGateNo)
        
        
        If (Glo_Display = "전광판(풀컬러)_FW7") Then
            stEmerg.iDisp1EmergColorReg = Glo_Disp1_Reg(iViewGateNo) '등록차량 첫번째 문구색상
            stEmerg.iDisp2EmergColorReg = Glo_Disp2_Reg(iViewGateNo) '등록차량 두번째 문구색상
            stEmerg.iDisp1EmergColorGuest = Glo_Disp1_Guest(iViewGateNo) '미등록차량 첫번째 문구색상
            stEmerg.iDisp2EmergColorGuest = Glo_Disp2_Guest(iViewGateNo) '미등록차량 두번째 문구색상
            stEmerg.iDisp1EmergColorNoRec = Glo_Disp1_NoRec(iViewGateNo) '미인식차량 첫번째 문구색상
            stEmerg.iDisp2EmergColorNoRec = Glo_Disp2_NoRec(iViewGateNo) '미인식차량 두번째 문구색상
            stEmerg.iDisp1EmergColorBKList = Glo_Disp1_BlackList(iViewGateNo) '블랙리스트 첫번째 문구색상
            stEmerg.iDisp2EmergColorBKList = Glo_Disp2_BlackList(iViewGateNo) '블랙리스트 두번째 문구색상
            stEmerg.iDisp1EmergColorTaxi = Glo_Disp1_Taxi(iViewGateNo) '영업용차량 첫번째 문구색상
            stEmerg.iDisp2EmergColorTaxi = Glo_Disp2_Taxi(iViewGateNo) '영업용차량 두번째 문구색상
            stEmerg.iDisp1EmergColorDay = Glo_Disp1_Day(iViewGateNo) '요일제차량 첫번째 문구색상
            stEmerg.iDisp2EmergColorDay = Glo_Disp2_Day(iViewGateNo) '요일제차량 두번째 문구색상
            stEmerg.iDisp1EmergColorRegExpDate = Glo_Disp1_RegExpDate(iViewGateNo) '등록기간초과차량 첫번째 문구색상
            stEmerg.iDisp2EmergColorRegExpDate = Glo_Disp2_RegExpDate(iViewGateNo) '등록기간초과차량 두번째 문구색상
            
        ElseIf (Glo_Display = "전광판" Or Glo_Display = "전광판(풀컬러)") Then  '황:2, 초:1, 적:0
            stEmerg.iDisp1EmergColorReg = 1
            stEmerg.iDisp2EmergColorReg = 2
            stEmerg.iDisp1EmergColorGuest = 0
            stEmerg.iDisp2EmergColorGuest = 2
            stEmerg.iDisp1EmergColorNoRec = 0
            stEmerg.iDisp2EmergColorNoRec = 2
            stEmerg.iDisp1EmergColorBKList = 0
            stEmerg.iDisp2EmergColorBKList = 2
            stEmerg.iDisp1EmergColorTaxi = 0
            stEmerg.iDisp2EmergColorTaxi = 2
            stEmerg.iDisp1EmergColorDay = 0
            stEmerg.iDisp2EmergColorDay = 2
            stEmerg.iDisp1EmergColorRegExpDate = 0
            stEmerg.iDisp2EmergColorRegExpDate = 2
        End If
    End If
    
    
    
    
    
    
    
'    s = InStr(4, sdata, "_", 1)
'    carnum = Mid(sdata, (i + 1), (s - i - 1))
'    Glo_CarNum = carnum
'    i = Len(sdata)
'    Tmp_Path = Mid(sdata, (s + 1), i)


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 수정 시작, 차량번호 재인식 시작
    s = InStr(4, sdata, "_", 1)
    carnum = Mid(sdata, (i + 1), (s - i - 1))
    i = Len(sdata)
    Tmp_Path = Mid(sdata, (s + 1), i)

'''    If (Glo_ReANPR_YN = "Y") Then
'''        Dim NewCarno, OldCarno As String
'''
'''        OldCarno = carnum
'''        NewCarno = GetPlateNumber(Tmp_Path)
'''
'''        If (NewCarno = "XXXXXXX") Then
'''        Else
'''            carnum = NewCarno
'''            Call DataLogger("차량번호재인식: " & OldCarno & " => " & NewCarno)
'''        End If
'''    End If
    Glo_CarNum = carnum
    ' 수정 끝
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    
    
    sPassDate = Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format(Timer * 1000 Mod 1000, "000")
    
    'stLPR.sGateNo = GateNo
    stLPR.sGateNo = iViewGateNo
    stLPR.sLprIP = sLprIP
    stLPR.sLaneInout = sLaneInout
    stLPR.sFreePass = sFreePass
    stLPR.sBlackList = sBlackList
    stLPR.sNoRecOpen = sNoRecOpen
    stLPR.sTaxiPass = sTaxiPass
    stLPR.sPassDate = sPassDate
    stLPR.sNoWork = sNoWork
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 수정 시작, 위즈넷 깨우기 위한 용도
'''    If (Glo_Device_Awake = "Y") Then
'''        If (iViewGateNo >= 0 And iViewGateNo < 6) Then
'''            Call Relay_Alive(0, iViewGateNo)
'''        End If
'''    End If
    ' 수정 끝
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    If (DB_Connect_F = True) Then   'DB 접속상태에서만 처리합니다.

        Call LPRIn_Proc(carnum, Tmp_Path, sLprIP, sLaneInout, sFreePass, sBlackList, sNoRecOpen, sTaxiPass, gateNo, sPassDate, sGateOpen, sGateStat, stSound, sNoWork, stEmerg)

        'stGate : 처리후 게이트 상태값을 리턴받습니다.
        'Call LPRIn_Proc(carnum, Tmp_Path, stLPR, stGate)
        '스크린 수에 따라서 분기
        If (iViewGateNo >= 0 And iViewGateNo < 6) Then
        
            If (Glo_Screen_No = 6) Then
                'If (GateNo < Glo_Screen_No) Then
                    'Call G6_23Show(carnum, GateNo, sPassDate)
                If (iViewGateNo < Glo_Screen_No) Then
                    Call G6_23Show(carnum, iViewGateNo, sPassDate)
                End If
            ElseIf (Glo_Screen_No = 4) Then
                If (iViewGateNo < Glo_Screen_No) Then
                    Call G4Mini_4INShow(carnum, iViewGateNo, sPassDate)
                End If
            ElseIf (Glo_Screen_No = 2) Then
                If (iViewGateNo < Glo_Screen_No) Then
                    Call Jung_Show(carnum, iViewGateNo, sPassDate)
                End If
            ElseIf (Glo_Screen_No = 1) Then
                If (iViewGateNo < Glo_Screen_No) Then
                    Call G1_Show(carnum, iViewGateNo, sPassDate)
                End If
            End If
        End If

        If (Glo_ParkFull_YN = "Y") Then
            Call ParkFull_Show
        End If

    Else
        If (iViewGateNo >= 0 And iViewGateNo < 6) Then
            FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다, 입출차기록 저장실패_차단기 자동 열림", 0
            Call DBaseCheck
            Call Relay_Out(0, iViewGateNo)   ' DB 미접속시 무조건 차단기 열림
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '방문객 처리 시작
    If (iViewGateNo >= 0 And iViewGateNo < 6) Then
'            If (sGuestLane = "Y") Then
'                If (sLaneInout = "입구") Then
            If (sGuestLane = "Y" And sGateStat <> "등록차량") Then
                    Call Glo_FrmGuest(iViewGateNo).Guest_Inputdata(carnum, sPassDate, Tmp_Path)
                    Glo_FrmGuest(iViewGateNo).WindowState = vbNormal
                    
                    '방문객 자동DB저장(입출차) - DB자동저장
                    If (sFreePass = "Y") Then
                        Call Glo_FrmGuest(iViewGateNo).Guest_In_Auto_Proc(carnum, sPassDate, Tmp_Path, sLaneInout)
                    End If
'                End If
            End If
            
            If (Glo_Guest_YN = "Y") Then '방문객레인 한개이상 사용할 경우
                If (sLaneInout = "출구") Then
                    'Call Glo_FrmGuest(iViewGateNo).Guest_Out_Auto_Proc(carnum, sPassDate, Tmp_Path, sFreePass, sLaneInout)
                    
                    '방문차량 주차시간 계산 후 DB저장
                    Call FormGuest1.Guest_Out_Auto_Proc(CStr(iViewGateNo), carnum, sPassDate, Tmp_Path, sLaneInout)
                End If
            End If
    End If
    '방문객 처리 끝
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If (iViewGateNo >= 0 And iViewGateNo < 6) Then
        If Glo_RemoteS_YN = "Y" Then
            'Glo_Remote_Str = (GateNo + Glo_RemoteS_ScrPos) & "_" & carnum & "_" & sPassDate
            'Glo_Remote_Str = gateNo & "_" & carnum & "_" & sPassDate
            Glo_Remote_Str = iViewGateNo & "_" & carnum & "_" & sPassDate
            FrmTcpServer.RemoteS_sock.SendData (Glo_Remote_Str)
            Call DataLogger("[Remote UDP 전송]  DATA = " & Glo_Remote_Str)
        End If
    End If

    If (iViewGateNo >= 0 And iViewGateNo < 6) Then
        If MVR_YN = "Y" Then
            MVR_Str = (iViewGateNo + 1) & " " & carnum
            FrmTcpServer.MvrSock.SendData (Trim(MVR_Str))
            Call DataLogger("[MVR UDP 전송]  DATA = " & MVR_Str)
        End If
    End If
    

Call LeaveCriticalSection(Glo_CS)

Exit Sub

Err_P:
    Call LeaveCriticalSection(Glo_CS)
    Call DataLogger(" [UDP Proc]  " & Err.Description)

End Sub


' 호스트 + 사전무인정산기 구조에서 실행됨
Public Sub PreAps_Proc(ByVal sCarno As String, ByVal iGateNo As Integer, ByVal sGateStat As String, ByVal sPassDate As String)

    Dim qry As String
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sPDate As String
    
On Error GoTo Err_P
    
    sPDate = Left(sPassDate, 19) 'yyyy-mm-dd hh:nn:ss
    
    If (Trim(sGateStat) = "출입제한") Then
        Exit Sub
    End If

    'If (iGateNo = 1 Or iGateNo = 3) Then
    If (InStr(sGateStat, "출차") > 0) Then
        If sCarno = "인식실패" Or sCarno = "미인식" Then
            'Call Sound_Out("미인식.WAV")
            Call GL_Emergency("[인실실패]", "인식실패", 0, 30, 10, 1, 2, 1, iGateNo)
            DataLogger ("인실실패")
            Exit Sub
        Else
            'Call Sound_Out("BELL.WAV")
        End If
        If Len(Trim(sCarno)) <= 6 Then
            Call GL_Emergency("[인실실패]", "인식실패", 0, 30, 10, 1, 2, 1, iGateNo)
            DataLogger ("차량번호 오류:" & Trim(sCarno))
            Exit Sub
        End If
        '계산결과 검색
        Set rs = New ADODB.Recordset
        qry = "SELECT * FROM tb_calculate WHERE (TICKET_NO = '" & Trim(sCarno) & "') AND (GREEN_NO='0') Order By REG_DATE Desc Limit 1"
        
        
        rs.Open qry, adoConn
        If Not (rs.EOF) Then
            '무인에서 정산했다
            'If Format(rs!REG_DATE, "YYYY-MM-DD HH:NN:SS") >= Format(DateAdd("s", -600, Now), "YYYY-MM-DD HH:NN:SS") Then
            'If Format(rs!REG_DATE, "YYYY-MM-DD HH:NN:SS") >= Format(DateAdd("s", -1 * (Glo_Grace_Time * 60), Now), "YYYY-MM-DD HH:NN:SS") Then
            If (DateDiff("n", Left(rs!PASS_DATE, 19), Format(Now, "YYYY-MM-DD HH:NN:SS")) <= Glo_Grace_Time) Then
            
                '무인에서 정산후 그레이스 타임 이내에 출차(정상)
                Call Relay_Out(0, 1)
                Call GL_Emergency("[감사합니다]", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                Set rs = Nothing
                adoConn.Execute "UPDATE tb_web_dc SET USE_YN = 'Y', DT_OUT='" & sPDate & "' WHERE DC_CARNO = '" & Trim(sCarno) & "' AND USE_YN = 'N'"
                adoConn.Execute "UPDATE tb_calculate SET GREEN_NO = '2' WHERE TICKET_NO = '" & Trim(sCarno) & "'"
                adoConn.Execute "Delete From tb_now Where CAR_NO= '" & Trim(sCarno) & "'"
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','사전정산(정상출차)',''," & 0 & ",'" & sPDate & "')"
                'Exit Sub
            Else
                '무인에서 정산후 그레이스 타임 이후 출차(재정산요구)
                Call GL_Emergency("[출차 지연]", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                Call Delay_Time(2)
                Call GL_Emergency("정산후 출차", "바랍니다.", 0, 30, 10, 1, 2, 1, iGateNo)
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','사전정산(출차시간초과)',''," & 0 & ",'" & sPDate & "')"
                DataLogger ("사전정산 후 출차시간초과 : " & Trim(sCarno))
            End If
        Else
            '정산결과가 없다
            '계산되지않은상황
            Set rs2 = New ADODB.Recordset
            qry = "SELECT * FROM tb_now WHERE CAR_NO = '" & Trim(sCarno) & "'"
            rs2.Open qry, adoConn
            If Not (rs2.EOF) Then
                If (DateDiff("n", Left(rs2!PASS_DATE, 19), Format(Now, "YYYY-MM-DD HH:NN:SS")) < Glo_Return_Time) Then
                    '회차차량이다.
                    Call Relay_Out(0, 1)
                    Call GL_Emergency("[감사합니다]", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                    
                    adoConn.Execute "UPDATE tb_web_dc SET USE_YN = 'Y', DT_OUT='" & sPDate & "' WHERE DC_CARNO = '" & Trim(sCarno) & "' AND USE_YN = 'N'"
                    'adoConn.Execute "UPDATE tb_calculate SET GREEN_NO = '2' WHERE TICKET_NO = '" & Trim(sCarNo) & "'"
                    adoConn.Execute "Delete From tb_now Where CAR_NO= '" & Trim(sCarno) & "'"
                    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','회차차량',''," & 0 & ",'" & sPDate & "')"
                    'Exit Sub
                Else

                    '웹할인을 받았는지 체크
                    'tb_now 의 CALC 필드를 이용하자
                    '최초 인서트될때 값이 0 이므로 웹할인을 하면 할인 프로그램에서 1로 세팅하자
                    If (rs2!CALC = "1") Then
                        '웹할인을 받았네???                                                        '10분 추가
                        'If (DateDiff("n", Left(rs2!PASS_DATE, 19), Format(Now, "YYYY-MM-DD HH:NN:SS")) < 60 + 10) Then
                        If (DateDiff("n", Left(rs2!PASS_DATE, 19), Format(Now, "YYYY-MM-DD HH:NN:SS")) <= Glo_Grace_Time) Then
                            '웹할인받고 1시간 이내에 출차(무료)
                            Call Relay_Out(0, 1)
                            Call GL_Emergency("[감사합니다]", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                            Set rs2 = Nothing
                            adoConn.Execute "UPDATE tb_web_dc SET USE_YN = 'Y', DT_OUT='" & sPDate & "' WHERE DC_CARNO = '" & Trim(sCarno) & "' AND USE_YN = 'N'"
                            'adoConn.Execute "UPDATE tb_calculate SET GREEN_NO = '2' WHERE TICKET_NO = '" & Trim(sCarNo) & "'"
                            adoConn.Execute "Delete From tb_now Where CAR_NO= '" & Trim(sCarno) & "'"
                            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','사전미정산 : 웹할인차량(무료)',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                            Exit Sub
                        Else
                            '웹할인받고 그레이스타임 이후에 출차(재정산요구)
                            Call GL_Emergency("[출차 지연]", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                            Call Delay_Time(2)
                            Call GL_Emergency("정산후 출차", "바랍니다.", 0, 30, 10, 1, 2, 1, iGateNo)
                            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','사전미정산 : 웹할인차량(제한시간초과)',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                            DataLogger ("웹할인후 출차지연 재정산요구 : " & Trim(sCarno))
                        End If
                    Else
                        '무인에서 정산도 안하고 웹할인도 안받았네???
                        Call GL_Emergency("[미정산차량]", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','사전미정산 : 웹할인(안함)',''," & 0 & ",'" & sPDate & "')"
                        DataLogger ("미정산차량 : " & Trim(sCarno))
                    End If
                End If
            Else
                '하물며 입차기록이없다?
                'Call Relay_Out(0, 1)
                Call GL_Emergency("입차정보없음", Trim(sCarno), 0, 30, 10, 1, 2, 1, iGateNo)
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(sCarno) & "', 'HOST','사전미정산 : 입차정보없음',''," & 0 & ",'" & sPDate & "')"
                DataLogger ("사전미정산 : 입차정보없음 : " & Trim(sCarno))
            End If
        End If
        Set rs = Nothing
        Set rs2 = Nothing
    End If
    
    
    
    Exit Sub
Err_P:
    Call DataLogger(" [PreAPS Proc]  " & Err.Description)
End Sub

Public Sub APS_Connect()
Dim bData() As Byte

On Error GoTo Err_P

With FrmTcpServer
    Call DataLogger("[APS_Connect] APS 접속시도 : " & Glo_Aps_IP & " " & 5889)
    If (.ApsS_sock.State <> sckClosed) Then
        .ApsS_sock.Close
    End If
    .ApsS_sock.Connect Glo_Aps_IP, 5889
End With

Exit Sub

Err_P:
    Call DataLogger("[APS_Connect] 에러내용 : " & Err.Description)

End Sub



Public Sub Jung_Show(ByVal carnum As String, ByVal sGateNo As String, ByVal sPassDate As String)
    Dim qry As String
    'Dim rs As ADODB.Recordset
    Dim Tmp_Path As String
    Dim itmX As ListItem
    Dim gateNo As Integer
    Dim inout As String
    Dim Gubun As String
    Dim i, s As Integer
    Dim ECHO As ICMP_ECHO_REPLY
    Dim bQryResult As Boolean
    Dim sGateName As String

On Error GoTo Err_P


    carnum = Trim(carnum)

    'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & carnum & "' And(PASS_DATE >= '" & Format(Now, "yyyy-mm-dd") & " " & "00:00:00" & "' AND PASS_DATE <= '" & Format(Now, "yyyy-mm-dd") & " " & "23:59:59" & "') Order By PASS_DATE Desc LIMIT 1"
    'Qry = "Select * From tb_inout Where CAR_NO = '" & carnum & "' AND PASS_DATE = '" & sPassDate & "' "
    
    '메인호스트에서 한글필터링 과정을 거치게되면 처음 차번인식번호가 달라짐.
    'tb_inout 검색 필드명을 CAR_NO 대신 REC_NO로 해야함
    qry = "Select * From tb_inout Where REC_NO = '" & carnum & "' AND PASS_DATE = '" & sPassDate & "' "
    
    Set rs = New ADODB.Recordset
    'rs.Open Qry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
        Call DataLogger("[JungShow]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    
    With Jung

    If Not (rs.EOF) Then
'        If (rs!PASS_INOUT = "IN") Then
        If (sGateNo = 0) Then
        
'            .lbl_title_in(0).Caption = "게이트 : "
'
'            If (Glo_User_Type = "구분1/구분2") Then
'                '.lbl_title_in(1).Caption = "이  름 : "
'                '.lbl_title_in(2).Caption = "연락처 : "
'                .lbl_title_in(1).Caption = "소  속 : "
'                .lbl_title_in(2).Caption = "직  급 : "
'                .lbl_info_in(1).Caption = "" & rs!DRIVER_NAME
'                .lbl_info_in(2).Caption = "" & rs!DRIVER_PHONE
'            Else
'                .lbl_title_in(1).Caption = "   동   : "
'                .lbl_title_in(2).Caption = " 호  수 : "
'                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
'                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
'            End If
'
'            .lbl_title_in(3).Caption = "인식번호 : "
'            .lbl_title_in(4).Caption = "종 료 일 : "
'            .lbl_title_in(5).Caption = "입출상태 : "
            .lbl_title_in(0).Caption = "   게이트 : "

            If (Glo_User_Type = "구분1/구분2") Then
                .lbl_title_in(1).Caption = "이        름 : "
                .lbl_title_in(2).Caption = "연  락  처 : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_NAME
                .lbl_info_in(2).Caption = "" & rs!DRIVER_PHONE
            Else
                .lbl_title_in(1).Caption = "       동     : "
                .lbl_title_in(2).Caption = "    호 수   : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
            End If

            .lbl_title_in(3).Caption = "인식번호 : "
            .lbl_title_in(4).Caption = "종  료  일 : "
            .lbl_title_in(5).Caption = "입출 상태 : "
            
            '.lbl_info_in(0).Caption = "" & rs!PASS_GATE
            If (rs!PASS_GATE = "0") Then
                .lbl_info_in(0).Caption = LANE1_Name
            ElseIf (rs!PASS_GATE = "1") Then
                .lbl_info_in(0).Caption = LANE2_Name
            ElseIf (rs!PASS_GATE = "2") Then
                .lbl_info_in(0).Caption = LANE3_Name
            ElseIf (rs!PASS_GATE = "3") Then
                .lbl_info_in(0).Caption = LANE4_Name
            ElseIf (rs!PASS_GATE = "4") Then
                .lbl_info_in(0).Caption = LANE5_Name
            ElseIf (rs!PASS_GATE = "5") Then
                .lbl_info_in(0).Caption = LANE6_Name
            End If
            .lbl_info_in(3).Caption = "" & rs!REC_NO
            .lbl_info_in(4).Caption = "" & rs!END_DATE
            '.lbl_info_in(5).Caption = "" & rs!PASS_RESULT
            .lbl_info_in(5).Caption = Get_InOutStrint(rs!PASS_RESULT)
            
'            Select Case Trim(rs!PASS_RESULT)
'                Case "정상입차"
'                    .Proc_Type(0).Caption = " " & "정기권입차"
'                    .Proc_Type(0).ForeColor = vbBlue
'                Case "정상출차"
'                    .Proc_Type(0).Caption = " " & "정기권출차"
'                    .Proc_Type(0).ForeColor = vbBlue
'                Case Else
'                    .Proc_Type(0).Caption = " " & rs!PASS_RESULT
'                    .Proc_Type(0).ForeColor = vbRed
'            End Select

            '.Proc_Type(0).Caption = "" & rs!PASS_RESULT
            .Proc_Type(0).Caption = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
            If rs!Pass_YN = "Y" Then
                .Proc_Type(0).ForeColor = vbBlue
            Else
                .Proc_Type(0).ForeColor = vbRed
            End If
            '==================================================================================================
            'Call Ping(rs!PASS_IP, ECHO)
            'If Left$(ECHO.Data, 1) <> Chr$(0) Then
'                Tmp_Path = Dir(rs!PASS_IMAGE)
'                If (Tmp_Path <> "") Then
                If (IsFile(rs!pass_image) = True) Then
                    .ImageIn(0).Picture = LoadPicture(rs!pass_image)
                Else
                    .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            'Else
            '    .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
            '    Call DataLogger(" [Jung Show]    " & "Ping Test Failure...!!")
            '    Call DataLogger(" [Jung Show]    " & CarNum & "  " & Tmp_Path)
            'End If
            '==================================================================================================
            .lbl_carno(0).Caption = rs!CAR_NO
            .lbl_time_now(0).Caption = Mid(rs!PASS_DATE, 1, Len(rs!PASS_DATE) - 4)
        Else
            .lbl_title_Out(0).Caption = "   게이트 : "

            If (Glo_User_Type = "구분1/구분2") Then
                .lbl_title_Out(1).Caption = "이        름 : "
                .lbl_title_Out(2).Caption = "연  락  처 : "
                .lbl_info_Out(1).Caption = "" & rs!DRIVER_NAME
                .lbl_info_Out(2).Caption = "" & rs!DRIVER_PHONE
            Else
                .lbl_title_Out(1).Caption = "       동     : "
                .lbl_title_Out(2).Caption = "    호 수   : "
                .lbl_info_Out(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_Out(2).Caption = "" & rs!DRIVER_CLASS
            End If

            .lbl_title_Out(3).Caption = "인식번호 : "
            .lbl_title_Out(4).Caption = "종  료  일 : "
            .lbl_title_Out(5).Caption = "입출 상태 : "
            '.lbl_info_out(0).Caption = "" & rs!PASS_GATE
            If (rs!PASS_GATE = "0") Then
                .lbl_info_Out(0).Caption = LANE1_Name
            ElseIf (rs!PASS_GATE = "1") Then
                .lbl_info_Out(0).Caption = LANE2_Name
            ElseIf (rs!PASS_GATE = "2") Then
                .lbl_info_Out(0).Caption = LANE3_Name
            ElseIf (rs!PASS_GATE = "3") Then
                .lbl_info_Out(0).Caption = LANE4_Name
            ElseIf (rs!PASS_GATE = "4") Then
                .lbl_info_Out(0).Caption = LANE5_Name
            ElseIf (rs!PASS_GATE = "5") Then
                .lbl_info_Out(0).Caption = LANE6_Name
            End If
            
            .lbl_info_Out(3).Caption = "" & rs!REC_NO
            .lbl_info_Out(4).Caption = "" & rs!END_DATE
            '.lbl_info_Out(5).Caption = "" & rs!PASS_RESULT
            .lbl_info_Out(5).Caption = "" & Get_InOutStrint(rs!PASS_RESULT)
'            Select Case Trim(rs!PASS_RESULT)
'                Case "정상입차"
'                    .Proc_Type(1).Caption = " " & "정기권입차"
'                    .Proc_Type(1).ForeColor = vbBlue
'                Case "정상출차"
'                    .Proc_Type(1).Caption = " " & "정기권출차"
'                    .Proc_Type(1).ForeColor = vbBlue
'                Case Else
'                    .Proc_Type(1).Caption = " " & rs!PASS_RESULT
'                    .Proc_Type(1).ForeColor = vbRed
'            End Select
            '.Proc_Type(1).Caption = "" & rs!PASS_RESULT
            .Proc_Type(1).Caption = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
            If rs!Pass_YN = "Y" Then
                .Proc_Type(1).ForeColor = vbBlue
            Else
                .Proc_Type(1).ForeColor = vbRed
            End If
            
            '==================================================================================================
            'Call Ping(rs!PASS_IP, ECHO)
            'If Left$(ECHO.Data, 1) <> Chr$(0) Then
                'Tmp_Path = Dir(rs!PASS_IMAGE)
                'If (Tmp_Path <> "") Then
                If (IsFile(rs!pass_image) = True) Then
                    .ImageIn(1).Picture = LoadPicture(rs!pass_image)
                Else
                    .ImageIn(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            'Else
            '    .ImageIn(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
            '    Call DataLogger(" [Jung Show]    " & "Ping Test Failure...!!")
            '    Call DataLogger(" [Jung Show]    " & CarNum & "  " & Tmp_Path)
            'End If
            '==================================================================================================
            .lbl_carno(1).Caption = rs!CAR_NO
            .lbl_time_now(1).Caption = Mid(rs!PASS_DATE, 1, Len(rs!PASS_DATE) - 4)
        End If
        Set itmX = .ListView2.ListItems.Add(, , "" & Left(rs!PASS_DATE, 19))
        itmX.SubItems(1) = "" & rs!CAR_NO
        'itmX.SubItems(2) = "" & rs!PASS_GATE
        If (rs!PASS_GATE = "0") Then
            sGateName = LANE1_Name
        ElseIf (rs!PASS_GATE = "1") Then
            sGateName = LANE2_Name
        ElseIf (rs!PASS_GATE = "2") Then
            sGateName = LANE3_Name
        ElseIf (rs!PASS_GATE = "3") Then
            sGateName = LANE4_Name
        ElseIf (rs!PASS_GATE = "4") Then
            sGateName = LANE5_Name
        ElseIf (rs!PASS_GATE = "5") Then
            sGateName = LANE6_Name
        End If
        itmX.SubItems(2) = sGateName

        itmX.SubItems(3) = "" & rs!DRIVER_NAME
        itmX.SubItems(4) = "" & rs!DRIVER_PHONE
        itmX.SubItems(5) = "" & rs!REC_NO
        itmX.SubItems(6) = "" & rs!END_DATE
        'itmX.SubItems(7) = "" & rs!PASS_RESULT
        itmX.SubItems(7) = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
'        itmX.SubItems(7) = "" & rs!PASS_DATE
        'itmX.SubItems(8) = "" & rs!PASS_INOUT
        itmX.SubItems(8) = "" & rs!pass_image
        .ListView2.Sorted = True
        
        '.ListView2.Sorted = False
        '.ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
        .ListView2.ListItems(1).Selected = True
        .ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible
    End If
    
    
    
    
Set rs = Nothing


'    If (Glo_ParkFull_YN = "Y") Then
'        .lbl_ParkFull.Caption = "만차현황 : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
'    End If

End With

Exit Sub

Err_P:
    Call DataLogger("[Jung Show] " & Err.Description)
End Sub



Public Sub G4Mini_4INShow(ByVal Data As String, ByVal sGateNo As String, ByVal sPassDate As String)
Dim i As Integer
Dim gateNo As Integer
Dim GateName As String
Dim carno As String
Dim rs As Recordset
Dim qry As String
Dim Tmp_File As String
Dim itmX
Dim bQryResult  As Boolean

With FrmG4Mini
'        GateNo = Left(Data, 1)
'        i = LenH(Data)
'        CarNo = Mid(Data, 3, (i - 2))
        gateNo = sGateNo
        carno = Data

        'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & CarNo & "' Order By PASS_DATE Desc Limit 1"
        'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & CarNo & "' And PASS_DATE = '" & sPassDate & "' "
        'Qry = "Select * From tb_inout Where CAR_NO = '" & CarNo & "' And PASS_DATE = '" & sPassDate & "' LIMIT 1"
        
        '메인호스트에서 한글필터링 과정을 거치게되면 처음 차번인식번호가 달라짐.
        'tb_inout 검색 필드명을 CAR_NO 대신 REC_NO로 해야함
        qry = "Select * From tb_inout Where REC_NO = '" & carno & "' And PASS_DATE = '" & sPassDate & "' LIMIT 1"

        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
            Call DataLogger("[G4Mini]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If

        If Not (rs.EOF) Then
                .lbl_carno(gateNo).Caption = "" & rs!CAR_NO
'                Tmp_File = Dir(rs!PASS_IMAGE)
'                If (Tmp_File <> "") Then

                If (IsFile(rs!pass_image) = True) Then
                    .ImageIn(gateNo).Picture = LoadPicture(rs!pass_image)
                End If

                For i = 0 To 3
                    .Shp_Rec(i).Visible = False
                Next i
                .Shp_Rec(gateNo).Visible = True
                .lbl_time_now(gateNo).Caption = "" & Left(rs!PASS_DATE, 19)
                '.lbl_RecState(gateNo).Caption = "" & rs!PASS_RESULT
                .lbl_RecState(gateNo).Caption = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
                If rs!Pass_YN = "Y" Then
                    .lbl_RecState(gateNo).ForeColor = vbBlue
                Else
                    .lbl_RecState(gateNo).ForeColor = vbRed
                End If
                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " GateNo : " & gateNo & ", 차량번호 : " & rs!CAR_NO & ", 처리결과 : " & rs!PASS_RESULT, 0
            
            
                Set itmX = .ListView2.ListItems.Add(, , "" & Left(rs!PASS_DATE, 19))
                itmX.SubItems(1) = "" & rs!CAR_NO
                itmX.SubItems(2) = "" & rs!CAR_GUBUN
                itmX.SubItems(3) = "" & rs!DRIVER_NAME
                itmX.SubItems(4) = "" & rs!DRIVER_PHONE
                itmX.SubItems(5) = "" & rs!START_DATE
                itmX.SubItems(6) = "" & rs!END_DATE
                'itmX.SubItems(7) = "" & rs!PASS_RESULT
                itmX.SubItems(7) = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
                'itmX.SubItems(7) = "" & rs!PASS_DATE
                itmX.SubItems(8) = "" & rs!pass_image
                '.ListView2.Sorted = False
                '.ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
                .ListView2.ListItems(1).Selected = True
                .ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible

        Else
            'Beep
        End If
        Set rs = Nothing

'        If (Glo_ParkFull_YN = "Y") Then
'            .lbl_ParkFull.Caption = "만차현황 : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
'        End If

End With

On Error Resume Next

End Sub


Public Sub G1_Show(ByVal Data As String, ByVal sGateNo As String, ByVal sPassDate As String)
    Dim i As Integer
    Dim gateNo As Integer
    Dim GateName As String
    Dim carno As String
    Dim rs As Recordset
    Dim qry As String
    Dim Tmp_File As String
    Dim itmX
    Dim bQryResult As Boolean
    Dim sGateName As String
    Dim sResult As String
    
With FrmG1
'        GateNo = Left(Data, 1)
'        i = LenH(Data)
'        CarNo = Mid(Data, 3, (i - 2))
        gateNo = sGateNo
        carno = Data

        'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & CarNo & "' Order By PASS_DATE Desc Limit 1"
        'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & CarNo & "' AND PASS_DATE = '" & sPassDate & "' "
        
        '메인호스트에서 한글필터링 과정을 거치게되면 처음 차번인식번호가 달라짐.
        'tb_inout 검색 필드명을 CAR_NO 대신 REC_NO 로 해야함
        qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And REC_NO = '" & carno & "' AND PASS_DATE = '" & sPassDate & "' "
        
        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
            Call DataLogger("[G1Show]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If


        If Not (rs.EOF) Then
                .lbl_carno(gateNo).Caption = "" & rs!CAR_NO
                If (IsFile(rs!pass_image) = True) Then
                    .ImageIn(gateNo).Picture = LoadPicture(rs!pass_image)
                End If


'            .lbl_title_in(0).Caption = "GATE : "
'
'            If (Glo_User_Type = "구분1/구분2") Then
'                .lbl_title_in(1).Caption = "이  름 : "
'                .lbl_title_in(2).Caption = "연락처 : "
'                .lbl_info_in(1).Caption = "" & rs!DRIVER_NAME
'                .lbl_info_in(2).Caption = "" & rs!DRIVER_PHONE
'            Else
'                .lbl_title_in(1).Caption = "동    : "
'                .lbl_title_in(2).Caption = "호    : "
'                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
'                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
'            End If
            .lbl_title_in(0).Caption = "이  름 : "
            .lbl_info_in(0).Caption = "" & rs!DRIVER_NAME
            
            If (Glo_User_Type = "구분1/구분2") Then
                .lbl_title_in(1).Caption = "소  속 : "
                .lbl_title_in(2).Caption = "직  급 : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
            Else
                .lbl_title_in(1).Caption = "  동  : "
                .lbl_title_in(2).Caption = "  호  : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
            End If
            
           
            .lbl_title_in(3).Caption = "인식번호 : "
            .lbl_title_in(4).Caption = "종료일 : "
            .lbl_title_in(5).Caption = "입출상태 : "
'            .lbl_info_in(0).Caption = "" & rs!PASS_GATE
            .lbl_info_in(3).Caption = "" & rs!REC_NO
            .lbl_info_in(4).Caption = "" & rs!END_DATE
            '.lbl_info_in(5).Caption = "" & rs!pass_result
            .lbl_info_in(5).Caption = Get_InOutStrint(rs!PASS_RESULT)

            
            
'            Select Case Trim(rs!PASS_RESULT)
'                Case "정상입차"
'                    .Proc_Type(0).Caption = " " & "정기권입차"
'                    .Proc_Type(0).ForeColor = vbBlue '성훈
'                Case "정상출차"
'                    .Proc_Type(0).Caption = " " & "정기권출차"
'                    .Proc_Type(0).ForeColor = vbBlue
'                Case Else
'                    .Proc_Type(0).Caption = " " & rs!PASS_RESULT
'                    .Proc_Type(0).ForeColor = vbRed '성훈
'            End Select
            '.Proc_Type(0).Caption = "" & rs!PASS_RESULT
            
            .Proc_Type(0).Caption = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
             If rs!Pass_YN = "Y" Then
                .Proc_Type(0).ForeColor = vbBlue
            Else
                .Proc_Type(0).ForeColor = vbRed
            End If
            '==================================================================================================
            'Call Ping(rs!PASS_IP, ECHO)
            'If Left$(ECHO.Data, 1) <> Chr$(0) Then
'                Tmp_File = Dir(rs!PASS_IMAGE)
'                If (Tmp_File <> "") Then

                If (IsFile(rs!pass_image) = True) Then
                    .ImageIn(0).Picture = LoadPicture(rs!pass_image)
                Else
                    .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            
            
                If (.Shape2.BorderColor = &HFFFF00) Then
                    .Shape2.BorderColor = &H80FF&
                Else
                    .Shape2.BorderColor = &HFFFF00
                End If
                .Shape2.Refresh

            '==================================================================================================
            .lbl_carno(0).Caption = rs!CAR_NO
            .lbl_time_now(0).Caption = Left(rs!PASS_DATE, 19)
        
        
            Set itmX = .ListView2.ListItems.Add(, , "" & .lbl_time_now(0).Caption)
            itmX.SubItems(1) = "" & rs!CAR_NO
            'itmX.SubItems(2) = "" & rs!PASS_GATE
            If (rs!PASS_GATE = "0") Then
                sGateName = LANE1_Name
            ElseIf (rs!PASS_GATE = "1") Then
                sGateName = LANE2_Name
            ElseIf (rs!PASS_GATE = "2") Then
                sGateName = LANE3_Name
            ElseIf (rs!PASS_GATE = "3") Then
                sGateName = LANE4_Name
            ElseIf (rs!PASS_GATE = "4") Then
                sGateName = LANE5_Name
            ElseIf (rs!PASS_GATE = "5") Then
                sGateName = LANE6_Name
            End If
            itmX.SubItems(2) = sGateName
        
            itmX.SubItems(3) = "" & rs!DRIVER_NAME
            itmX.SubItems(4) = "" & rs!DRIVER_PHONE
            itmX.SubItems(5) = "" & rs!REC_NO
            itmX.SubItems(6) = "" & rs!END_DATE
            'itmX.SubItems(7) = "" & rs!PASS_RESULT
            itmX.SubItems(7) = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
            'itmX.SubItems(8) = "" & rs!PASS_INOUT
            itmX.SubItems(8) = "" & rs!pass_image
            .ListView2.Sorted = True
            
            '.ListView2.Sorted = False
            '.ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
            .ListView2.ListItems(1).Selected = True
            .ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible
            
        Else
            'Beep
        End If
        Set rs = Nothing

        
'        If (Glo_ParkFull_YN = "Y") Then
'            .lbl_ParkFull.Caption = "만차현황 : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
'        End If
        
End With

On Error Resume Next

End Sub


Public Sub G6_23Show(ByVal Data As String, ByVal sGateNo As String, ByVal sPassDate As String)
Dim i As Integer
Dim gateNo As Integer
Dim GateName As String
Dim carno As String
Dim rs As Recordset
Dim qry As String
Dim Tmp_File As String
Dim itmX
Dim bQryResult As Boolean




With FrmG6_23
'        GateNo = Left(Data, 1)
'        i = LenH(Data)
'        CarNo = Mid(Data, 3, (i - 2))
        gateNo = sGateNo
        carno = Data
         
        'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & CarNo & "' Order By PASS_DATE Desc Limit 1"
        'Qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And CAR_NO = '" & CarNo & "' AND PASS_DATE = '" & sPassDate & "' "
        
        '메인호스트에서 한글필터링 과정을 거치게되면 처음 차번인식번호가 달라짐.
        'tb_inout 검색 필드명을 CAR_NO 대신 REC_NO로 해야함
        qry = "Select * From tb_inout Where PASS_GATE = '" & sGateNo & "' And REC_NO = '" & carno & "' AND PASS_DATE = '" & sPassDate & "' "
        
        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
            Call DataLogger("[G6_23Show]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        If Not (rs.EOF) Then
                .lbl_carno(gateNo).Caption = "" & rs!CAR_NO
'                Tmp_File = Dir(rs!PASS_IMAGE)
'                If (Tmp_File <> "") Then
                If (IsFile(rs!pass_image) = True) Then
                    .ImageIn(gateNo).Picture = LoadPicture(rs!pass_image)
                End If
                For i = 0 To 5
                    .Shp_Rec(i).Visible = False
                Next i
                .Shp_Rec(gateNo).Visible = True
                .lbl_time_now(gateNo).Caption = "" & Left(rs!PASS_DATE, 19)
                '''.lbl_RecState(gateNo).Caption = "" & rs!PASS_RESULT
                .lbl_RecState(gateNo).Caption = "" & Get_ResultString(rs!PASS_RESULT, rs!PASS_GATE)
                If rs!Pass_YN = "Y" Then
                    .lbl_RecState(gateNo).ForeColor = vbBlue
                Else
                    .lbl_RecState(gateNo).ForeColor = vbRed
                End If
'                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " GateNo : " & GateNo & ", 차량번호 : " & rs!CAR_NO & ", 처리결과 : " & rs!PASS_RESULT, 0
            
            
'                Set itmX = .ListView2.ListItems.Add(, , "" & rs!PASS_DATE)
'                itmX.SubItems(1) = "" & rs!CAR_NO
'                itmX.SubItems(2) = "" & rs!CAR_GUBUN
'                itmX.SubItems(3) = "" & rs!DRIVER_NAME
'                itmX.SubItems(4) = "" & rs!DRIVER_PHONE
'                itmX.SubItems(5) = "" & rs!Start_Date
'                itmX.SubItems(6) = "" & rs!End_Date
'                itmX.SubItems(7) = "" & rs!PASS_RESULT
'                'itmX.SubItems(7) = "" & rs!PASS_DATE
'                itmX.SubItems(8) = "" & rs!PASS_IMAGE
'                '.ListView2.Sorted = False
'                '.ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
'                '.ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible
        
        Else
            'Beep
        End If
        Set rs = Nothing
        
        
        
'        If (Glo_ParkFull_YN = "Y") Then
'            .lbl_ParkFull.Caption = "만차현황 : " & Glo_ParkNow_Count & " / " & Glo_ParkFull_Count
'        End If
        


End With

On Error Resume Next

End Sub





Public Sub G6_23_Freepass(ByVal sFreePass As String, ByVal sGateNo As String, ByVal sYN As String)
    Dim iValue As Integer
    
    If sYN = "Y" Then
        iValue = 1
    Else
        iValue = 0
    End If
    
    
    With FrmG6_23
    If (sFreePass = "FREEPASS") Then ' 일반차량 프리패스
            
            Select Case sGateNo
            Case 0
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane1_YN = "Y"
                    Else
                        Glo_FreePassLane1_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane1_YN", Glo_FreePassLane1_YN)
            Case 1
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane2_YN = "Y"
                    Else
                        Glo_FreePassLane2_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane2_YN", Glo_FreePassLane2_YN)
            Case 2
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane3_YN = "Y"
                    Else
                        Glo_FreePassLane3_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane3_YN", Glo_FreePassLane3_YN)
            Case 3
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane4_YN = "Y"
                    Else
                        Glo_FreePassLane4_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane4_YN", Glo_FreePassLane4_YN)
            Case 4
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane5_YN = "Y"
                    Else
                        Glo_FreePassLane5_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane5_YN", Glo_FreePassLane5_YN)
            Case 5
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane6_YN = "Y"
                    Else
                        Glo_FreePassLane6_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane6_YN", Glo_FreePassLane6_YN)
            End Select
            
    
    ElseIf (sFreePass = "TAXI") Then ' 영업용차량 프리패스
            
            Select Case sGateNo
                Case 0
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI1_YN = "Y"
                    Else
                        Glo_TAXI1_YN = "N"
                    End If
        
                    Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
                
                Case 1
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI2_YN = "Y"
                    Else
                        Glo_TAXI2_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI2_YN", Glo_TAXI2_YN)
                Case 2
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI3_YN = "Y"
                    Else
                        Glo_TAXI3_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI3_YN", Glo_TAXI3_YN)
                Case 3
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI4_YN = "Y"
                    Else
                        Glo_TAXI4_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI4_YN", Glo_TAXI4_YN)
                Case 4
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI5_YN = "Y"
                    Else
                        Glo_TAXI5_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI5_YN", Glo_TAXI5_YN)
                Case 5
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI6_YN = "Y"
                    Else
                        Glo_TAXI6_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI6_YN", Glo_TAXI6_YN)
            End Select
            
    ElseIf (sFreePass = "NOWORK") Then ' 자리비움
    
            Dim sNoWork As String
            If iValue = 0 Then
                sNoWork = "근무중"
            ElseIf iValue = 1 Then
                sNoWork = "자리비움"
            End If
            
            Select Case sGateNo
                Case 0
                    '.chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        .NoWork(sGateNo).Caption = sNoWork
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        .NoWork(sGateNo).Caption = sNoWork
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 1
                    '.chk_NoWork(sGateNo).value = iValue
                    .NoWork(sGateNo).Caption = sNoWork
                    If (iValue = 1) Then
                        Glo_Lane2_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane2_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 2
                    '.chk_NoWork(sGateNo).value = iValue
                    .NoWork(sGateNo).Caption = sNoWork
                    If (iValue = 1) Then
                        Glo_Lane3_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane3_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 3
                    '.chk_NoWork(sGateNo).value = iValue
                    .NoWork(sGateNo).Caption = sNoWork
                    If (iValue = 1) Then
                        Glo_Lane4_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane4_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 4
                    '.chk_NoWork(sGateNo).value = iValue
                    .NoWork(sGateNo).Caption = sNoWork
                    If (iValue = 1) Then
                        Glo_Lane5_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane5_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 5
                    '.chk_NoWork(sGateNo).value = iValue
                    .NoWork(sGateNo).Caption = sNoWork
                    If (iValue = 1) Then
                        Glo_Lane6_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane6_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                    'Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
            End Select
            
            Dim sLog As String
            sLog = "Lane" & sGateNo + 1 & ":" & sNoWork
            Call DataLogger(sLog)
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('자리비움제어', 'Remote','" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            
    End If
    End With
End Sub
Public Sub G4Mini_4IN_Freepass(ByVal sFreePass As String, ByVal sGateNo As String, ByVal sYN As String)
    Dim iValue As Integer
    
    If sYN = "Y" Then
        iValue = 1
    Else
        iValue = 0
    End If
    
    
    With FrmG4Mini
    If (sFreePass = "FREEPASS") Then ' 일반차량 프리패스
            
            Select Case sGateNo
            Case 0
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane1_YN = "Y"
                    Else
                        Glo_FreePassLane1_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane1_YN", Glo_FreePassLane1_YN)
            Case 1
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane2_YN = "Y"
                    Else
                        Glo_FreePassLane2_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane2_YN", Glo_FreePassLane2_YN)
            Case 2
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane3_YN = "Y"
                    Else
                        Glo_FreePassLane3_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane3_YN", Glo_FreePassLane3_YN)
            Case 3
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane4_YN = "Y"
                    Else
                        Glo_FreePassLane4_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane4_YN", Glo_FreePassLane4_YN)
            End Select
            
    
    ElseIf (sFreePass = "TAXI") Then ' 영업용차량 프리패스
            
            Select Case sGateNo
                Case 0
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI1_YN = "Y"
                    Else
                        Glo_TAXI1_YN = "N"
                    End If
        
                    Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
                
                Case 1
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI2_YN = "Y"
                    Else
                        Glo_TAXI2_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI2_YN", Glo_TAXI2_YN)
                Case 2
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI3_YN = "Y"
                    Else
                        Glo_TAXI3_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI3_YN", Glo_TAXI3_YN)
                Case 3
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI4_YN = "Y"
                    Else
                        Glo_TAXI4_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI4_YN", Glo_TAXI4_YN)
            End Select
            
    ElseIf (sFreePass = "NOWORK") Then ' 자리비움
    
            Dim sNoWork As String
            If iValue = 0 Then
                sNoWork = "근무중"
            ElseIf iValue = 1 Then
                sNoWork = "자리비움"
            End If
            
            Select Case sGateNo
                Case 0
                    .chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 1
                    .chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_Lane2_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane2_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 2
                    .chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_Lane3_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane3_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 3
                    .chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_Lane4_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane4_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                    'Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
            End Select
            
            Dim sLog As String
            sLog = "Lane" & sGateNo + 1 & ":" & sNoWork
            Call DataLogger(sLog)
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('자리비움제어', 'Remote','" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            
    End If
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''.
' Jung_Freepass : 일반차량/영업차량, 게이트번호, 프리패스ON/OFF
Public Sub Jung_Freepass(ByVal sFreePass As String, ByVal sGateNo As String, ByVal sYN As String)
    Dim iValue As Integer
    
    If sYN = "Y" Then
        iValue = 1
    Else
        iValue = 0
    End If
    
    
    With Jung
    If (sFreePass = "FREEPASS") Then ' 일반차량 프리패스
            
            Select Case sGateNo
            Case 0
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane1_YN = "Y"
                    Else
                        Glo_FreePassLane1_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane1_YN", Glo_FreePassLane1_YN)
            Case 1
                    .Chk_FreePass(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_FreePassLane2_YN = "Y"
                    Else
                        Glo_FreePassLane2_YN = "N"
                    End If
                    
                    Call Put_Ini("System Config", "FreePassLane2_YN", Glo_FreePassLane2_YN)
            End Select
            
    
    ElseIf (sFreePass = "TAXI") Then ' 영업용차량 프리패스
            
            Select Case sGateNo
                Case 0
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI1_YN = "Y"
                    Else
                        Glo_TAXI1_YN = "N"
                    End If
        
                    Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
                
                Case 1
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI2_YN = "Y"
                    Else
                        Glo_TAXI2_YN = "N"
                    End If
                    Call Put_Ini("System Config", "TAXI2_YN", Glo_TAXI2_YN)
            End Select
            
    ElseIf (sFreePass = "NOWORK") Then ' 자리비움
    
            Dim sNoWork As String
            If iValue = 0 Then
                sNoWork = "근무중"
            ElseIf iValue = 1 Then
                sNoWork = "자리비움"
            End If
            
            Select Case sGateNo
                Case 0
                    .chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                Case 1
                    .chk_NoWork(sGateNo).value = iValue
                    If (iValue = 1) Then
                        Glo_Lane2_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane2_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                    'Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
            End Select
            
            Dim sLog As String
            sLog = "Lane" & sGateNo + 1 & ":" & sNoWork
            Call DataLogger(sLog)
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('자리비움제어', 'Remote','" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            
    End If
    End With
End Sub
Public Sub G1_Freepass(ByVal sFreePass As String, ByVal sGateNo As String, ByVal sYN As String)
    Dim iValue As Integer
    
    If sYN = "Y" Then
        iValue = 1
    Else
        iValue = 0
    End If
    
    With FrmG1
    If (sFreePass = "FREEPASS") Then ' 일반차량 프리패스
            
            Select Case sGateNo
                Case 0
                        .Chk_FreePass(sGateNo).value = iValue
                        If (iValue = 1) Then
                            Glo_FreePassLane1_YN = "Y"
                        Else
                            Glo_FreePassLane1_YN = "N"
                        End If
                        
                        Call Put_Ini("System Config", "FreePassLane1_YN", Glo_FreePassLane1_YN)
                End Select


    ElseIf (sFreePass = "TAXI") Then ' 영업용차량 프리패스
            
            Select Case sGateNo
                Case 0
                    .chk_Taxi(sGateNo).value = iValue
                     
                    If (iValue = 1) Then
                        Glo_TAXI1_YN = "Y"
                    Else
                        Glo_TAXI1_YN = "N"
                    End If
        
                    Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
                End Select
                
                
    ElseIf (sFreePass = "NOWORK") Then ' 자리비움
            Dim sNoWork As String
            If iValue = 0 Then
                sNoWork = "근무중"
            ElseIf iValue = 1 Then
                sNoWork = "자리비움"
            End If
            Select Case sGateNo
                Case 0
                    .chk_NoWork(sGateNo).value = iValue
                    
                    If (iValue = 1) Then
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = False
                        .Chk_FreePass(sGateNo).Enabled = False
                    Else
                        Glo_Lane1_NoWork = sNoWork
                        .chk_Taxi(sGateNo).Enabled = True
                        .Chk_FreePass(sGateNo).Enabled = True
                    End If
                    'Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
            End Select
            
            Dim sLog As String
            sLog = "Lane" & sGateNo + 1 & ":" & sNoWork
            Call DataLogger(sLog)
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('자리비움제어', 'Remote','" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    End If
    End With
End Sub





Private Sub sOutput(strIP As String, strText As String)
    If (FrmTcpServer.Check2.value = 1) Then
        FrmTcpServer.ListData.AddItem Format(Now, "YYYY-MM-DD HH:NN:SS") & "    " & strIP & "     " & strText, 0
    End If
End Sub

Public Function Get_InOutStrint(ByVal sRes As String)
    If (InStr(1, sRes, "입차")) Then
        Get_InOutStrint = "입차"
    Else
        Get_InOutStrint = "출차"
    End If
End Function

Public Function Get_ResultString(ByVal sRes As String, ByVal sGATE As String)
    If (InStr(1, sRes, "방문예약")) Then
        Get_ResultString = "방문예약"
    ElseIf (InStr(1, sRes, "정상")) Then
        Get_ResultString = Glo_Str_Reg(Val(sGATE))
    ElseIf (InStr(1, sRes, "미등록")) Then
        Get_ResultString = Glo_Str_Guest(Val(sGATE))
    ElseIf (InStr(1, sRes, "미인식")) Then
        Get_ResultString = Glo_Str_NoRec(Val(sGATE))
    ElseIf (InStr(1, sRes, "출입제한")) Then
        Get_ResultString = Glo_Str_BlackList(Val(sGATE))
    ElseIf (InStr(1, sRes, "영업")) Then
        Get_ResultString = Glo_Str_Taxi(Val(sGATE))
    ElseIf (InStr(1, sRes, "부제위반")) Then
        Get_ResultString = "부제위반"
    ElseIf (InStr(1, sRes, "요일")) Then
        Get_ResultString = Glo_Str_Day(Val(sGATE))
    End If
    
    If (InStr(1, sRes, "입차")) Then
        Get_ResultString = Get_ResultString & "입차"
    Else
        Get_ResultString = Get_ResultString & "출차"
    End If
    
End Function



Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
    Dim wFlags, Placement As Long
   Const SWPNOSIZE = &H1
   Const SWPNOMOVE = &H2
   Const SWPNOACTIVATE = &H10
   Const SWPSHOWWINDOW = &H40
   Const HWNDTOPMOST = -1
   Const HWNDNOTOPMOST = -2
    
   wFlags = SWPNOMOVE Or SWPNOSIZE Or SWPSHOWWINDOW Or SWPNOACTIVATE
    
   Select Case bTopMost
   Case True
       Placement = HWNDTOPMOST
   Case False
       Placement = HWNDNOTOPMOST
   End Select
    
   SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub


Public Sub RemoveCancelMenuItem(frm As Form)
    Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(frm.hwnd, 0)
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    
End Sub



Public Sub GuestForm_WindowState(value As Integer)

    Dim i As Integer

    For i = 0 To MAX_LANE_COUNT - 1
        If (Not Glo_FrmGuest(i) Is Nothing) Then '만들어져 있다면
            Glo_FrmGuest(i).WindowState = value
        End If
    Next i

End Sub


Public Function SaveClientKey(sSiteCode As String, sSiteName As String) As Boolean
    
    Dim rs As Recordset
    Dim qry As String
    
    On Error GoTo Err_P
    
    qry = "SELECT * From tb_certify WHERE HASHCODE = '" & Glo_PhyHDDKey & "' "
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    
    If (rs.EOF) Then
        adoConn.Execute "INSERT INTO tb_certify (IP, MAC, HASHCODE, SITECODE, SITENAME, C2DATE) VALUE ('" & Glo_IPAddr & "', '" & Glo_MacAddr & "', '" & Glo_PhyHDDKey & "', '" & sSiteCode & "', '" & sSiteName & "', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "') "
    Else
        adoConn.Execute "UPDATE tb_certify SET SITECODE = '" & sSiteCode & "', C2DATE = '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "' "
    End If
    
    '모니터링 호스트와 함께 운영할 경우 대비하여
    '메인호스트, 모니터링호스트 모두 같은 SiteCode, SiteName 을 갖도록 한다.
    adoConn.Execute "UPDATE tb_certify SET SITECODE='" & sSiteCode & "', SITENAME='" & sSiteName & "' "
    
    SaveClientKey = True
    
    Set rs = Nothing
    
    Exit Function
    
Err_P:
    SaveClientKey = False
    Set rs = Nothing
End Function


Public Sub DeleteClientKey(Scode As String)
    adoConn.Execute "UPDATE tb_certify SET SITECODE = '" & Scode & "', C2DATE = '' "
End Sub


Public Function GetSiteCode()
    Dim rs As Recordset
    Dim HDDKey As String
On Error GoTo Err_P

    GetSiteCode = ""
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT SITECODE FROM tb_Certify WHERE HASHCODE is NOT NULL", adoConn

    If Not (rs.EOF) Then
        GetSiteCode = "" & rs!SiteCode
    End If
    Set rs = Nothing
    
    
'    Call GetClienKey(HDDKey)
    
    Exit Function
    
Err_P:
    
    Call DebugLogger("[MainForm Activate ERR] " & Err.Description)
    Set rs = Nothing
End Function

Public Function GetSiteName()
    Dim rs As Recordset
On Error GoTo Err_P

    
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT SITENAME FROM tb_Certify WHERE HASHCODE is NOT NULL", adoConn

    If Not (rs.EOF) Then
        GetSiteName = "" & rs!SiteName
    Else
        GetSiteName = ""
    End If
    Set rs = Nothing
    
    Exit Function
    
Err_P:
    
    Call DebugLogger("[MainForm Activate ERR] " & Err.Description)
    Set rs = Nothing
End Function

'
'
''키생성
'Private Sub GetClienKey(sKey As String)
'    Dim msg
'
'    On Error Resume Next
'
''    msg = GetHDDID
'    msg = GetCPUID
'
'
'    If (Len(msg) > 0) Then
'        sKey = msg
'    Else
'        sKey = "키값 획득실패!!"
'    End If
'
'End Sub
'
'Public Function GetHDDID() As String
'    Dim oList, oObject As Object
'    Dim sDrive, msg As String
'
'On Error Resume Next
'
'    Set oList = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDisk") 'HDD고유값
'    For Each oObject In oList
'        sDrive = oObject.Path_.RelPath
'        If (InStr(sDrive, "C:") > 0) Then
'            'msg = msg & object.VolumeSerialNumber & vbCrLf
'            msg = oObject.VolumeSerialNumber
'            Exit For
'        End If
'    Next
'
'    GetHDDID = msg
'End Function
'
'Public Function GetCPUID() As String
'    Dim oWMI, oCPU As Object
'    Dim sCPUID As String
'On Error Resume Next
'    Set oWMI = GetObject("winmgmts:")
'    For Each oCPU In oWMI.InstancesOf("Win32_Processor")
'        sCPUID = sCPUID + oCPU.ProcessorID
'    Next
'
'    GetCPUID = sCPUID
'End Function


Public Sub ProtectMainMenuButton(argForm As Form)
    Dim i As Integer

    With argForm

        .Lblbutton(1).Caption = "보호해제"
        For i = 0 To 10
            .Lblbutton(i).Enabled = False
            .Imgbutton(i).Enabled = False
            .Lblbutton(i).Visible = False
            .Imgbutton(i).Visible = False
        Next i
        For i = 0 To 6 '입출차조회, 보호해제, 정기권관리, 정기권이력, 근무자관리, 환경설정, 종료
            .Lblbutton(i).Visible = True
            .Imgbutton(i).Visible = True
        Next i
        
        .Lblbutton(6).Enabled = True '시스템종료 버튼
        .Imgbutton(6).Enabled = True '
        
        .Lblbutton(1).Enabled = True '보호해제 버튼
        .Imgbutton(1).Enabled = True
        
        .Lblbutton(3).Visible = False '정기권이력, 미사용이므로 visible = false로 처리함
        .Imgbutton(3).Visible = False
    End With
End Sub

'1, 2, 4화면 전용 메뉴버튼 활성화
Public Sub ReleaseMainMenuButton(argForm As Form, argMenu As Variant)
    
    Dim i, j As Integer
    
    With argForm
    
        For i = 0 To UBound(argMenu)
            If (argMenu(i) = "입출차조회") Then
                .Lblbutton(0).ForeColor = vbBlack
                .Lblbutton(0).Enabled = True
                .Imgbutton(0).Enabled = True
            End If
            If (argMenu(i) = "정기권관리") Then
                .Lblbutton(2).ForeColor = vbBlack
                .Lblbutton(2).Enabled = True
                .Imgbutton(2).Enabled = True
            End If
            If (argMenu(i) = "정기권이력") Then
                .Lblbutton(3).ForeColor = vbBlack
                .Lblbutton(3).Enabled = True
                .Imgbutton(3).Enabled = True
                .Lblbutton(3).Visible = True
                .Imgbutton(3).Visible = True
            End If
            If (argMenu(i) = "방문예약") Then
                .Lblbutton(10).ForeColor = vbBlack
                .Lblbutton(10).Enabled = True
                .Imgbutton(10).Enabled = True
                .Lblbutton(10).Visible = True
                .Imgbutton(10).Visible = True
            End If
            If (argMenu(i) = "근무자관리") Then
                .Lblbutton(4).ForeColor = vbBlack
                .Lblbutton(4).Enabled = True
                .Imgbutton(4).Enabled = True
            End If
            If (argMenu(i) = "환경설정") Then
                .Lblbutton(5).ForeColor = vbBlack
                .Lblbutton(5).Enabled = True
                .Imgbutton(5).Enabled = True
            End If
            If (argMenu(i) = "무인정산기") Then
                .Lblbutton(7).ForeColor = vbBlack
                .Lblbutton(7).Enabled = True
                .Imgbutton(7).Enabled = True
                .Lblbutton(7).Visible = True
                .Imgbutton(7).Visible = True
                .Lblbutton(7).Caption = "무인정산기"
            End If
            If (argMenu(i) = "결제내역") Then
                .Lblbutton(7).ForeColor = vbBlack
                .Lblbutton(7).Enabled = True
                .Imgbutton(7).Enabled = True
                .Lblbutton(7).Visible = True
                .Imgbutton(7).Visible = True
                .Lblbutton(7).Caption = "결제내역"
            End If
        Next i

        .Lblbutton(1).Caption = "보호모드"
       
    End With
End Sub


'6화면 전용 메뉴버튼 활성화
Public Sub ReleaseMainMenuButton6Form(argForm As Form, argMenu As Variant)

    Dim i, j As Integer
    
    With argForm
    
        For i = 0 To UBound(argMenu)
            If (argMenu(i) = "입출차조회") Then
                .cmd_menu(0).Enabled = True
            End If
            If (argMenu(i) = "정기권관리") Then
                .cmd_menu(2).Enabled = True
            End If
            If (argMenu(i) = "정기권이력") Then
                .Lblbutton(3).ForeColor = vbBlack
                .Lblbutton(3).Enabled = True
                .Imgbutton(3).Enabled = True
            End If
            If (argMenu(i) = "방문예약") Then
                .cmd_menu(10).Enabled = True
                .cmd_menu(10).Visible = True
            End If
            If (argMenu(i) = "근무자관리") Then
                .cmd_menu(4).Enabled = True
            End If
            If (argMenu(i) = "환경설정") Then
                .cmd_menu(5).Enabled = True
            End If
            If (argMenu(i) = "무인정산기") Then
                .cmd_menu(7).Enabled = True
                .cmd_menu(7).Visible = True
                .cmd_menu(7).Caption = "무인정산기"
            End If
            If (argMenu(i) = "결제내역") Then
                .cmd_menu(7).Enabled = True
                .cmd_menu(7).Visible = True
                .cmd_menu(7).Caption = "결제내역"
            End If
        Next i

        .cmd_menu(1).Caption = "보호모드"
       
    End With
End Sub


Public Sub ProtectMainMenuButton6Form(argForm As Form)
    Dim i As Integer

    With argForm

        For i = 0 To 10
            .cmd_menu(i).Enabled = False
            .cmd_menu(i).Visible = False
        Next i
        For i = 0 To 6 '입출차조회, 보호해제, 정기권관리, 정기권이력, 근무자관리, 환경설정, 종료
            .cmd_menu(i).Visible = True
        Next i
        .cmd_menu(6).Enabled = True '시스템종료 버튼
        .cmd_menu(1).Enabled = True '보호해제 버튼
        
        .cmd_menu(3).Visible = False '정기권이력, 미사용이므로 visible = false로 처리함
        
        .cmd_menu(1).Caption = "보호해제"

    End With
End Sub

Public Sub ShowTitlebarSiteCode()
    Dim sSiteName As String
    Dim sSiteCode As String
    Dim sFrmTitle As String
    Dim sCustCode As String

    If (Glo_App_Cust_Code = "") Then
        sCustCode = ""
    Else
        sCustCode = " - 현장코드:" & Glo_App_Cust_Code
    End If
    
    Select Case Glo_Screen_No
        Case 1
            FrmG1.Caption = "주차관제 시스템 동작중" & sCustCode
        Case 2
            Jung.Caption = "주차관제 시스템 동작중" & sCustCode
        Case 4
            FrmG4Mini.Caption = "주차관제 시스템 동작중" & sCustCode
        Case 6
            FrmG6_23.Caption = "주차관제 시스템 동작중" & sCustCode
    End Select
End Sub



Public Sub Chk_FreePassEnable(argForm As Form, Index As Integer, bVal As Boolean)
    If (Index < Glo_Screen_No) Then
        argForm.Chk_FreePass(Index).Enabled = bVal
    End If
End Sub

Public Sub Chk_NormalPassEnable(argForm As Form, ByVal sLaneUse As String, ByVal sNormalUse As String, ByVal iIdx As Integer, ByVal sLaneName As String)
    
    If (iIdx < Glo_Screen_No) Then
        If (sLaneUse = "Y") Then
            argForm.Chk_FreePass(iIdx).Caption = sLaneName
            argForm.Chk_FreePass(iIdx).Enabled = True
            
            If (sNormalUse = "Y") Then
                argForm.Chk_FreePass(iIdx).value = 1
            End If
        Else
            argForm.Chk_FreePass(iIdx).Caption = "미사용"
            argForm.Chk_FreePass(iIdx).Enabled = False
            argForm.Chk_FreePass(iIdx).value = 0
        End If
    End If

End Sub

Public Sub Chk_TaxiPassEnable(argForm As Form, ByVal sLaneUse As String, ByVal sTaxiUse As String, ByVal iIdx As Integer, ByVal sLaneName As String)
    If (iIdx < Glo_Screen_No) Then
        If (sLaneUse = "Y") Then
            argForm.chk_Taxi(iIdx).Caption = sLaneName
            argForm.chk_Taxi(iIdx).Enabled = True
            
            If (sTaxiUse = "Y") Then
                argForm.chk_Taxi(iIdx).value = 1
            End If
        Else
            argForm.chk_Taxi(iIdx).Caption = "미사용"
            argForm.chk_Taxi(iIdx).Enabled = False
            argForm.chk_Taxi(iIdx).value = 0
        End If
    End If
End Sub
'임시 테스트 끝



Public Sub SelectMenuButton(argForm As Form, Index As Integer)
    Dim i As Integer

    Call GuestForm_WindowState(vbMinimized)
    
    argForm.MousePointer = 11
    Select Case Index
        Case 0
             'FrmInOut.Show 1
             FrmInOut.Show 0
             argForm.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "입출차 보고서 화면 접근")
        Case 2
             'FrmReg.Show 1
             FrmReg.Show 0
             argForm.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "정기권관리 화면 접근")
        Case 5
             If (Glo_Login_GUBUN = "총괄관리자") Then
                FrmTcpServer.Show 0
                argForm.MousePointer = 0
                Call DataLogger("[HOST Button]    " & "TCP Server 화면 접근")
             'ElseIf (Glo_Login_GUBUN = "관리자") Then
             Else
                FrmTcpServer2.Show 0
                argForm.MousePointer = 0
                Call DataLogger("[HOST Button]    " & "TCP Server2 화면 접근")
             End If
        Case 6
             Call DataLogger("[HOST Button]    " & "주차관제 시스템 종료!!!")
             Unload argForm
        Case 1
    '''         If (Lblbutton(1).Caption = "보호모드") Then
    '''            Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
    '''            Lblbutton(1).Caption = "보호해제"
    '''            For i = 0 To 8
    '''                Lblbutton(i).Enabled = False
    '''                Imgbutton(i).Enabled = False
    '''            Next i
    '''            Lblbutton(6).Enabled = True '시스템종료
    '''            Lblbutton(1).Enabled = True '보호해제
    '''            Imgbutton(6).Enabled = True
    '''            Imgbutton(1).Enabled = True
    '''
    '''            Lblbutton(7).Visible = False '무인정산기
    '''            Imgbutton(7).Visible = False
    '''
    '''            Lblbutton(10).Visible = False '방문예약
    '''            Imgbutton(10).Visible = False
    '''            Lblbutton(10).Enabled = False '방문예약
    '''            Imgbutton(10).Enabled = False
    '''
    '''            Put_Ini "System Config", "보호모드", "True"
    '''
    '''            Glo_Login_ID = ""
    '''            Glo_Login_PW = ""
    '''            Glo_Login_GUBUN = ""
    '''         Else
    '''            frmLogin.Show 1
    '''            Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
    '''            Lblbutton(1).Caption = "보호모드"
    '''            ListView1.SetFocus
    '''         End If
    '''         argForm.MousePointer = 0
    
             If (argForm.Lblbutton(1).Caption = "보호모드") Then
                'Call UnloadForms(Me) '모든 폼 제거(Jung, FrmTcpServer 폼은 제외)
                Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
                Call ProtectMainMenuButton(argForm)
                
                Glo_Login_ID = ""
                Glo_Login_PW = ""
                Glo_Login_GUBUN = ""
                Put_Ini "System Config", "보호모드", "True"
    
             Else
                Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
                frmLogin.Show 1
                'Lblbutton(1).Caption = "보호모드"
                argForm.ListView1.SetFocus
             End If
             argForm.MousePointer = 0
        Case 3
             'FrmRegHistory.Show 1
             FrmRegHistory.Show 0
             argForm.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "정기권 이력 화면 접근")
        Case 4
                'FrmId.Show 1
                FrmId.Show 0
                argForm.MousePointer = 0
                Call DataLogger("[HOST Button]    " & "아이디 관리 화면 접근")
        Case 7
            argForm.MousePointer = 0
            
'            If (Not (FrmG6_23.ActiveControl Is Nothing)) Then
'                If (argForm.ActiveControl = FrmG6_23.ActiveControl) Then
'                  sMenu = argForm.cmd_menu(Index).Caption
'                Else
'                  sMenu = argForm.cmd_menu(Index).Caption
'                End If
'             Else
'                sMenu = argForm.Lblbutton(Index).Caption
'             End If

            If (argForm.Lblbutton(Index).Caption = "무인정산기") Then
                FrmAccnt.Show 0
            ElseIf (argForm.Lblbutton(Index).Caption = "결제내역") Then
                frmResult.Show 1
            End If
            Call DataLogger("[HOST Button]    " & "무인정산기 관리 화면 접근")
        Case 8
            argForm.MousePointer = 1
            frmResult.Show 0
            Call DataLogger("[HOST Button]    " & "결제내역 화면 접근")
        Case 9
            argForm.MousePointer = 1
            'FrmGuestLog.Show 1
            FrmGuestLog.Show 0
            Call DataLogger("[HOST Button]    " & "방문객내역 화면 접근")
            
        Case 10  '방문차량 사전방문
            argForm.MousePointer = 1
            'FrmGuestRegLog.Show 1
            FrmGuestRegLog.Show 0
            Call DataLogger("[HOST Button]    " & "방문예약 화면 접근")
            Exit Sub
    
    End Select

End Sub


Public Sub RunHomeNet()
    Shell ("taskkill /f /im HomeNet.exe")
    If (IsFile("C:\HomeNet\HomeNet.exe") = True) Then
            
        Delay_Time (1)
        Shell ("C:\HomeNet\HomeNet.exe")
        Delay_Time (2)
'        FrmTcpServer.HomeSock.Close
'        FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
'        FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
'        FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
    End If
End Sub

