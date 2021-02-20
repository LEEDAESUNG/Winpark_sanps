Attribute VB_Name = "Definition"
Option Explicit

Public Glo_WebDC_YN As String '웹할인 기능 사용유무
Public Glo_GuestReg_YN As String '방문예약 기능 사용유무
Public Glo_MobileAlarm_YN As String '모바일알림 사용유무

Public Glo_COMPANY As String

Public Const MAX_LANE_COUNT = 6

'전광판 2단6열
Public Const Glo_DISP_ROW = 2
Public Const Glo_DISP_COL = 6
Public Enum enumDISP_EMG_TIME
    e1sec = 1       '긴급문구 표시시간(1초)
    e2sec = 2       '긴급문구 표시시간(2초)
    e3sec = 3       '긴급문구 표시시간(3초)
    e4sec = 4       '긴급문구 표시시간(4초)
    e5sec = 5       '긴급문구 표시시간(5초)
    e6sec = 6       '긴급문구 표시시간(6초)
    e7sec = 7       '긴급문구 표시시간(7초)
    e8sec = 8       '긴급문구 표시시간(8초)
    e9sec = 9       '긴급문구 표시시간(9초)
    e10sec = 10     '긴급문구 표시시간(10초)
    
'    e3sec = 3       '긴급문구 표시시간(3초)
'    e6sec = 6       '긴급문구 표시시간(6초)
'    e6sec = 10       '긴급문구 표시시간(6초)
'    e10sec = 20     '긴급문구 표시시간(10초)
'    e20sec = 40     '긴급문구 표시시간(20초)
'    e30sec = 60     '긴급문구 표시시간(30초)
End Enum

Public Enum enumDISP_NML_SHIFT
    eSTOP = &H1       '일반문구 정지
    eSHIFT = &H6      '일반문구 왼쪽으로 이동
End Enum
Public Glo_LANE_DISP_NML_SHIFT(MAX_LANE_COUNT) As Byte

Public Enum enumDIS_COLORs '전광판(풀컬러)
    eRED = &H1     '빨강
    eGreen = &H2   '녹색
    eYellow = &H3 '노랑
    eBLUE = &H4    '파랑
    eWINE = &H5    '자주색
    eSKY = &H6     '하늘색
    eWHITE = &H7   '흰색
End Enum
Public Enum enumDIS_COLOR2s '전광판(풀컬러)_FW7
    eRED = &H1     '빨강
    eBLUE = &H2    '파랑
    eWINE = &H3    '자주색(와인)
    eGreen = &H4   '녹색
    eYellow = &H5 '노랑
    eSKY = &H6     '하늘색
    eWHITE = &H7   '흰색
End Enum

Public Glo_App_Cust_Code As String


Public Glo_FrmIPCameraPlayer(MAX_LANE_COUNT) As Object ' 모바일에서 CCTV 보기위해 RTSP 주소를 저장한다

Public Glo_GateAgent_YN As String '차단기 TCP 제어용 중계서버 접속사용유무(VB특성상 TCP 처리 속도느리기 때문에 중계서버사용함)
Public Glo_GATE_AGENT1_PORT As Long
Public Glo_GATE_AGENT2_PORT As Long
Public Glo_GATE_AGENT3_PORT As Long
Public Glo_GATE_AGENT4_PORT As Long
Public Glo_GATE_AGENT5_PORT As Long
Public Glo_GATE_AGENT6_PORT As Long

Public Glo_APP_CHG_DAY As Long

Public Glo_SiteCode As String    '현장코드
Public Glo_SiteName As String    '현장명

Public Glo_Certify_PC As Boolean 'PC인증
Public Glo_IPAddr As String      'PC 외부IP
Public Glo_MacAddr As String     'PC 맥주소
Public Glo_PhyHDDKey As String   'PC 물리적 HDD 시리얼
Public Glo_CertServerIP As String 'PC인증 서버IP
Public Glo_CertServerPORT As Long 'PC인증 서버포트
Public GlO_CertPC_TcpData As String 'PC인증용

Public Enum enumCertify
    eCertNoTry = 0 '인증시도를 처음부터 안햇을 경우
    eCertTry = 1 '인증시도 했을 경우
    eCertOK = 2 '인증완료
End Enum
Public Glo_Certify As Integer           '밴더인증
Public Glo_Cert_LimitDate As String     '밴더인증 만료일
Public Glo_Cert_NoticeSDate As String   '밴더인증 만료일 안내 시작일
Public Glo_Cert_Month As Integer        '밴더인증 기간

Public Glo_GuestLogBackup_YN As String
Public Glo_GuestLogBackup_Month As Integer
Public Glo_GuestLogBackup_Time As String

Public glo_check As Boolean

Public Glo_GateNo_StartNo As Integer



Public Glo_Gate_ReconnCnt(MAX_LANE_COUNT) As Integer


'모바일 =====> 호스트
Public Const MO_GATE = "11"
Public Const MO_GATE_OPEN = "01"


'방문객정보
Public Type stGuest
    CarGubun As String
    ReserveSDate As String
    ReserveEDate As String
    Pass_YN As String
    InCarNo As String
    GuestName As String
    Dong  As String
    Ho  As String
    Tel  As String
    object  As String
    InGateNo  As String
    InDate  As String
    InImagePath  As String
    RegDate  As String
    ParkTime  As String
    
    OutCarNo  As String
    OutGateNo  As String
    OutDate  As String
    OutImagePath  As String
End Type

'Public Type stImageButton
'    Left As String
'    Top As String
'End Type
'

Public Type structLPR
    sLprIP As String
    sLaneInout As String
    sFreePass As String
    sBlackList As String
    sNoRecOpen As String
    sGateNo As String
    sTaxiPass As String
    sPassDate As String
    sNoWork As String '자리비움
End Type

Public Type structGate
    sGateOpen As String
    sGateStat As String
End Type

Public Type structSound
    sSnd_YN As String
    sSndReg_YN As String
    sSndGuest_YN As String
    sSndNoRec_YN As String
    sSndBlackList_YN As String
    sSndTaxi_YN As String
    sSndDay_YN As String
    sSndRegExpDate_YN As String
    sSndFName_Reg As String
    sSndFName_Guest As String
    sSndFName_NoRec As String
    sSndFName_BlackList As String
    sSndFName_Taxi As String
    sSndFName_Day As String '요일제위반사운드
    sSndFName_RegExpDate As String
    sSndFName_GuestRegCar As String
    sSndFName_GuestRegCarExpDate As String
End Type

Public Type structEmerg
    '긴급문구
    sEmergReg As String
    sEmergGuest As String
    sEmergNoRec As String
    sEmergBlackList As String
    sEmergTaxi As String
    sEmergDay As String
    sEmergRegExpDate As String
    sEmergGuestRegCar As String
    sEmergGuestRegCarExpDate As String
    
    '긴급문구 색상
    iDisp1EmergColorReg As Byte '등록차량 첫번째 문구색상
    iDisp2EmergColorReg As Byte '등록차량 두번째 문구색상
    iDisp1EmergColorGuest As Byte '미등록차량 첫번째 문구색상
    iDisp2EmergColorGuest As Byte '미등록차량 두번째 문구색상
    iDisp1EmergColorNoRec As Byte '미인식차량 첫번째 문구색상
    iDisp2EmergColorNoRec As Byte '미인식차량 두번째 문구색상
    iDisp1EmergColorBKList As Byte '블랙리스트 첫번째 문구색상
    iDisp2EmergColorBKList As Byte '블랙리스트 두번째 문구색상
    iDisp1EmergColorTaxi As Byte '영업용차량 첫번째 문구색상
    iDisp2EmergColorTaxi As Byte '영업용차량 두번째 문구색상
    iDisp1EmergColorDay As Byte '요일제차량 첫번째 문구색상
    iDisp2EmergColorDay As Byte '요일제차량 두번째 문구색상
    iDisp1EmergColorRegExpDate As Byte '등록기간초과차량 첫번째 문구색상
    iDisp2EmergColorRegExpDate As Byte '등록기간초과차량 두번째 문구색상
    iDisp1EmergColorGuestRegCar As Byte '사전방문예약차량 첫번째 문구색상
    iDisp2EmergColorGuestRegCar As Byte '사전방문예약차량 두번째 문구색상
    iDisp1EmergColorGuestRegCarExpDate As Byte '사전방문예약 기간초과차량 첫번째 문구색상
    iDisp2EmergColorGuestRegCarExpDate As Byte '사전방문예약 기간초과차량 두번째 문구색상
End Type

Public Glo_FrmGuest(MAX_LANE_COUNT) As Object
Public Glo_Guest_Print_Model(MAX_LANE_COUNT) As String '영수증프린터 모델("NONE", "랴", "LPT2", "FILE", "COM1~COM12")
Public Glo_Guest_Print_Port(MAX_LANE_COUNT) As String '영수증프린터 포트번호("NONE", "LPT1", "LPT2", "FILE", "COM1~COM12")
Public Glo_Guest_Print_Open(MAX_LANE_COUNT) As String '영수증프린터 오픈("Y", "N")
Public Glo_Guest_Gate_OpenDelay(MAX_LANE_COUNT) As Single
Public Glo_Receipt_Paper_Cut As String '영수증용지 절단 유무
Public Glo_Guest_YN As String                         '방문차량 처리가부(레인 1개라도 사용하면 "Y")

Public Const NWERR_GATE_OPEN = True ' 네트워크 에러 OR DB 장애일 경우 정상처리 불가능할경우, 차량진입시 차단기오픈 처리함
Public Const NWERR_GATE_STAY = False

Public APS_INFO_CarNo As String
Public APS_INFO_ParkTime As String
Public APS_INFO_AMT As String
Public APS_INFO_Pay As String
Public APS_INFO_DC As String

Public Glo_LPRBoard As String

Public Const MAX_LISTBOX_LINE = 100

Public Glo_RegMonFee_YN As String

Public Const MAX_REG_GUBUN = 10
Public Glo_RegGubun(MAX_REG_GUBUN) As String

Public Glo_Device_Awake As String


Public Glo_ParkFull_YN As String
Public Glo_ParkFull_Count As Long
Public Glo_ParkNow_Count As Long
Public Glo_ParkRegIn_YN As String
Public Glo_ParkFull_Status As enumParkFullStatus
Public Enum enumParkFullStatus
    pkfStayFULL = 1     '만차상태
    pkfChangeFULL = 2   '정상->만차로 변경
    pkfChangeNML = 3    '만차->정상으로 변경
    pkfStayNML = 4      '정상상태
End Enum
Public Glo_ParkFullLIGHT_YN As String '만차등
Public Glo_ParkFullLIGHT_EMPTY As String '여유
Public Glo_ParkFullLIGHT_BUSY As String '혼잡
Public Glo_ParkFullLIGHT_FULL As String '만차
Public Glo_ParkFullLIGHT_GUIDE As Long '기준:75%
Public Glo_ParkFullLight_DispMode As String
Public Glo_ParkFullLIGHT_IP As String '만차등
Public Glo_ParkFullLIGHT_PORT As Long '만차등
Public GlO_ParkFullLight_BData() As Byte
Public Glo_ParkFullLigth_Toggle As Boolean

Public Glo_Lane1_NoWork As String
Public Glo_Lane2_NoWork As String
Public Glo_Lane3_NoWork As String
Public Glo_Lane4_NoWork As String
Public Glo_Lane5_NoWork As String
Public Glo_Lane6_NoWork As String


Public Enum EnumEmergToggleOrder
    enumCarNo = 1
    enumCarStat = 2
End Enum

'긴급문구 세로출력 타이머 처리(차량번호, 처리결과 교대로 출력)
Public Type structEmergVertical
    
    CarNoCount As Integer     '차번출력할 횟수
    CarNo1 As String          '전광판 오른쪽 출력:차번 마지막 4자리 제외한 왼쪽부분
    CarNo2 As String          '전광판 왼쪽 출력:차번 마지막 4자리
    CarNoColor1 As Byte       '차번출력 색상
    CarNoColor2 As Byte       '차번출력 색상
    
    ToggleSelect As String '다음 출력할 내용 차례("차량번호" 또는 "처리결과", 초기값은 "처리결과")
    
    CarStatCount As Integer   '처리결과 출력할 횟수
    CarStat1 As String        '전광판 오른쪽 출력:처리결과1
    CarStat2 As String        '전광판 왼쪽 출력:처리결과2
    CarStatColor1 As Byte      '처리결과 색상
    CarStatColor2 As Byte      '처리결과 색상
End Type
Public Glo_Emerg_Vertical(MAX_LANE_COUNT) As structEmergVertical
'Public Const Glo_Emerg_Vertical_ToggleCount = 2 '1이상이어야 함
Public Glo_Emerg_Vertical_ToggleCount As Integer '1이상이어야 함

'긴급문구 "유지시간"이 3sec로 설정되어 있으므로 토글타이머는 이보다 같거나 짧은 시간으로 해야만 됨
'코드:GL_Emergency_Vertical => Head_Up(16) = "&H" & Hex(enumDISP_EMG_TIME.e3sec)
'Public Const Glo_Emerg_Vertical_ToggleTime = 2700 '세로 긴급문구 유지시간(차량번호, 처리결과) : 2700 ms
Public Glo_Emerg_Vertical_ToggleTime As Integer '세로 긴급문구 유지시간(차량번호, 처리결과) : 2700 ms



Public Glo_SOUND_YN As String
Public Glo_SND_Lane1_Reg_YN As String
Public Glo_SND_Lane1_Guest_YN As String
Public Glo_SND_Lane1_NoRec_YN As String
Public Glo_SND_Lane1_BlackList_YN As String
Public Glo_SND_Lane1_Taxi_YN As String
Public Glo_SND_Lane1_Day_YN As String
Public Glo_SND_Lane1_RegExpDate_YN As String
Public Glo_SND_Lane1_GuestRegCar_YN As String          '방문예약차량
Public Glo_SND_Lane1_GuestRegCarExpDate_YN As String   '방문예약만료
Public Glo_SND_Lane2_Reg_YN As String
Public Glo_SND_Lane2_Guest_YN As String
Public Glo_SND_Lane2_NoRec_YN As String
Public Glo_SND_Lane2_BlackList_YN As String
Public Glo_SND_Lane2_Taxi_YN As String
Public Glo_SND_Lane2_Day_YN As String
Public Glo_SND_Lane2_RegExpDate_YN As String
Public Glo_SND_Lane2_GuestRegCar_YN As String          '방문예약차량
Public Glo_SND_Lane2_GuestRegCarExpDate_YN As String   '방문예약만료
Public Glo_SND_Lane3_Reg_YN As String
Public Glo_SND_Lane3_Guest_YN As String
Public Glo_SND_Lane3_NoRec_YN As String
Public Glo_SND_Lane3_BlackList_YN As String
Public Glo_SND_Lane3_Taxi_YN As String
Public Glo_SND_Lane3_Day_YN As String
Public Glo_SND_Lane3_RegExpDate_YN As String
Public Glo_SND_Lane3_GuestRegCar_YN As String          '방문예약차량
Public Glo_SND_Lane3_GuestRegCarExpDate_YN As String   '방문예약만료
Public Glo_SND_Lane4_Reg_YN As String
Public Glo_SND_Lane4_Guest_YN As String
Public Glo_SND_Lane4_NoRec_YN As String
Public Glo_SND_Lane4_BlackList_YN As String
Public Glo_SND_Lane4_Taxi_YN As String
Public Glo_SND_Lane4_Day_YN As String
Public Glo_SND_Lane4_RegExpDate_YN As String
Public Glo_SND_Lane4_GuestRegCar_YN As String          '방문예약차량
Public Glo_SND_Lane4_GuestRegCarExpDate_YN As String   '방문예약만료
Public Glo_SND_Lane5_Reg_YN As String
Public Glo_SND_Lane5_Guest_YN As String
Public Glo_SND_Lane5_NoRec_YN As String
Public Glo_SND_Lane5_BlackList_YN As String
Public Glo_SND_Lane5_Taxi_YN As String
Public Glo_SND_Lane5_Day_YN As String
Public Glo_SND_Lane5_RegExpDate_YN As String
Public Glo_SND_Lane5_GuestRegCar_YN As String          '방문예약차량
Public Glo_SND_Lane5_GuestRegCarExpDate_YN As String   '방문예약만료
Public Glo_SND_Lane6_Reg_YN As String
Public Glo_SND_Lane6_Guest_YN As String
Public Glo_SND_Lane6_NoRec_YN As String
Public Glo_SND_Lane6_BlackList_YN As String
Public Glo_SND_Lane6_Taxi_YN As String
Public Glo_SND_Lane6_Day_YN As String
Public Glo_SND_Lane6_RegExpDate_YN As String
Public Glo_SND_Lane6_GuestRegCar_YN As String          '방문예약차량
Public Glo_SND_Lane6_GuestRegCarExpDate_YN As String   '방문예약만료

Public Glo_SNDFILE_Reg(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_Guest(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_NoRec(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_BlackList(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_Taxi(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_Day(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_RegExpDate(MAX_LANE_COUNT) As String
Public Glo_SNDFILE_GuestRegCar(MAX_LANE_COUNT) As String        '방문예약차량
Public Glo_SNDFILE_GuestRegCarExpDate(MAX_LANE_COUNT) As String '방문예약만료

Public Tmp_SNDFILE_Reg(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_Guest(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_NoRec(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_BlackList(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_Taxi(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_Day(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_RegExpDate(MAX_LANE_COUNT) As String
Public Tmp_SNDFILE_GuestRegCar(MAX_LANE_COUNT) As String        '방문예약차량
Public Tmp_SNDFILE_GuestRegCarExpDate(MAX_LANE_COUNT) As String '방문예약만료

Public Glo_Str_Reg(MAX_LANE_COUNT) As String
Public Glo_Str_Guest(MAX_LANE_COUNT) As String
Public Glo_Str_NoRec(MAX_LANE_COUNT) As String
Public Glo_Str_BlackList(MAX_LANE_COUNT) As String
Public Glo_Str_Taxi(MAX_LANE_COUNT) As String
Public Glo_Str_Day(MAX_LANE_COUNT) As String
Public Glo_Str_RegExpDate(MAX_LANE_COUNT) As String
Public Glo_Str_GuestRegCar(MAX_LANE_COUNT) As String        '방문예약차량
Public Glo_Str_GuestRegCarExpDate(MAX_LANE_COUNT) As String '방문예약만료

Public Glo_Disp1_Reg(MAX_LANE_COUNT) As String
Public Glo_Disp2_Reg(MAX_LANE_COUNT) As String
Public Glo_Disp1_Guest(MAX_LANE_COUNT) As String
Public Glo_Disp2_Guest(MAX_LANE_COUNT) As String
Public Glo_Disp1_NoRec(MAX_LANE_COUNT) As String
Public Glo_Disp2_NoRec(MAX_LANE_COUNT) As String
Public Glo_Disp1_BlackList(MAX_LANE_COUNT) As String
Public Glo_Disp2_BlackList(MAX_LANE_COUNT) As String
Public Glo_Disp1_Taxi(MAX_LANE_COUNT) As String
Public Glo_Disp2_Taxi(MAX_LANE_COUNT) As String
Public Glo_Disp1_Day(MAX_LANE_COUNT) As String
Public Glo_Disp2_Day(MAX_LANE_COUNT) As String
Public Glo_Disp1_RegExpDate(MAX_LANE_COUNT) As String
Public Glo_Disp2_RegExpDate(MAX_LANE_COUNT) As String
Public Glo_Disp1_GuestRegCar(MAX_LANE_COUNT) As String        '방문예약차량
Public Glo_Disp2_GuestRegCar(MAX_LANE_COUNT) As String '방문예약만료
Public Glo_Disp1_GuestRegCarExpDate(MAX_LANE_COUNT) As String        '방문예약차량
Public Glo_Disp2_GuestRegCarExpDate(MAX_LANE_COUNT) As String '방문예약만료

Public DB_Connect_F As Boolean
Public DB_Server_IP As String
Public DB_Server_Port As Long
Public DB_Rcv_LastTime As Long
Public DB_Conn_Msg As String

Public Glo_TestMode As String
Public Glo_Display  As String
Public Glo_Display_Direct  As String
Public Glo_NoRecOpen As String
Public Glo_FreePassLane As String ' 차량별 레인 통과허용 여부
Public Glo_WEEK_YN As String '요일제
Public Glo_ROTATION As String 'x부제
Public Glo_User_Type As String ' 구분1/구분2 or 동/호수

Public IniFileName$

Public MissMatch_YN As String
Public MissMatch_HomeNet_YN As String
Public MissMatch_F As Boolean

'Public TAXI_YN As String
'Public Glo_TAXI_IN_YN As String
'Public Glo_TAXI_OUT_YN As String
Public Glo_TAXI1_YN As String
Public Glo_TAXI2_YN As String
Public Glo_TAXI3_YN As String
Public Glo_TAXI4_YN As String
Public Glo_TAXI5_YN As String
Public Glo_TAXI6_YN As String
Public Taxi_F As Boolean

Public Glo_NOWORK1_YN As String
Public Glo_NOWORK2_YN As String
Public Glo_NOWORK3_YN As String
Public Glo_NOWORK4_YN As String
Public Glo_NOWORK5_YN As String
Public Glo_NOWORK6_YN As String

Public Glo_GUEST_LANE1_YN As String
Public Glo_GUEST_LANE2_YN As String
Public Glo_GUEST_LANE3_YN As String
Public Glo_GUEST_LANE4_YN As String
Public Glo_GUEST_LANE5_YN As String
Public Glo_GUEST_LANE6_YN As String

'Public HomeAlarm_Mode As Integer
'Public HomeSvr_IP As String
'Public HomeSvr_Port As Long
'Public Homesvr_ID As String
'Public Homesvr_PW As String

Public Glo_Mon_Lane(6) As Boolean
Public Glo_MonStat_Lane(6) As String
Public Glo_Mon_LastInTime As Long

Public Glo_Screen1 As Integer
Public Glo_Screen2 As Integer
Public Glo_Screen3 As Integer
Public Glo_Screen4 As Integer
Public Glo_Screen5 As Integer
Public Glo_Screen6 As Integer



Public Glo_Screen_No As Long
Public Glo_FreePass As String
Public Glo_Lane_Inout As String
Public Normal_Search_F As Boolean


Public Glo_Login_ID As String
Public Glo_Login_PW As String
Public Glo_Login_GUBUN As String

Public Glo_FreePassLane1_YN As String
Public Glo_FreePassLane2_YN As String
Public Glo_FreePassLane3_YN As String
Public Glo_FreePassLane4_YN As String
Public Glo_FreePassLane5_YN As String
Public Glo_FreePassLane6_YN As String

Public Glo_NoRecOpen1_YN As String
Public Glo_NoRecOpen2_YN As String
Public Glo_NoRecOpen3_YN As String
Public Glo_NoRecOpen4_YN As String
Public Glo_NoRecOpen5_YN As String
Public Glo_NoRecOpen6_YN As String

Public Glo_BlackList1_YN As String
Public Glo_BlackList2_YN As String
Public Glo_BlackList3_YN As String
Public Glo_BlackList4_YN As String
Public Glo_BlackList5_YN As String
Public Glo_BlackList6_YN As String

Public HostType As Integer
Public Glo_CarNum As String
Public Glo_RecNo As String
Public Glo_ProcNo As String
Public Glo_BlackList As String

Public Glo_Disp1 As String
Public Glo_Disp2 As String
Public Glo_Gate As String
Public Glo_GateNo As Integer
'Public Glo_GateGubun As Integer
Public Glo_Lpr_IP As String

Public User_Type As String
Public rs As ADODB.Recordset
Public rsGuestReg As ADODB.Recordset
Public rsGuestRegAdmin As ADODB.Recordset

'FrmAccnt
Public tmpValue As Long

Public Glo_Index As Long
Public Glo_SQL_REG As String

Public Glo_INOUT_USING_DATE As Long '입출차기록 보유기간

Public Glo_Reg_Qry As String    '정기권 등록창 쿼리
Public Glo_EndDate As Integer
Public Pwd_Cancel As Boolean

Public Record_Source As String
Public Report_Path_Name$
Public Doc_Path_Name$
Public Db_Path_Name$

Public PassWord As String
Public kyo_str(33) As String * 30

Public Glo_MsgRet As Boolean
Public Server_IP As String
Public Server_Port As Long
Public Const Server_WEBDCPort = 8000
Public Glo_GateBar_IP As String


Public Glo_JungSearch As String
Public Glo_JIOSch As String
Public Glo_APS1_Port As Long
Public Glo_APS2_Port As Long
Public Glo_APSCMD_Str As String * 22

Public Glo_SQL_PARKING As String
Public Glo_SQL_PART As String
Public Glo_SQL_PART_COUNT As String
Public Glo_SQL_PART_INFO As String
Public Glo_SQL_PART_REG As String
Public Glo_SQL_LPR_ID As String
Public Glo_SQL_SEARCH As String
Public Glo_SQL_COUNT As String

Public Glo_CLEAR_COUNT As String
Public Glo_DELETE_OPTION As String
Public Glo_SAVE_TERM As String

Public Glo_cmd_menu_index As Integer
Public Glo_PartName As String

'RemotePC
Public Glo_RemoteS_YN As String
Public Glo_RemoteS_IP As String
Public Glo_RemoteS_Port As Long
Public Glo_RemoteS_ScrPos As Integer
Public Glo_Remote_Str As String
Public Glo_RemoteR_YN As String
'Public Glo_RemoteR_IP As String
Public Glo_RemoteR_Port As Long

Public Glo_FreepassS_YN As String
Public Glo_FreepassS_IP As String
Public Glo_FreepassS_Port As Long
Public Glo_FreepassR_YN As String
Public Glo_FreepassR_Port As Long



'HomeNet
Public HomeNet_YN As String
Public HomeNet_IP As String
Public HomeNet_Port As Long

Public HomeNet_Str As String
Public HomeNet_Dong As String * 4
Public HomeNet_Ho As String * 4
Public HomeNet_CarNo As String * 16

'MVR
Public MVR_Str As String
Public MVR_YN As String
Public MVR_IP As String
Public MVR_Port As Long


'재 번호인식
Public Glo_ReANPR_YN As String



'APS(출구무인기)
Public Glo_ApsYN As String
Public Glo_Aps_IP As String
Public Glo_Aps_PORT As Long
Public Glo_APSCMD_Port As Long
Public Glo_APS_Str As String
Public Glo_APSCmdR_Port As Long

'APS(사전무인기)
Public Glo_PreApsYN As String
Public Glo_Grace_Time As Long
Public Glo_Return_Time As Long


'LPR Config
Public LANE1_YN As String
Public LANE1_Name As String
Public LANE1_Inout As String
Public LANE1_LPRMode As String
Public LANE1_LPRIP As String
Public LANE1_LPRPort As Long
Public LANE1_DeviceMode As String
Public LANE1_DeviceIP As String
Public LANE1_DisplayMode As String
Public LANE1_DispIP As String
Public LANE1_DispPort As Long
Public LANE1_RelayPort As Long
Public LANE1_DispComPort As Integer
Public LANE1_RelayComPort As Integer
Public LANE1_Disp1Msg As String
Public LANE1_Disp2Msg As String
Public LANE1_Disp1Color As String
Public LANE1_Disp2Color As String
Public LANE1_DispSpeed As String

Public LANE2_YN As String
Public LANE2_Name As String
Public LANE2_Inout As String
Public LANE2_LPRMode As String
Public LANE2_LPRIP As String
Public LANE2_LPRPort As Long
Public LANE2_DeviceMode As String
Public LANE2_DeviceIP As String
Public LANE2_DisplayMode As String
Public LANE2_DispIP As String
Public LANE2_DispPort As Long
Public LANE2_RelayPort As Long
Public LANE2_DispComPort As Integer
Public LANE2_RelayComPort As Integer
Public LANE2_Disp1Msg As String
Public LANE2_Disp2Msg As String
Public LANE2_Disp1Color As String
Public LANE2_Disp2Color As String
Public LANE2_DispSpeed As String

Public LANE3_YN As String
Public LANE3_Name As String
Public LANE3_Inout As String
Public LANE3_LPRMode As String
Public LANE3_LPRIP As String
Public LANE3_LPRPort As Long
Public LANE3_DeviceMode As String
Public LANE3_DeviceIP As String
Public LANE3_DisplayMode As String
Public LANE3_DispIP As String
Public LANE3_DispPort As Long
Public LANE3_RelayPort As Long
Public LANE3_DispComPort As Integer
Public LANE3_RelayComPort As Integer
Public LANE3_Disp1Msg As String
Public LANE3_Disp2Msg As String
Public LANE3_Disp1Color As String
Public LANE3_Disp2Color As String
Public LANE3_DispSpeed As String

Public LANE4_YN As String
Public LANE4_Name As String
Public LANE4_Inout As String
Public LANE4_LPRMode As String
Public LANE4_LPRIP As String
Public LANE4_LPRPort As Long
Public LANE4_DeviceMode As String
Public LANE4_DeviceIP As String
Public LANE4_DisplayMode As String
Public LANE4_DispIP As String
Public LANE4_DispPort As Long
Public LANE4_RelayPort As Long
Public LANE4_DispComPort As Integer
Public LANE4_RelayComPort As Integer
Public LANE4_Disp1Msg As String
Public LANE4_Disp2Msg As String
Public LANE4_Disp1Color As String
Public LANE4_Disp2Color As String
Public LANE4_DispSpeed As String

Public LANE5_YN As String
Public LANE5_Name As String
Public LANE5_Inout As String
Public LANE5_LPRMode As String
Public LANE5_LPRIP As String
Public LANE5_LPRPort As Long
Public LANE5_DeviceMode As String
Public LANE5_DeviceIP As String
Public LANE5_DisplayMode As String
Public LANE5_DispIP As String
Public LANE5_DispPort As Long
Public LANE5_RelayPort As Long
Public LANE5_DispComPort As Integer
Public LANE5_RelayComPort As Integer
Public LANE5_Disp1Msg As String
Public LANE5_Disp2Msg As String
Public LANE5_Disp1Color As String
Public LANE5_Disp2Color As String
Public LANE5_DispSpeed As String

Public LANE6_YN As String
Public LANE6_Name As String
Public LANE6_Inout As String
Public LANE6_LPRMode As String
Public LANE6_LPRIP As String
Public LANE6_LPRPort As Long
Public LANE6_DeviceMode As String
Public LANE6_DeviceIP As String
Public LANE6_DisplayMode As String
Public LANE6_DispIP As String
Public LANE6_DispPort As Long
Public LANE6_RelayPort As Long
Public LANE6_DispComPort As Integer
Public LANE6_RelayComPort As Integer
Public LANE6_Disp1Msg As String
Public LANE6_Disp2Msg As String
Public LANE6_Disp1Color As String
Public LANE6_Disp2Color As String
Public LANE6_DispSpeed As String

Public AdoHome_Str As String
Public Homers As ADODB.Recordset

'호스트 =====> 무인정산기
Public Const CM_DAY = "00"      ' 일일최대금액 (보류)
Public Const CM_PER = "01"      ' %할인
Public Const CM_HOUR = "02"     ' 시간할인
Public Const CM_WON = "03"      ' 금액 할인
Public Const CM_DATE = "04"     ' 입차시간 강제입력 20151022143005

Public Const CM_CANCEL = "10"   ' 정산취소
Public Const CM_INITAL = "11"   ' 초기화면으로 전환
Public Const CM_GATE = "12"     ' 차단기 오픈
Public Const CM_PRINT = "13"    ' 영수증 출력
Public Const CM_REPRINT = "14"    ' 영수증 재출력
Public Const CM_CARDCANCEL = "15"  '카드결제 취소

'무인정산기 =====> 호스트
Public Const CM_START = "40"    ' 무인정산기 START
Public Const CM_END = "41"    ' 무인정산기 END
Public Const CM_RESPONSE = "42"   ' 호스트 명령 응답
Public Const CM_NOPAY = "43"   ' 무료처리
Public Const CM_MSG = "44"
Public Const CM_JUNGSANCANCEL = "50"    ' 정산취소버튼 누름
Public Const CM_CHANGEOUTERR = "51"    ' 거스름돈 배출에러
Public Const CM_DISPENSER1000ERR = "52"    ' 1000원권 지폐방출기에러
Public Const CM_DISPENSER5000ERR = "53"    ' 5000원권 지폐방출기에러
Public Const CM_COINERR = "54"    ' 코인기에러
Public Const CM_BILLERR = "55"    ' 지폐인식기에러
Public Const CM_CAROUT = "56"    ' 입차 정보
Public Const CM_FILTER = "57"    ' 입차 정보가 필터링을 통해 과금되었슴
Public Const CM_NOCAR = "58"    ' 입차 정보가 없슴
Public Const CM_SERVICECARDERR = "59"   ' 할인권에러
Public Const CM_CREDITCARDERR = "60"    ' 신용카드에러
Public Const CM_CREDITCARDCANCEL = "61"    ' 신용카드 결제취소

Public F_Key1 As String
Public F_Key2 As String
Public F_Key3 As String
Public F_Key4 As String
Public F_Key5 As String
Public F_Key6 As String
Public F_Key7 As String
Public F_Key8 As String
Public F_Key9 As String
Public F_Key10 As String
Public F_Key11 As String
Public F_Key12 As String


'후방카메라 관련 설정
Public Glo_Lane1_Back_YN As String 'Lane1 후방카메라 사용유무
Public Glo_Lane2_Back_YN As String 'Lane2 후방카메라 사용유무
Public Glo_Lane3_Back_YN As String 'Lane3 후방카메라 사용유무
Public Glo_Lane4_Back_YN As String 'Lane4 후방카메라 사용유무
Public Glo_Lane5_Back_YN As String 'Lane5 후방카메라 사용유무
Public Glo_Lane6_Back_YN As String 'Lane6 후방카메라 사용유무

Public Glo_Lane1_Front_CarNo As String 'Lane1 전방 차량번호
Public Glo_Lane1_Front_PassDate As String 'Lane1 전방 차량번호 인식시간
Public Glo_Lane2_Front_CarNo As String 'Lane2 전방 차량번호
Public Glo_Lane2_Front_PassDate As String 'Lane2 전방 차량번호 인식시간
Public Glo_Lane3_Front_CarNo As String 'Lane3 전방 차량번호
Public Glo_Lane3_Front_PassDate As String 'Lane3 전방 차량번호 인식시간
Public Glo_Lane4_Front_CarNo As String 'Lane4 전방 차량번호
Public Glo_Lane4_Front_PassDate As String 'Lane4 전방 차량번호 인식시간
Public Glo_Lane5_Front_CarNo As String 'Lane5 전방 차량번호
Public Glo_Lane5_Front_PassDate As String 'Lane5 전방 차량번호 인식시간
Public Glo_Lane6_Front_CarNo As String 'Lane6 전방 차량번호
Public Glo_Lane6_Front_PassDate As String 'Lane6 전방 차량번호 인식시간

'차단기닫기 버튼 사용유무
Public Glo_Lane1_GateClose_YN As String 'Lane1 차단기닫기 버튼 사용유무
Public Glo_Lane2_GateClose_YN As String 'Lane2 차단기닫기 버튼 사용유무
Public Glo_Lane3_GateClose_YN As String 'Lane3 차단기닫기 버튼 사용유무
Public Glo_Lane4_GateClose_YN As String 'Lane4 차단기닫기 버튼 사용유무
Public Glo_Lane5_GateClose_YN As String 'Lane5 차단기닫기 버튼 사용유무
Public Glo_Lane6_GateClose_YN As String 'Lane6 차단기닫기 버튼 사용유무


'Account 내역 읽어오기
Public Sub Read_Account()
    Dim qry As String
    Dim rs As ADODB.Recordset
    Dim SQL_SEARCH As String
    Dim bQryResult  As Boolean
    Dim i As Long
    Dim RegDate As String
    
On Error GoTo Err_p

    With FrmAccnt
    
        qry = "SELECT * From tb_account"
        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            Call DataLogger("[Read Account]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        If (rs.EOF = False) Then
            
            .LblBill10000.Caption = rs!BILL_S10000
            .LblBill5000.Caption = rs!BILL_S5000
            .LblBill1000.Caption = rs!BILL_S1000
            
            .txt_500.text = rs!COIN_C500 + rs!COIN_H500
            .txt_100.text = rs!COIN_C100 + rs!COIN_H100
            
            .txt_5000.text = rs!BILL_H5000
            .txt_1000.text = rs!BILL_H1000
            
            .lbl_Update.Caption = "Update Date : " & rs!Update_date
            
            i = (10000 * rs!BILL_S10000) + (5000 * rs!BILL_S5000) + (1000 * rs!BILL_S1000) + (5000 * rs!BILL_H5000) + (1000 * rs!BILL_H1000) + (500 * rs!COIN_C500) + (100 * rs!COIN_C100) + (500 * rs!COIN_H500) + (100 * rs!COIN_H100)
            .lbl_TotalAccnt.Caption = i
        Else
            'Insert 0 초기화
            RegDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
            
            qry = "INSERT INTO tb_account (BILL_S10000, BILL_S5000, BILL_S1000, BILL_H5000, BILL_H1000, COIN_H500, COIN_H100, COIN_C500, COIN_C100, COIN_C50, COIN_C10, REG_DATE, UPDATE_DATE, GUBUN ) VALUES (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '" & RegDate & "', '" & RegDate & "', 'POS1')"
            Set rs = New ADODB.Recordset
            'rs.Open Qry, adoConn
            bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
            If (bQryResult = False) Then
                Call DataLogger("[Read Account 0]    " & "네트워크 및 DB 점검바랍니다")
                Exit Sub
            End If
            
            
            .LblBill10000.Caption = 0
            .LblBill5000.Caption = 0
            .LblBill1000.Caption = 0
            
            .txt_500.text = 0
            .txt_100.text = 0
            
            .txt_5000.text = 0
            .txt_1000.text = 0
            
            .lbl_Update.Caption = "Update Date : " & "-"
            
            i = 0
            .lbl_TotalAccnt.Caption = i
            
            
        End If
        
        Set rs = Nothing
        
    End With
    
    Exit Sub
    
Err_p:
    Call DataLogger("[Read_Account Err] " & Err.Description)
End Sub


'전역변수 초기화
Public Sub CFG_Init()


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 표준 호스트PC는 LPR과 같이 설치되므로, 각 레인별 초기값은 모니터링용 레인이 아닌 것으로 설정합니다.
    ' 하지만, RemoteR_Socket으로부터 레인별 상태 데이터를 한번이라도 받으면, 모니터링용 레인으로 전환/처리된다.
    Dim i As Integer
    For i = 0 To MAX_LANE_COUNT - 1
        Glo_Mon_Lane(i) = False
        Glo_MonStat_Lane(i) = ""
    Next i
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim sip_pos As Integer
    Dim eip_pos As Integer
    
    
    Glo_GateAgent_YN = Get_Ini("System Config", "SocketAgent_YN", "N")
    Glo_GATE_AGENT1_PORT = Val(Get_Ini("System Config", "SOCKET_AGENT1_PORT", "30101"))
    Glo_GATE_AGENT2_PORT = Val(Get_Ini("System Config", "SOCKET_AGENT2_PORT", "30102"))
    Glo_GATE_AGENT3_PORT = Val(Get_Ini("System Config", "SOCKET_AGENT3_PORT", "30103"))
    Glo_GATE_AGENT4_PORT = Val(Get_Ini("System Config", "SOCKET_AGENT4_PORT", "30104"))
    Glo_GATE_AGENT5_PORT = Val(Get_Ini("System Config", "SOCKET_AGENT5_PORT", "30105"))
    Glo_GATE_AGENT6_PORT = Val(Get_Ini("System Config", "SOCKET_AGENT6_PORT", "30106"))
    
    
    ' 정기권 구분 항목 로드
    For i = 1 To MAX_REG_GUBUN
        Glo_RegGubun(i) = Get_Ini("System Config", "RegGubun" & i, "")
    Next
    
   
    Glo_App_Cust_Code = Get_Ini("System Config", "APP_CUST_CODE", "")
    
    
    Glo_APP_CHG_DAY = Val(Get_Ini("System Config", "APP_PW_CHG_DAY", "0"))
    
    
    Glo_Cert_Month = Val(Get_Ini("System Config", "CERT_MONTH", "12"))
    If (Glo_Cert_Month > 12) Then
        Glo_Cert_Month = 12
    ElseIf (Glo_Cert_Month < 2) Then
        Glo_Cert_Month = 2
    End If
    
    '방문객DB 백업설정
    Glo_GuestLogBackup_YN = Get_Ini("System Config", "GuestLogBackup_YN", "N")
    Glo_GuestLogBackup_Month = Val(Get_Ini("System Config", "GuestLogBackup_Month", "3"))
    Glo_GuestLogBackup_Time = Val(Get_Ini("System Config", "GuestLogBackup_Time", "02:00"))
    
    '만차
    Glo_ParkFull_YN = Get_Ini("System Config", "ParkFull_YN", "N")
    Glo_ParkFull_Count = Val(Get_Ini("System Config", "ParkFull_Count", "N"))
    Glo_ParkNow_Count = Val(Get_Ini("System Config", "ParkNow_Count", "N"))
    Glo_ParkRegIn_YN = Get_Ini("System Config", "ParkRegIn_YN", "N")
    
    '만차등
    Glo_ParkFullLIGHT_YN = Get_Ini("System Config", "ParkFullLIGHT_YN", "N")
    Glo_ParkFullLIGHT_EMPTY = Get_Ini("System Config", "ParkFullLight_EMPTY", "여유")
    Glo_ParkFullLIGHT_BUSY = Get_Ini("System Config", "ParkFullLight_BUSY", "혼잡")
    Glo_ParkFullLIGHT_FULL = Get_Ini("System Config", "ParkFullLight_FULL", "만차")
    Glo_ParkFullLIGHT_GUIDE = Val(Get_Ini("System Config", "ParkFullLight_GUIDE ", "75"))
    Glo_ParkFullLIGHT_IP = Get_Ini("System Config", "ParkFullLIGHT_IP", "255.255.255.255")
    Glo_ParkFullLight_DispMode = Get_Ini("System Config", "ParkFullLight_DispMode", "0")
    
    Glo_LPRBoard = Get_Ini("System Config", "LPRBoard", "위즈넷")
    
    
    Glo_RegMonFee_YN = Get_Ini("System Config", "RegMonFee_YN", "N")
    Glo_Device_Awake = Get_Ini("System Config", "Device_Awake", "N")
    
    Glo_TestMode = Get_Ini("System Config", "TestMode", "N")
    
    Glo_Display = Get_Ini("System Config", "Display", "전광판(풀컬러)")
    Glo_Display_Direct = Get_Ini("System Config", "Display_Direct", "가로")
    
    'LPR시작 번호
    Glo_GateNo_StartNo = Val(Get_Ini("System Config", "GateNo_StartNo", "0"))
    
    'Remote Config
    Glo_RemoteS_YN = Get_Ini("System Config", "RemoteS_YN", "N")
    Glo_RemoteS_IP = Get_Ini("System Config", "RemoteS_IP", "127.0.0.1")
    Glo_RemoteS_Port = Val(Get_Ini("System Config", "RemoteS_Port", "4000"))
    Glo_RemoteS_ScrPos = Val(Get_Ini("System Config", "RemoteS_ScrPos", "0"))
    Glo_RemoteR_YN = Get_Ini("System Config", "RemoteR_YN", "N")
    'Glo_RemoteR_IP = Get_Ini("System Config", "RemoteR_IP", "127.0.0.1")
    Glo_RemoteR_Port = Val(Get_Ini("System Config", "RemoteR_Port", "4000"))
    
    Glo_FreepassS_YN = Get_Ini("System Config", "FreepassS_YN", "N")
    Glo_FreepassS_IP = Get_Ini("System Config", "FreepassS_IP", "127.0.0.1")
    Glo_FreepassS_Port = Val(Get_Ini("System Config", "FreepassS_Port", "18280"))
    Glo_FreepassR_YN = Get_Ini("System Config", "FreepassR_YN", "N")
    Glo_FreepassR_Port = Val(Get_Ini("System Config", "FreepassR_Port", "18280"))
    
    
    HomeNet_YN = Get_Ini("System Config", "HomeNet_YN", "N")
    HomeNet_IP = Get_Ini("System Config", "HomeNet_IP", "127.0.0.1")
    HomeNet_Port = Val(Get_Ini("System Config", "HomeNet_Port", "18497"))
    
    MVR_YN = Get_Ini("System Config", "MVR_YN", "N")
    MVR_IP = Get_Ini("System Config", "MVR_IP", "127.0.0.1")
    MVR_Port = Val(Get_Ini("System Config", "MVR_Port", "18496"))
    
    Glo_ReANPR_YN = Get_Ini("System Config", "ReANPR_YN", "N")
    
    
    
    
    Glo_NOWORK1_YN = Get_Ini("System Config", "NOWORK1_YN", "N")
    Glo_NOWORK2_YN = Get_Ini("System Config", "NOWORK2_YN", "N")
    Glo_NOWORK3_YN = Get_Ini("System Config", "NOWORK3_YN", "N")
    Glo_NOWORK4_YN = Get_Ini("System Config", "NOWORK4_YN", "N")
    Glo_NOWORK5_YN = Get_Ini("System Config", "NOWORK5_YN", "N")
    Glo_NOWORK6_YN = Get_Ini("System Config", "NOWORK6_YN", "N")
    
    
        
    
'    TAXI_YN = Get_Ini("System Config", "TAXI_YN", "N")
'    Glo_TAXI_IN_YN = Get_Ini("System Config", "TAXI_IN_YN", "N")
'    Glo_TAXI_OUT_YN = Get_Ini("System Config", "TAXI_OUT_YN", "N")
    Glo_TAXI1_YN = Get_Ini("System Config", "TAXI1_YN", "N")
    Glo_TAXI2_YN = Get_Ini("System Config", "TAXI2_YN", "N")
    Glo_TAXI3_YN = Get_Ini("System Config", "TAXI3_YN", "N")
    Glo_TAXI4_YN = Get_Ini("System Config", "TAXI4_YN", "N")
    Glo_TAXI5_YN = Get_Ini("System Config", "TAXI5_YN", "N")
    Glo_TAXI6_YN = Get_Ini("System Config", "TAXI6_YN", "N")
    
    
    '후방카메라 로드
    Glo_Lane1_Back_YN = Get_Ini("System Config", "LANE1_BACK_YN", "N")
    Glo_Lane2_Back_YN = Get_Ini("System Config", "Lane2_BACK_YN", "N")
    Glo_Lane3_Back_YN = Get_Ini("System Config", "Lane3_BACK_YN", "N")
    Glo_Lane4_Back_YN = Get_Ini("System Config", "Lane4_BACK_YN", "N")
    Glo_Lane5_Back_YN = Get_Ini("System Config", "Lane5_BACK_YN", "N")
    Glo_Lane6_Back_YN = Get_Ini("System Config", "Lane6_BACK_YN", "N")
    
    
    '차단기닫기버튼 사용유무
    Glo_Lane1_GateClose_YN = Get_Ini("System Config", "Lane1_GateClose_YN", "N")
    Glo_Lane2_GateClose_YN = Get_Ini("System Config", "Lane2_GateClose_YN", "N")
    Glo_Lane3_GateClose_YN = Get_Ini("System Config", "Lane3_GateClose_YN", "N")
    Glo_Lane4_GateClose_YN = Get_Ini("System Config", "Lane4_GateClose_YN", "N")
    Glo_Lane5_GateClose_YN = Get_Ini("System Config", "Lane5_GateClose_YN", "N")
    Glo_Lane6_GateClose_YN = Get_Ini("System Config", "Lane6_GateClose_YN", "N")
    
    
    '사운드 로드
    Glo_SOUND_YN = Get_Ini("System Config", "SOUND_YN", "N")
    Glo_SND_Lane1_Reg_YN = Get_Ini("System Config", "SND_Lane1_Reg_YN", "N")
    Glo_SND_Lane1_Guest_YN = Get_Ini("System Config", "SND_Lane1_Guest_YN", "N")
    Glo_SND_Lane1_NoRec_YN = Get_Ini("System Config", "SND_Lane1_NoRec_YN", "N")
    Glo_SND_Lane1_BlackList_YN = Get_Ini("System Config", "SND_Lane1_BlackList_YN", "N")
    Glo_SND_Lane1_Taxi_YN = Get_Ini("System Config", "SND_Lane1_Taxi_YN", "N")
    Glo_SND_Lane1_Day_YN = Get_Ini("System Config", "SND_Lane1_Day_YN", "N")
    Glo_SND_Lane1_RegExpDate_YN = Get_Ini("System Config", "SND_Lane1_RegExpDate_YN", "N")
    Glo_SND_Lane1_GuestRegCar_YN = Get_Ini("System Config", "SND_Lane1_GuestRegCar_YN", "N")
    Glo_SND_Lane1_GuestRegCarExpDate_YN = Get_Ini("System Config", "SND_Lane1_GuestRegCarExpDate_YN", "N")
    
    Glo_SND_Lane2_Reg_YN = Get_Ini("System Config", "SND_Lane2_Reg_YN", "N")
    Glo_SND_Lane2_Guest_YN = Get_Ini("System Config", "SND_Lane2_Guest_YN", "N")
    Glo_SND_Lane2_NoRec_YN = Get_Ini("System Config", "SND_Lane2_NoRec_YN", "N")
    Glo_SND_Lane2_BlackList_YN = Get_Ini("System Config", "SND_Lane2_BlackList_YN", "N")
    Glo_SND_Lane2_Taxi_YN = Get_Ini("System Config", "SND_Lane2_Taxi_YN", "N")
    Glo_SND_Lane2_Day_YN = Get_Ini("System Config", "SND_Lane2_Day_YN", "N")
    Glo_SND_Lane2_RegExpDate_YN = Get_Ini("System Config", "SND_Lane2_RegExpDate_YN", "N")
    Glo_SND_Lane2_GuestRegCar_YN = Get_Ini("System Config", "SND_Lane2_GuestRegCar_YN", "N")
    Glo_SND_Lane2_GuestRegCarExpDate_YN = Get_Ini("System Config", "SND_Lane2_GuestRegCarExpDate_YN", "N")
    Glo_SND_Lane3_Reg_YN = Get_Ini("System Config", "SND_Lane3_Reg_YN", "N")
    Glo_SND_Lane3_Guest_YN = Get_Ini("System Config", "SND_Lane3_Guest_YN", "N")
    Glo_SND_Lane3_NoRec_YN = Get_Ini("System Config", "SND_Lane3_NoRec_YN", "N")
    Glo_SND_Lane3_BlackList_YN = Get_Ini("System Config", "SND_Lane3_BlackList_YN", "N")
    Glo_SND_Lane3_Taxi_YN = Get_Ini("System Config", "SND_Lane3_Taxi_YN", "N")
    Glo_SND_Lane3_Day_YN = Get_Ini("System Config", "SND_Lane3_Day_YN", "N")
    Glo_SND_Lane3_RegExpDate_YN = Get_Ini("System Config", "SND_Lane3_RegExpDate_YN", "N")
    Glo_SND_Lane3_GuestRegCar_YN = Get_Ini("System Config", "SND_Lane3_GuestRegCar_YN", "N")
    Glo_SND_Lane3_GuestRegCarExpDate_YN = Get_Ini("System Config", "SND_Lane3_GuestRegCarExpDate_YN", "N")
    Glo_SND_Lane4_Reg_YN = Get_Ini("System Config", "SND_Lane4_Reg_YN", "N")
    Glo_SND_Lane4_Guest_YN = Get_Ini("System Config", "SND_Lane4_Guest_YN", "N")
    Glo_SND_Lane4_NoRec_YN = Get_Ini("System Config", "SND_Lane4_NoRec_YN", "N")
    Glo_SND_Lane4_BlackList_YN = Get_Ini("System Config", "SND_Lane4_BlackList_YN", "N")
    Glo_SND_Lane4_Taxi_YN = Get_Ini("System Config", "SND_Lane4_Taxi_YN", "N")
    Glo_SND_Lane4_Day_YN = Get_Ini("System Config", "SND_Lane4_Day_YN", "N")
    Glo_SND_Lane4_RegExpDate_YN = Get_Ini("System Config", "SND_Lane4_RegExpDate_YN", "N")
    Glo_SND_Lane4_GuestRegCar_YN = Get_Ini("System Config", "SND_Lane4_GuestRegCar_YN", "N")
    Glo_SND_Lane4_GuestRegCarExpDate_YN = Get_Ini("System Config", "SND_Lane4_GuestRegCarExpDate_YN", "N")
    Glo_SND_Lane5_Reg_YN = Get_Ini("System Config", "SND_Lane5_Reg_YN", "N")
    Glo_SND_Lane5_Guest_YN = Get_Ini("System Config", "SND_Lane5_Guest_YN", "N")
    Glo_SND_Lane5_NoRec_YN = Get_Ini("System Config", "SND_Lane5_NoRec_YN", "N")
    Glo_SND_Lane5_BlackList_YN = Get_Ini("System Config", "SND_Lane5_BlackList_YN", "N")
    Glo_SND_Lane5_Taxi_YN = Get_Ini("System Config", "SND_Lane5_Taxi_YN", "N")
    Glo_SND_Lane5_Day_YN = Get_Ini("System Config", "SND_Lane5_Day_YN", "N")
    Glo_SND_Lane5_RegExpDate_YN = Get_Ini("System Config", "SND_Lane5_RegExpDate_YN", "N")
    Glo_SND_Lane5_GuestRegCar_YN = Get_Ini("System Config", "SND_Lane5_GuestRegCar_YN", "N")
    Glo_SND_Lane5_GuestRegCarExpDate_YN = Get_Ini("System Config", "SND_Lane5_GuestRegCarExpDate_YN", "N")
    Glo_SND_Lane6_Reg_YN = Get_Ini("System Config", "SND_Lane6_Reg_YN", "N")
    Glo_SND_Lane6_Guest_YN = Get_Ini("System Config", "SND_Lane6_Guest_YN", "N")
    Glo_SND_Lane6_NoRec_YN = Get_Ini("System Config", "SND_Lane6_NoRec_YN", "N")
    Glo_SND_Lane6_BlackList_YN = Get_Ini("System Config", "SND_Lane6_BlackList_YN", "N")
    Glo_SND_Lane6_Taxi_YN = Get_Ini("System Config", "SND_Lane6_Taxi_YN", "N")
    Glo_SND_Lane6_Day_YN = Get_Ini("System Config", "SND_Lane6_Day_YN", "N")
    Glo_SND_Lane6_RegExpDate_YN = Get_Ini("System Config", "SND_Lane6_RegExpDate_YN", "N")
    Glo_SND_Lane6_GuestRegCar_YN = Get_Ini("System Config", "SND_Lane6_GuestRegCar_YN", "N")
    Glo_SND_Lane6_GuestRegCarExpDate_YN = Get_Ini("System Config", "SND_Lane6_GuestRegCarExpDate_YN", "N")
    
    Glo_Str_Reg(0) = Get_Ini("System Config", "Str_Lane1_Reg", "정상")
    Glo_Str_Guest(0) = Get_Ini("System Config", "Str_Lane1_Guest", "미등록")
    Glo_Str_NoRec(0) = Get_Ini("System Config", "Str_Lane1_NoRec", "미인식")
    Glo_Str_BlackList(0) = Get_Ini("System Config", "Str_Lane1_BlackList", "출입제한")
    Glo_Str_Taxi(0) = Get_Ini("System Config", "Str_Lane1_Taxi", "영업차량")
    Glo_Str_Day(0) = Get_Ini("System Config", "Str_Lane1_Day", "요일제위반")
    Glo_Str_RegExpDate(0) = Get_Ini("System Config", "Str_Lane1_RegExpDate", "기간만료")
    Glo_Str_GuestRegCar(0) = Get_Ini("System Config", "Str_Lane1_GuestRegCar", "방문예약")
    Glo_Str_GuestRegCarExpDate(0) = Get_Ini("System Config", "Str_Lane1_GuestRegCarExpDate", "방문예약만료")
    Glo_Str_Reg(1) = Get_Ini("System Config", "Str_Lane2_Reg", "정상")
    Glo_Str_Guest(1) = Get_Ini("System Config", "Str_Lane2_Guest", "미등록")
    Glo_Str_NoRec(1) = Get_Ini("System Config", "Str_Lane2_NoRec", "미인식")
    Glo_Str_BlackList(1) = Get_Ini("System Config", "Str_Lane2_BlackList", "출입제한")
    Glo_Str_Taxi(1) = Get_Ini("System Config", "Str_Lane2_Taxi", "영업차량")
    Glo_Str_Day(1) = Get_Ini("System Config", "Str_Lane2_Day", "요일제위반")
    Glo_Str_RegExpDate(1) = Get_Ini("System Config", "Str_Lane2_RegExpDate", "기간만료")
    Glo_Str_GuestRegCar(1) = Get_Ini("System Config", "Str_Lane2_GuestRegCar", "방문예약")
    Glo_Str_GuestRegCarExpDate(1) = Get_Ini("System Config", "Str_Lane2_GuestRegCarExpDate", "방문예약만료")
    Glo_Str_Reg(2) = Get_Ini("System Config", "Str_Lane3_Reg", "정상")
    Glo_Str_Guest(2) = Get_Ini("System Config", "Str_Lane3_Guest", "미등록")
    Glo_Str_NoRec(2) = Get_Ini("System Config", "Str_Lane3_NoRec", "미인식")
    Glo_Str_BlackList(2) = Get_Ini("System Config", "Str_Lane3_BlackList", "출입제한")
    Glo_Str_Taxi(2) = Get_Ini("System Config", "Str_Lane3_Taxi", "영업차량")
    Glo_Str_Day(2) = Get_Ini("System Config", "Str_Lane3_Day", "요일제위반")
    Glo_Str_RegExpDate(2) = Get_Ini("System Config", "Str_Lane3_RegExpDate", "기간만료")
    Glo_Str_GuestRegCar(2) = Get_Ini("System Config", "Str_Lane3_GuestRegCar", "방문예약")
    Glo_Str_GuestRegCarExpDate(2) = Get_Ini("System Config", "Str_Lane3_GuestRegCarExpDate", "방문예약만료")
    Glo_Str_Reg(3) = Get_Ini("System Config", "Str_Lane4_Reg", "정상")
    Glo_Str_Guest(3) = Get_Ini("System Config", "Str_Lane4_Guest", "미등록")
    Glo_Str_NoRec(3) = Get_Ini("System Config", "Str_Lane4_NoRec", "미인식")
    Glo_Str_BlackList(3) = Get_Ini("System Config", "Str_Lane4_BlackList", "출입제한")
    Glo_Str_Taxi(3) = Get_Ini("System Config", "Str_Lane4_Taxi", "영업차량")
    Glo_Str_Day(3) = Get_Ini("System Config", "Str_Lane4_Day", "요일제위반")
    Glo_Str_RegExpDate(3) = Get_Ini("System Config", "Str_Lane4_RegExpDate", "기간만료")
    Glo_Str_GuestRegCar(3) = Get_Ini("System Config", "Str_Lane4_GuestRegCar", "방문예약")
    Glo_Str_GuestRegCarExpDate(3) = Get_Ini("System Config", "Str_Lane4_GuestRegCarExpDate", "방문예약만료")
    Glo_Str_Reg(4) = Get_Ini("System Config", "Str_Lane5_Reg", "정상")
    Glo_Str_Guest(4) = Get_Ini("System Config", "Str_Lane5_Guest", "미등록")
    Glo_Str_NoRec(4) = Get_Ini("System Config", "Str_Lane5_NoRec", "미인식")
    Glo_Str_BlackList(4) = Get_Ini("System Config", "Str_Lane5_BlackList", "출입제한")
    Glo_Str_Taxi(4) = Get_Ini("System Config", "Str_Lane5_Taxi", "영업차량")
    Glo_Str_Day(4) = Get_Ini("System Config", "Str_Lane5_Day", "요일제위반")
    Glo_Str_RegExpDate(4) = Get_Ini("System Config", "Str_Lane5_RegExpDate", "기간만료")
    Glo_Str_GuestRegCar(4) = Get_Ini("System Config", "Str_Lane5_GuestRegCar", "방문예약")
    Glo_Str_GuestRegCarExpDate(4) = Get_Ini("System Config", "Str_Lane5_GuestRegCarExpDate", "방문예약만료")
    Glo_Str_Reg(5) = Get_Ini("System Config", "Str_Lane6_Reg", "정상")
    Glo_Str_Guest(5) = Get_Ini("System Config", "Str_Lane6_Guest", "미등록")
    Glo_Str_NoRec(5) = Get_Ini("System Config", "Str_Lane6_NoRec", "미인식")
    Glo_Str_BlackList(5) = Get_Ini("System Config", "Str_Lane6_BlackList", "출입제한")
    Glo_Str_Taxi(5) = Get_Ini("System Config", "Str_Lane6_Taxi", "영업차량")
    Glo_Str_Day(5) = Get_Ini("System Config", "Str_Lane6_Day", "요일제위반")
    Glo_Str_RegExpDate(5) = Get_Ini("System Config", "Str_Lane6_RegExpDate", "기간만료")
    Glo_Str_GuestRegCar(5) = Get_Ini("System Config", "Str_Lane6_GuestRegCar", "방문예약")
    Glo_Str_GuestRegCarExpDate(5) = Get_Ini("System Config", "Str_Lane6_GuestRegCarExpDate", "방문예약만료")
    
    Glo_SNDFILE_Reg(0) = Get_Ini("System Config", "SNDFILE_Lane1_Reg", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Guest(0) = Get_Ini("System Config", "SNDFILE_Lane1_Guest", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_NoRec(0) = Get_Ini("System Config", "SNDFILE_Lane1_NoRec", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_BlackList(0) = Get_Ini("System Config", "SNDFILE_Lane1_BlackList", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Taxi(0) = Get_Ini("System Config", "SNDFILE_Lane1_Taxi", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Day(0) = Get_Ini("System Config", "SNDFILE_Lane1_Day", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_RegExpDate(0) = Get_Ini("System Config", "SNDFILE_Lane1_RegExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCar(0) = Get_Ini("System Config", "SNDFILE_Lane1_GuestRegCar", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCarExpDate(0) = Get_Ini("System Config", "SNDFILE_Lane1_GuestRegCarExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Reg(1) = Get_Ini("System Config", "SNDFILE_Lane2_Reg", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Guest(1) = Get_Ini("System Config", "SNDFILE_Lane2_Guest", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_NoRec(1) = Get_Ini("System Config", "SNDFILE_Lane2_NoRec", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_BlackList(1) = Get_Ini("System Config", "SNDFILE_Lane2_BlackList", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Taxi(1) = Get_Ini("System Config", "SNDFILE_Lane2_Taxi", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Day(1) = Get_Ini("System Config", "SNDFILE_Lane2_Day", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_RegExpDate(1) = Get_Ini("System Config", "SNDFILE_Lane2_RegExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCar(1) = Get_Ini("System Config", "SNDFILE_Lane2_GuestRegCar", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCarExpDate(1) = Get_Ini("System Config", "SNDFILE_Lane2_GuestRegCarExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Reg(2) = Get_Ini("System Config", "SNDFILE_Lane3_Reg", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Guest(2) = Get_Ini("System Config", "SNDFILE_Lane3_Guest", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_NoRec(2) = Get_Ini("System Config", "SNDFILE_Lane3_NoRec", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_BlackList(2) = Get_Ini("System Config", "SNDFILE_Lane3_BlackList", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Taxi(2) = Get_Ini("System Config", "SNDFILE_Lane3_Taxi", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Day(2) = Get_Ini("System Config", "SNDFILE_Lane3_Day", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_RegExpDate(2) = Get_Ini("System Config", "SNDFILE_Lane3_RegExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCar(2) = Get_Ini("System Config", "SNDFILE_Lane3_GuestRegCar", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCarExpDate(2) = Get_Ini("System Config", "SNDFILE_Lane3_GuestRegCarExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Reg(3) = Get_Ini("System Config", "SNDFILE_Lane4_Reg", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Guest(3) = Get_Ini("System Config", "SNDFILE_Lane4_Guest", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_NoRec(3) = Get_Ini("System Config", "SNDFILE_Lane4_NoRec", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_BlackList(3) = Get_Ini("System Config", "SNDFILE_Lane4_BlackList", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Taxi(3) = Get_Ini("System Config", "SNDFILE_Lane4_Taxi", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Day(3) = Get_Ini("System Config", "SNDFILE_Lane4_Day", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_RegExpDate(3) = Get_Ini("System Config", "SNDFILE_Lane4_RegExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCar(3) = Get_Ini("System Config", "SNDFILE_Lane4_GuestRegCar", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCarExpDate(3) = Get_Ini("System Config", "SNDFILE_Lane4_GuestRegCarExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Reg(4) = Get_Ini("System Config", "SNDFILE_Lane5_Reg", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Guest(4) = Get_Ini("System Config", "SNDFILE_Lane5_Guest", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_NoRec(4) = Get_Ini("System Config", "SNDFILE_Lane5_NoRec", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_BlackList(4) = Get_Ini("System Config", "SNDFILE_Lane5_BlackList", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Taxi(4) = Get_Ini("System Config", "SNDFILE_Lane5_Taxi", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Day(4) = Get_Ini("System Config", "SNDFILE_Lane5_Day", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_RegExpDate(4) = Get_Ini("System Config", "SNDFILE_Lane5_RegExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCar(4) = Get_Ini("System Config", "SNDFILE_Lane5_GuestRegCar", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCarExpDate(4) = Get_Ini("System Config", "SNDFILE_Lane5_GuestRegCarExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Reg(5) = Get_Ini("System Config", "SNDFILE_Lane6_Reg", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Guest(5) = Get_Ini("System Config", "SNDFILE_Lane6_Guest", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_NoRec(5) = Get_Ini("System Config", "SNDFILE_Lane6_NoRec", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_BlackList(5) = Get_Ini("System Config", "SNDFILE_Lane6_BlackList", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Taxi(5) = Get_Ini("System Config", "SNDFILE_Lane6_Taxi", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_Day(5) = Get_Ini("System Config", "SNDFILE_Lane6_Day", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_RegExpDate(5) = Get_Ini("System Config", "SNDFILE_Lane6_RegExpDate", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCar(5) = Get_Ini("System Config", "SNDFILE_Lane6_GuestRegCar", App.Path & "\sound\Bell.wav")
    Glo_SNDFILE_GuestRegCarExpDate(5) = Get_Ini("System Config", "SNDFILE_Lane6_GuestRegCarExpDate", App.Path & "\sound\Bell.wav")
    
    

    
    
    
    
    
    Glo_INOUT_USING_DATE = Val(Get_Ini("System Config", "INOUT_USING_DATE", "99"))
    Glo_Screen_No = Val(Get_Ini("System Config", "Screen_No", "4"))
    
    '0 : OnlyTwo 1: OnlyTwoIN
    'HostType = Val(Get_Ini("System Config", "HostType", "0"))
    
    Glo_EndDate = Val(Get_Ini("System Config", "END_Date", "99"))
    AdoConn_Str = Get_Ini("System Config", "Conn_Str", "")
    
    
    sip_pos = InStr(UCase(AdoConn_Str), "SERVER=") + Len("SERVER=")
    eip_pos = InStr(UCase(AdoConn_Str), "DATABASE=")
    DB_Server_IP = Mid(AdoConn_Str, sip_pos, eip_pos - sip_pos - 1)
    DB_Server_Port = 3306
    

    '출구무인정산기
    Glo_ApsYN = Get_Ini("System Config", "APS_YN", "N")
    Glo_Aps_IP = Get_Ini("System Config", "APS_IP", Glo_Aps_IP)
    Glo_Aps_PORT = Val(Get_Ini("System Config", "APS_Port", 5889))
    Glo_APSCMD_Port = Val(Get_Ini("System Config", "CMD_Port", 5888))
    Glo_APSCmdR_Port = Val(Get_Ini("System Config", "CmdR_Port", 5888))
    
    '사전무인정산기
    Glo_PreApsYN = Get_Ini("System Config", "PreAPS_YN", "N")
    Glo_Grace_Time = Val(Get_Ini("System Config", "GRACE_TIME", "0"))
    Glo_Return_Time = Val(Get_Ini("System Config", "RETURN_TIME", "0"))
    
    
    
    
    Glo_User_Type = Get_Ini("System Config", "User_Type", "구분1/구분2")
    'Glo_NoRec = Get_Ini("System Config", "NoRecOpen_YN", "N")
    
    Glo_NoRecOpen1_YN = Get_Ini("System Config", "NoRecOpen1_YN", "N")
    Glo_NoRecOpen2_YN = Get_Ini("System Config", "NoRecOpen2_YN", "N")
    Glo_NoRecOpen3_YN = Get_Ini("System Config", "NoRecOpen3_YN", "N")
    Glo_NoRecOpen4_YN = Get_Ini("System Config", "NoRecOpen4_YN", "N")
    Glo_NoRecOpen5_YN = Get_Ini("System Config", "NoRecOpen5_YN", "N")
    Glo_NoRecOpen6_YN = Get_Ini("System Config", "NoRecOpen6_YN", "N")
    
    'Glo_BlackList = Get_Ini("System Config", "Glo_BlackList", "N")
    Glo_BlackList1_YN = Get_Ini("System Config", "BlackList1_YN", "N")
    Glo_BlackList2_YN = Get_Ini("System Config", "BlackList2_YN", "N")
    Glo_BlackList3_YN = Get_Ini("System Config", "BlackList3_YN", "N")
    Glo_BlackList4_YN = Get_Ini("System Config", "BlackList4_YN", "N")
    Glo_BlackList5_YN = Get_Ini("System Config", "BlackList5_YN", "N")
    Glo_BlackList6_YN = Get_Ini("System Config", "BlackList6_YN", "N")

    
'    '세대통보 정보
'    'HomeAlarm_Mode = Val(Get_Ini("System Config", "HomeAlarm", "0"))
'    HomeSvr_IP = Trim(Get_Ini("System Config", "HomeSvr_IP", ""))
'    HomeSvr_Port = Get_Ini("System Config", "HomeSvr_Port", "")
'    Homesvr_ID = Trim(Get_Ini("System Config", "HomeSvr_ID", ""))
'    Homesvr_PW = Trim(Get_Ini("System Config", "HomeSvr_PW", ""))
    
    '프리패스 기능
    Glo_FreePassLane1_YN = Get_Ini("System Config", "FreePassLane1_YN", "N")
    Glo_FreePassLane2_YN = Get_Ini("System Config", "FreePassLane2_YN", "N")
    Glo_FreePassLane3_YN = Get_Ini("System Config", "FreePassLane3_YN", "N")
    Glo_FreePassLane4_YN = Get_Ini("System Config", "FreePassLane4_YN", "N")
    Glo_FreePassLane5_YN = Get_Ini("System Config", "FreePassLane5_YN", "N")
    Glo_FreePassLane6_YN = Get_Ini("System Config", "FreePassLane6_YN", "N")
    
    '한글 오인식 필터링 사용여부
    MissMatch_YN = Get_Ini("System Config", "MissMatch_YN", "N")
    MissMatch_HomeNet_YN = Get_Ini("System Config", "MissMatch_HomeNet_YN", "N")
    
    '요일제, 부제 사용여부
    Glo_WEEK_YN = Get_Ini("System Config", "WEEK_YN", "N")
    Glo_ROTATION = Get_Ini("System Config", "ROTATION", "미적용")
    
    'GyeYoung 세대통보
    AdoHome_Str = Get_Ini("System Config", "Home_Str", "")
    
    'LPR용 서버 TCP포트
    Server_Port = 10100
    
    LANE1_YN = Get_Ini("System Config", "LANE1_YN", "N")
    LANE1_Name = Get_Ini("System Config", "LANE1_Name", "입구")
    LANE1_Inout = Get_Ini("System Config", "LANE1_Inout", "입구")
    LANE1_LPRIP = Get_Ini("System Config", "LANE1_LPRIP", "192.168.0.211")
    LANE1_LPRPort = 10101
    LANE1_DeviceIP = Get_Ini("System Config", "LANE1_DeviceIP", "192.168.0.221")
    LANE1_DispIP = Get_Ini("System Config", "LANE1_DispIP", "192.168.0.211")
    LANE1_DispPort = 1000
    LANE1_RelayPort = 1100
    
    LANE1_Disp1Msg = Get_Ini("System Config", "LANE1_Disp1Msg", "UP String")
    LANE1_Disp2Msg = Get_Ini("System Config", "LANE1_Disp2Msg", "DOWN String")
    LANE1_Disp1Color = Get_Ini("System Config", "LANE1_Disp1Color", "1")
    LANE1_Disp2Color = Get_Ini("System Config", "LANE1_Disp2Color", "1")
    LANE1_DispSpeed = Get_Ini("System Config", "LANE_DispSpeed", "2")
    Glo_LANE_DISP_NML_SHIFT(0) = Val(Get_Ini("System Config", "LANE1_DispShift", "6"))
        
    LANE2_YN = Get_Ini("System Config", "LANE2_YN", "N")
    LANE2_Name = Get_Ini("System Config", "LANE2_Name", "출구")
    LANE2_Inout = Get_Ini("System Config", "LANE2_Inout", "입구")
    LANE2_LPRIP = Get_Ini("System Config", "LANE2_LPRIP", "192.168.0.212")
    LANE2_LPRPort = 10102
    LANE2_DeviceIP = Get_Ini("System Config", "LANE2_DeviceIP", "192.168.0.222")
    LANE2_DispIP = Get_Ini("System Config", "LANE2_DispIP", "192.168.0.212")
    LANE2_DispPort = 1000
    LANE2_RelayPort = 1100
    LANE2_Disp1Msg = Get_Ini("System Config", "LANE2_Disp1Msg", "UP String")
    LANE2_Disp2Msg = Get_Ini("System Config", "LANE2_Disp2Msg", "DOWN String")
    LANE2_Disp1Color = Get_Ini("System Config", "LANE2_Disp1Color", "1")
    LANE2_Disp2Color = Get_Ini("System Config", "LANE2_Disp2Color", "1")
    LANE2_DispSpeed = Get_Ini("System Config", "LANE2_DispSpeed", "2")
    Glo_LANE_DISP_NML_SHIFT(1) = Val(Get_Ini("System Config", "LANE2_DispShift", "6"))
    
    LANE3_YN = Get_Ini("System Config", "LANE3_YN", "N")
    LANE3_Name = Get_Ini("System Config", "LANE3_Name", "입구")
    LANE3_Inout = Get_Ini("System Config", "LANE3_Inout", "출구")
    LANE3_LPRIP = Get_Ini("System Config", "LANE3_LPRIP", "192.168.0.213")
    LANE3_LPRPort = 10103
    LANE3_DeviceIP = Get_Ini("System Config", "LANE3_DeviceIP", "192.168.0.223")
    LANE3_DispIP = Get_Ini("System Config", "LANE3_DispIP", "192.168.0.213")
    LANE3_DispPort = 1000
    LANE3_RelayPort = 1100
    LANE3_Disp1Msg = Get_Ini("System Config", "LANE3_Disp1Msg", "UP String")
    LANE3_Disp2Msg = Get_Ini("System Config", "LANE3_Disp2Msg", "DOWN String")
    LANE3_Disp1Color = Get_Ini("System Config", "LANE3_Disp1Color", "1")
    LANE3_Disp2Color = Get_Ini("System Config", "LANE3_Disp2Color", "1")
    LANE3_DispSpeed = Get_Ini("System Config", "LANE3_DispSpeed", "2")
    Glo_LANE_DISP_NML_SHIFT(2) = Val(Get_Ini("System Config", "LANE3_DispShift", "6"))
    
    LANE4_YN = Get_Ini("System Config", "LANE4_YN", "N")
    LANE4_Name = Get_Ini("System Config", "LANE4_Name", "출구")
    LANE4_Inout = Get_Ini("System Config", "LANE4_Inout", "출구")
    LANE4_LPRIP = Get_Ini("System Config", "LANE4_LPRIP", "192.168.0.214")
    LANE4_LPRPort = 10104
    LANE4_DeviceIP = Get_Ini("System Config", "LANE4_DeviceIP", "192.168.0.224")
    LANE4_DispIP = Get_Ini("System Config", "LANE4_DispIP", "192.168.0.214")
    LANE4_DispPort = 1000
    LANE4_RelayPort = 1100
    LANE4_Disp1Msg = Get_Ini("System Config", "LANE4_Disp1Msg", "UP String")
    LANE4_Disp2Msg = Get_Ini("System Config", "LANE4_Disp2Msg", "DOWN String")
    LANE4_Disp1Color = Get_Ini("System Config", "LANE4_Disp1Color", "1")
    LANE4_Disp2Color = Get_Ini("System Config", "LANE4_Disp2Color", "1")
    LANE4_DispSpeed = Get_Ini("System Config", "LANE4_DispSpeed", "2")
    Glo_LANE_DISP_NML_SHIFT(3) = Val(Get_Ini("System Config", "LANE4_DispShift", "6"))
    
    LANE5_YN = Get_Ini("System Config", "LANE5_YN", "N")
    LANE5_Name = Get_Ini("System Config", "LANE5_Name", "출구")
    LANE5_Inout = Get_Ini("System Config", "LANE5_Inout", "출구")
    LANE5_LPRIP = Get_Ini("System Config", "LANE5_LPRIP", "192.168.0.214")
    LANE5_LPRPort = 10105
    LANE5_DeviceIP = Get_Ini("System Config", "LANE5_DeviceIP", "192.168.0.224")
    LANE5_DispIP = Get_Ini("System Config", "LANE5_DispIP", "192.168.0.215")
    LANE5_DispPort = 1000
    LANE5_RelayPort = 1100
    LANE5_Disp1Msg = Get_Ini("System Config", "LANE5_Disp1Msg", "UP String")
    LANE5_Disp2Msg = Get_Ini("System Config", "LANE5_Disp2Msg", "DOWN String")
    LANE5_Disp1Color = Get_Ini("System Config", "LANE5_Disp1Color", "1")
    LANE5_Disp2Color = Get_Ini("System Config", "LANE5_Disp2Color", "1")
    LANE5_DispSpeed = Get_Ini("System Config", "LANE5_DispSpeed", "2")
    Glo_LANE_DISP_NML_SHIFT(4) = Val(Get_Ini("System Config", "LANE5_DispShift", "6"))
    
    LANE6_YN = Get_Ini("System Config", "LANE6_YN", "N")
    LANE6_Name = Get_Ini("System Config", "LANE6_Name", "출구")
    LANE6_Inout = Get_Ini("System Config", "LANE6_Inout", "출구")
    LANE6_LPRIP = Get_Ini("System Config", "LANE6_LPRIP", "192.168.0.216")
    LANE6_LPRPort = 10106
    LANE6_DeviceIP = Get_Ini("System Config", "LANE6_DeviceIP", "192.168.0.226")
    LANE6_DispIP = Get_Ini("System Config", "LANE6_DispIP", "192.168.0.216")
    LANE6_DispPort = 1000
    LANE6_RelayPort = 1100
    LANE6_Disp1Msg = Get_Ini("System Config", "LANE6_Disp1Msg", "UP String")
    LANE6_Disp2Msg = Get_Ini("System Config", "LANE6_Disp2Msg", "DOWN String")
    LANE6_Disp1Color = Get_Ini("System Config", "LANE6_Disp1Color", "1")
    LANE6_Disp2Color = Get_Ini("System Config", "LANE6_Disp2Color", "1")
    LANE6_DispSpeed = Get_Ini("System Config", "LANE6_DispSpeed", "2")
    Glo_LANE_DISP_NML_SHIFT(5) = Val(Get_Ini("System Config", "LANE6_DispShift", "6"))
    
    
    
    LANE1_LPRMode = Get_Ini("System Config", "LPRMode", "0")
    LANE2_LPRMode = LANE1_LPRMode
    LANE3_LPRMode = LANE1_LPRMode
    LANE4_LPRMode = LANE1_LPRMode
    LANE5_LPRMode = LANE1_LPRMode
    LANE6_LPRMode = LANE1_LPRMode
    
    
    LANE1_DeviceMode = Get_Ini("System Config", "DeviceMode", "0")
    LANE2_DeviceMode = LANE1_DeviceMode
    LANE3_DeviceMode = LANE1_DeviceMode
    LANE4_DeviceMode = LANE1_DeviceMode
    LANE5_DeviceMode = LANE1_DeviceMode
    LANE6_DeviceMode = LANE1_DeviceMode
    
    LANE1_DisplayMode = Get_Ini("System Config", "DisplayMode", "0")
    LANE2_DisplayMode = LANE1_DisplayMode
    LANE3_DisplayMode = LANE1_DisplayMode
    LANE4_DisplayMode = LANE1_DisplayMode
    LANE5_DisplayMode = LANE1_DisplayMode
    LANE6_DisplayMode = LANE1_DisplayMode
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Glo_GUEST_LANE1_YN = Get_Ini("System Config", "GUEST1_YN", "N")
    Glo_GUEST_LANE2_YN = Get_Ini("System Config", "GUEST2_YN", "N")
    Glo_GUEST_LANE3_YN = Get_Ini("System Config", "GUEST3_YN", "N")
    Glo_GUEST_LANE4_YN = Get_Ini("System Config", "GUEST4_YN", "N")
    Glo_GUEST_LANE5_YN = Get_Ini("System Config", "GUEST5_YN", "N")
    Glo_GUEST_LANE6_YN = Get_Ini("System Config", "GUEST6_YN", "N")
    
    Glo_Guest_Print_Model(0) = Get_Ini("System Config", "GUEST1_PRINT_MODEL", "NONE")
    Glo_Guest_Print_Model(1) = Get_Ini("System Config", "GUEST2_PRINT_MODEL", "NONE")
    Glo_Guest_Print_Model(2) = Get_Ini("System Config", "GUEST3_PRINT_MODEL", "NONE")
    Glo_Guest_Print_Model(3) = Get_Ini("System Config", "GUEST4_PRINT_MODEL", "NONE")
    Glo_Guest_Print_Model(4) = Get_Ini("System Config", "GUEST5_PRINT_MODEL", "NONE")
    Glo_Guest_Print_Model(5) = Get_Ini("System Config", "GUEST6_PRINT_MODEL", "NONE")
    
    Glo_Guest_Print_Port(0) = Get_Ini("System Config", "GUEST1_PRINT_PORT", "COM1")
    Glo_Guest_Print_Port(1) = Get_Ini("System Config", "GUEST2_PRINT_PORT", "COM2")
    Glo_Guest_Print_Port(2) = Get_Ini("System Config", "GUEST3_PRINT_PORT", "COM3")
    Glo_Guest_Print_Port(3) = Get_Ini("System Config", "GUEST4_PRINT_PORT", "COM4")
    Glo_Guest_Print_Port(4) = Get_Ini("System Config", "GUEST5_PRINT_PORT", "COM5")
    Glo_Guest_Print_Port(5) = Get_Ini("System Config", "GUEST6_PRINT_PORT", "COM6")
    
    Glo_Guest_Print_Port(0) = Get_Ini("System Config", "GUEST1_PRINT_PORT", "COM1")
    Glo_Guest_Print_Port(1) = Get_Ini("System Config", "GUEST2_PRINT_PORT", "COM2")
    Glo_Guest_Print_Port(2) = Get_Ini("System Config", "GUEST3_PRINT_PORT", "COM3")
    Glo_Guest_Print_Port(3) = Get_Ini("System Config", "GUEST4_PRINT_PORT", "COM4")
    Glo_Guest_Print_Port(4) = Get_Ini("System Config", "GUEST5_PRINT_PORT", "COM5")
    Glo_Guest_Print_Port(5) = Get_Ini("System Config", "GUEST6_PRINT_PORT", "COM6")
    
    Glo_Guest_Gate_OpenDelay(0) = CSng(Get_Ini("System Config", "GUEST1_GATE_OPENDELAY_TIME", 0))
    Glo_Guest_Gate_OpenDelay(1) = CSng(Get_Ini("System Config", "GUEST2_GATE_OPENDELAY_TIME", 0))
    Glo_Guest_Gate_OpenDelay(2) = CSng(Get_Ini("System Config", "GUEST3_GATE_OPENDELAY_TIME", 0))
    Glo_Guest_Gate_OpenDelay(3) = CSng(Get_Ini("System Config", "GUEST4_GATE_OPENDELAY_TIME", 0))
    Glo_Guest_Gate_OpenDelay(4) = CSng(Get_Ini("System Config", "GUEST5_GATE_OPENDELAY_TIME", 0))
    Glo_Guest_Gate_OpenDelay(5) = CSng(Get_Ini("System Config", "GUEST6_GATE_OPENDELAY_TIME", 0))
    
    Glo_Receipt_Paper_Cut = Get_Ini("System Config", "RECEIPT_PAPER_CUT", "0")
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub


'전광판 색상 코드
Public Function GetDispColorData(sColor As String)
    Dim upColor As Byte
    Dim downColor As Byte
    
    Select Case sColor
        Case "적"
            GetDispColorData = enumDIS_COLORs.eRED
        Case "황"
            GetDispColorData = enumDIS_COLORs.eYellow
        Case "녹"
            GetDispColorData = enumDIS_COLORs.eGreen
        Case "파"
            GetDispColorData = enumDIS_COLORs.eBLUE
        Case "자"
            GetDispColorData = enumDIS_COLORs.eWINE
        Case "하"
            GetDispColorData = enumDIS_COLORs.eSKY
        Case "백"
            GetDispColorData = enumDIS_COLORs.eWHITE
        Case Else
            GetDispColorData = enumDIS_COLORs.eGreen
    End Select

End Function
Public Sub LoadDBConfig()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tb_config ", adoConn

    Do While Not (rs.EOF)

        GetDispColorData (rs!Content)
        If (rs!name = "LANE1_Disp1EmgColorReg") Then
            Glo_Disp1_Reg(0) = GetDispColorData(rs!Content):
        ElseIf (rs!name = "LANE1_Disp2EmgColorReg") Then Glo_Disp2_Reg(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorGuest") Then Glo_Disp1_Guest(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorGuest") Then Glo_Disp2_Guest(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorNoRec") Then Glo_Disp1_NoRec(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorNoRec") Then Glo_Disp2_NoRec(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorBKList") Then Glo_Disp1_BlackList(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorBKList") Then Glo_Disp2_BlackList(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorTaxi") Then Glo_Disp1_Taxi(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorTaxi") Then Glo_Disp2_Taxi(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorDay") Then Glo_Disp1_Day(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorDay") Then Glo_Disp2_Day(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorRegExpDate") Then Glo_Disp1_RegExpDate(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorRegExpDate") Then Glo_Disp2_RegExpDate(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorGuestRegCar") Then Glo_Disp1_GuestRegCar(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorGuestRegCar") Then Glo_Disp2_GuestRegCar(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp1EmgColorGuestRegCarExpDate") Then Glo_Disp1_GuestRegCarExpDate(0) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE1_Disp2EmgColorGuestRegCarExpDate") Then Glo_Disp2_GuestRegCarExpDate(0) = GetDispColorData(rs!Content)

        ElseIf (rs!name = "LANE2_Disp1EmgColorReg") Then Glo_Disp1_Reg(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorReg") Then Glo_Disp2_Reg(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorGuest") Then Glo_Disp1_Guest(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorGuest") Then Glo_Disp2_Guest(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorNoRec") Then Glo_Disp1_NoRec(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorNoRec") Then Glo_Disp2_NoRec(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorBKList") Then Glo_Disp1_BlackList(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorBKList") Then Glo_Disp2_BlackList(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorTaxi") Then Glo_Disp1_Taxi(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorTaxi") Then Glo_Disp2_Taxi(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorDay") Then Glo_Disp1_Day(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorDay") Then Glo_Disp2_Day(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorRegExpDate") Then Glo_Disp1_RegExpDate(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorRegExpDate") Then Glo_Disp2_RegExpDate(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorGuestRegCar") Then Glo_Disp1_GuestRegCar(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorGuestRegCar") Then Glo_Disp2_GuestRegCar(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp1EmgColorGuestRegCarExpDate") Then Glo_Disp1_GuestRegCarExpDate(1) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE2_Disp2EmgColorGuestRegCarExpDate") Then Glo_Disp2_GuestRegCarExpDate(1) = GetDispColorData(rs!Content)
        
        ElseIf (rs!name = "LANE3_Disp1EmgColorReg") Then Glo_Disp1_Reg(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorReg") Then Glo_Disp2_Reg(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorGuest") Then Glo_Disp1_Guest(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorGuest") Then Glo_Disp2_Guest(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorNoRec") Then Glo_Disp1_NoRec(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorNoRec") Then Glo_Disp2_NoRec(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorBKList") Then Glo_Disp1_BlackList(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorBKList") Then Glo_Disp2_BlackList(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorTaxi") Then Glo_Disp1_Taxi(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorTaxi") Then Glo_Disp2_Taxi(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorDay") Then Glo_Disp1_Day(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorDay") Then Glo_Disp2_Day(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorRegExpDate") Then Glo_Disp1_RegExpDate(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorRegExpDate") Then Glo_Disp2_RegExpDate(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorGuestRegCar") Then Glo_Disp1_GuestRegCar(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorGuestRegCar") Then Glo_Disp2_GuestRegCar(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp1EmgColorGuestRegCarExpDate") Then Glo_Disp1_GuestRegCarExpDate(2) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE3_Disp2EmgColorGuestRegCarExpDate") Then Glo_Disp2_GuestRegCarExpDate(2) = GetDispColorData(rs!Content)
        
        ElseIf (rs!name = "LANE4_Disp1EmgColorReg") Then Glo_Disp1_Reg(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorReg") Then Glo_Disp2_Reg(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorGuest") Then Glo_Disp1_Guest(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorGuest") Then Glo_Disp2_Guest(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorNoRec") Then Glo_Disp1_NoRec(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorNoRec") Then Glo_Disp2_NoRec(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorBKList") Then Glo_Disp1_BlackList(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorBKList") Then Glo_Disp2_BlackList(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorTaxi") Then Glo_Disp1_Taxi(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorTaxi") Then Glo_Disp2_Taxi(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorDay") Then Glo_Disp1_Day(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorDay") Then Glo_Disp2_Day(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorRegExpDate") Then Glo_Disp1_RegExpDate(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorRegExpDate") Then Glo_Disp2_RegExpDate(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorGuestRegCar") Then Glo_Disp1_GuestRegCar(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorGuestRegCar") Then Glo_Disp2_GuestRegCar(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp1EmgColorGuestRegCarExpDate") Then Glo_Disp1_GuestRegCarExpDate(3) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE4_Disp2EmgColorGuestRegCarExpDate") Then Glo_Disp2_GuestRegCarExpDate(3) = GetDispColorData(rs!Content)
        
        ElseIf (rs!name = "LANE5_Disp1EmgColorReg") Then Glo_Disp1_Reg(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorReg") Then Glo_Disp2_Reg(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorGuest") Then Glo_Disp1_Guest(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorGuest") Then Glo_Disp2_Guest(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorNoRec") Then Glo_Disp1_NoRec(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorNoRec") Then Glo_Disp2_NoRec(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorBKList") Then Glo_Disp1_BlackList(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorBKList") Then Glo_Disp2_BlackList(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorTaxi") Then Glo_Disp1_Taxi(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorTaxi") Then Glo_Disp2_Taxi(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorDay") Then Glo_Disp1_Day(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorDay") Then Glo_Disp2_Day(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorRegExpDate") Then Glo_Disp1_RegExpDate(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorRegExpDate") Then Glo_Disp2_RegExpDate(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorGuestRegCar") Then Glo_Disp1_GuestRegCar(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorGuestRegCar") Then Glo_Disp2_GuestRegCar(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp1EmgColorGuestRegCarExpDate") Then Glo_Disp1_GuestRegCarExpDate(4) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE5_Disp2EmgColorGuestRegCarExpDate") Then Glo_Disp2_GuestRegCarExpDate(4) = GetDispColorData(rs!Content)
        
        ElseIf (rs!name = "LANE6_Disp1EmgColorReg") Then Glo_Disp1_Reg(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorReg") Then Glo_Disp2_Reg(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorGuest") Then Glo_Disp1_Guest(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorGuest") Then Glo_Disp2_Guest(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorNoRec") Then Glo_Disp1_NoRec(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorNoRec") Then Glo_Disp2_NoRec(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorBKList") Then Glo_Disp1_BlackList(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorBKList") Then Glo_Disp2_BlackList(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorTaxi") Then Glo_Disp1_Taxi(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorTaxi") Then Glo_Disp2_Taxi(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorDay") Then Glo_Disp1_Day(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorDay") Then Glo_Disp2_Day(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorRegExpDate") Then Glo_Disp1_RegExpDate(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorRegExpDate") Then Glo_Disp2_RegExpDate(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorGuestRegCar") Then Glo_Disp1_GuestRegCar(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorGuestRegCar") Then Glo_Disp2_GuestRegCar(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp1EmgColorGuestRegCarExpDate") Then Glo_Disp1_GuestRegCarExpDate(5) = GetDispColorData(rs!Content)
        ElseIf (rs!name = "LANE6_Disp2EmgColorGuestRegCarExpDate") Then Glo_Disp2_GuestRegCarExpDate(5) = GetDispColorData(rs!Content)
        
        ElseIf (rs!name = "Disp_Vertical_ToggleCount") Then Glo_Emerg_Vertical_ToggleCount = rs!Content
        ElseIf (rs!name = "Disp_Vertical_ToggleTime") Then Glo_Emerg_Vertical_ToggleTime = rs!Content

        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

