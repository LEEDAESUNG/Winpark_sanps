VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  '없음
   Caption         =   "ParkingManager™"
   ClientHeight    =   6015
   ClientLeft      =   9990
   ClientTop       =   5310
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   6015
   ScaleWidth      =   9000
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   1500
      TabIndex        =   8
      Top             =   3960
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   120
      Top             =   5400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "본 프로그램은 프로그램등록법에 의거 정식등록 되어있으므로, 무단복제하는 경우에는 저작권의 침해가 되므로 주의 바랍니다. "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   2670
      TabIndex        =   7
      Top             =   4785
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "ⓒ 2003-2017 All Right Reserved. See the patent and legal notice in the about box."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   780
      TabIndex        =   6
      Top             =   5310
      Width           =   7860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9615
      TabIndex        =   5
      Top             =   1995
      Width           =   1455
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   8775
   End
   Begin VB.Label lblLicenseTo 
      BackStyle       =   0  '투명
      Caption         =   "LicenseTo"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '투명
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   645
      TabIndex        =   2
      Top             =   1335
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "Parking Manager™  for LPR System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  '투명
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   645
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim tm_cnt As Integer

Private Sub Form_Load()
Dim Tmp_File As String
Dim i As Integer

On Error GoTo Err_p

If App.PrevInstance = True Then
    End
End If

Left = (Screen.width - width) / 2
Top = (Screen.height - height) / 2
lblVersion.Caption = "Version " & App.Major & " . " & App.Minor & " . " & App.Revision
'lblCompany.Caption = App.CompanyName
lblCompany.Caption = "주차관제시스템"
lblLicenseTo.Caption = App.FileDescription
IniFileName$ = App.Path & "\Winpark.ini"
Doc_Path_Name$ = App.Path & "\Doc\"
Db_Path_Name$ = App.Path & "\Data\"
Report_Path_Name$ = App.Path & "\Data\"

'전역변수 초기화

Call CFG_Init

    DB_Connect_F = True

    i = 0
    If (adoConn.State = adStateOpen) Then
        Call DataBaseClose(adoConn)
    End If

    Do While DataBaseOpen(adoConn) = False
        Call DataLogger("DB Connection Failure..!!")
        Call Delay_Time(1)
        i = i + 1
'''        If i > 3 Then
            DB_Connect_F = False
            Call MsgBox("DataBase IP 주소와 Name을 확인후 재실행하세요", vbCritical Or vbMsgBoxSetForeground, "DataBase 설정 에러")
            Exit Do
            'End
'''        End If
    Loop

    If (DB_Connect_F = True) Then
        
        Call DB_Table_Check
        'Call DB_CFG_Init("SOUND")
        
        DB_Rcv_LastTime = Timer
        If (adoConn.State = adStateOpen) Then
            adoConn.Execute "SET GLOBAL max_connect_errors=99999999"
        End If
        Call Time_Sync

        Call Check_SuperAdminAccount
    End If
    
    Call LoadDBConfig
    

'''    Call Get_MacAddress

    'Call Genuine(Key)
    
'If DataBaseOpen(adoConn) Then
'    FrmG4Mini.List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "Database Connect...!!", 0
'    adoConn.Execute "SET GLOBAL max_connect_errors=99999999"
'    Call Time_Sync
'Else
'    MsgBox "데이터베이스 접속 실패..!!       ", vbCritical, "  www.jawootek.com"
'    End
'End If

ProgressBar1.Max = 100
Timer1.Enabled = True

Exit Sub

Err_p:
'    DataLogger ("DB:" & DB_Connect_F)
End Sub

Private Sub Check_SuperAdminAccount()
    Dim rs As ADODB.Recordset
    Dim sQry As String
    Dim sPasswordEncode As String
    
    sQry = "SELECT * FROM tb_id WHERE GUBUN = '총괄관리자' "
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    If (rs.EOF) Then
        
        sPasswordEncode = EncodeNDE01("11111111", "www.jawootek.com") '암호화
        adoConn.Execute "INSERT INTO tb_id(ID, PASSWORD, GUBUN, MENU1, MENU2, MENU3, MENU4, MENU5, MENU6, MENU7, MENU8, MENU9, MENU10, REG_DATE ) VALUES ('11111111', '" & sPasswordEncode & "','총괄관리자','입출차조회','정기권관리','정기권이력','근무자관리','환경설정','','','','','','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('[ID생성]', 'HOST','[총괄관리자 ID 자동생성]',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        Call DataLogger("[ 총괄관리자 ID 자동생성] ")
    End If
    Set rs = Nothing

    Exit Sub

Err_p:
    Set rs = Nothing
End Sub
Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()

On Error Resume Next

If (tm_cnt <= 99) Then
    tm_cnt = tm_cnt + 1
    ProgressBar1.value = tm_cnt
Else
    Timer1.Enabled = False
    
    If Glo_Screen_No = 6 Then '6화면
        FrmG6_23.Show 0
    ElseIf Glo_Screen_No = 4 Then
'        Jung.Hide
        FrmG4Mini.Show 0    '4화면
    ElseIf Glo_Screen_No = 2 Then
'        FrmG4Mini.Hide
        Jung.Show 0         '2화면
    ElseIf Glo_Screen_No = 1 Then
        FrmG1.Show 0         '1화면
    End If
    
    Unload Me
End If

End Sub

Private Sub DB_Table_Check()
    Dim sFindTB_F As String

On Error GoTo Err_p


'''    sFindTBName = "tb_config"
'''    sFindTB_F = DB_Table_Find(sFindTBName)
'''
'''    If (sFindTB_F = "Y") Then
'''    Else
'''        Call DB_Table_Create(sFindTBName)
'''
'''
'''
'''    End If
    
    If (DB_Table_Find("tb_certify") = "Y") Then
    Else
        Call DB_Table_Create("tb_certify")
    End If
    
    Exit Sub

Err_p:
    
End Sub

Private Function DB_Table_Find(ByVal sFindTBName As String)
    Dim rs As Recordset
    Dim bQryResult As Boolean
    
    DB_Table_Find = "N"
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, "show tables", False)
    
    Do While Not (rs.EOF)
        'Debug.Print rs(0)
        If (rs(0) = sFindTBName) Then
            DB_Table_Find = "Y"
            Exit Do
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    
End Function


Private Function DB_Table_Create(ByVal sFindTBName As String)
    
    Dim qry As String
    Dim bQryResult As Boolean
    
    If (sFindTBName = "tb_config") Then
        'Qry = "CREATE TABLE `tb_config` (`Seq` int(10) unsigned NOT NULL AUTO_INCREMENT, `Title` varchar(32) DEFAULT NULL, `Name` varchar(32) DEFAULT NULL, `Content` varchar(256) DEFAULT NULL, `RegDate` varchar(32) DEFAULT NULL, PRIMARY KEY (`Seq`)) ENGINE=InnoDB AUTO_INCREMENT=20 DEFAULT CHARSET=euckr"
        qry = "CREATE TABLE `tb_config` (`Title` varchar(32) DEFAULT NULL,`Name` varchar(32) NOT NULL,`Content` varchar(256) DEFAULT NULL,`RegDate` varchar(32) DEFAULT NULL,PRIMARY KEY (`Name`)) ENGINE=InnoDB DEFAULT CHARSET=euckr"
        bQryResult = DataBaseQueryExec(adoConn, qry, False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','SOUND_YN','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane1_NoReg','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane1_NoRec','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane1_BlackList','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane2_NoReg','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane2_NoRec','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane2_BlackList','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane3_NoReg','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane3_NoRec','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane3_BlackList','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane4_NoReg','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane4_NoRec','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane4_BlackList','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane5_NoReg','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane5_NoRec','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane5_BlackList','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane6_NoReg','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane6_NoRec','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO " & sFindTBName & " (Title,Name,Content,RegDate) VALUES ('사운드','Lane6_BlackList','N', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') ", False)
        
    ElseIf (sFindTBName = "tb_certify") Then
    
        'jwt_sanps 전용 테이블
        qry = "CREATE TABLE `tb_certify` (  `SEQ` int(11) unsigned NOT NULL AUTO_INCREMENT,  `LOCKDATE` varchar(10) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `UNLOCKDATE` varchar(10) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `IP` varchar(32) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `MAC` varchar(17) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `HASHCODE` varchar(256) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `SITECODE` varchar(6) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `SITENAME` varchar(32) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL',  `MEMO` varchar(256) CHARACTER SET euckr COLLATE euckr_korean_ci DEFAULT NULL,  `C2DATE` datetime DEFAULT NULL',  PRIMARY KEY (`SEQ`)) ENGINE=MyISAM AUTO_INCREMENT=24 DEFAULT CHARSET=euckr"
        bQryResult = DataBaseQueryExec(adoConn, qry, False)
    End If
End Function




