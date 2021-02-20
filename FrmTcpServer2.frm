VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmTcpServer2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   10350
   ClientLeft      =   7065
   ClientTop       =   2880
   ClientWidth     =   19290
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   19290
   Begin VB.CheckBox chk_HomeNet_YN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3405
      TabIndex        =   112
      Top             =   2010
      Width           =   720
   End
   Begin VB.CheckBox chk_ParkFull_YN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6630
      TabIndex        =   129
      ToolTipText     =   "출구LPR 설치시에만 사용하세요. 만차 기능사용:만차시 전광판 안내문구 미표시 + ""차단기 자동열림"" 기능 중지됩니다."
      Top             =   2010
      Width           =   720
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10560
      TabIndex        =   140
      ToolTipText     =   "만차 기능설정시 전광판 안내문구가 나타나지 않습니다."
      Top             =   2010
      Width           =   720
   End
   Begin VB.CheckBox chk_GuestLogBackup_YN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   14250
      TabIndex        =   131
      Top             =   2010
      Width           =   720
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   18510
      TabIndex        =   154
      Top             =   2010
      Width           =   720
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 블랙리스트 자동등록 설정 *** 작업중 ***"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Left            =   15060
      TabIndex        =   155
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   163
         Text            =   "60"
         Top             =   1200
         Width           =   630
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   162
         ToolTipText     =   "출구LPR 설치시에만 사용가능합니다"
         Top             =   1260
         Width           =   720
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   975
         TabIndex        =   161
         Text            =   "60"
         Top             =   1200
         Width           =   630
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   158
         Text            =   "5"
         Top             =   495
         Width           =   630
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   157
         Top             =   555
         Width           =   720
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   975
         TabIndex        =   156
         Text            =   "3"
         Top             =   495
         Width           =   630
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "일 주차"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3315
         TabIndex        =   165
         Top             =   1260
         Width           =   660
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "개월 이내"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   10
         Left            =   1650
         TabIndex        =   164
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "회 입차"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3315
         TabIndex        =   160
         Top             =   555
         Width           =   660
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "일 이내"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   9
         Left            =   1650
         TabIndex        =   159
         Top             =   555
         Width           =   840
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 만차등 설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Left            =   7395
      TabIndex        =   141
      Top             =   2160
      Width           =   3900
      Begin VB.TextBox txt_ParkFullLight_FullRate 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   151
         Text            =   "100"
         Top             =   330
         Width           =   435
      End
      Begin VB.TextBox txt_ParkFullLight_BusyRate 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   148
         Text            =   "75"
         Top             =   750
         Width           =   435
      End
      Begin VB.TextBox txt_ParkFullLight_Empty 
         Height          =   315
         Left            =   1125
         TabIndex        =   147
         Text            =   "여유"
         Top             =   1155
         Width           =   630
      End
      Begin VB.TextBox txt_ParkFullLight_Busy 
         Height          =   315
         Left            =   1125
         TabIndex        =   145
         Text            =   "혼잡"
         Top             =   750
         Width           =   630
      End
      Begin VB.TextBox txt_ParkFullLight_Full 
         Height          =   315
         Left            =   1125
         TabIndex        =   142
         Text            =   "만차"
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   3525
         TabIndex        =   153
         Top             =   390
         Width           =   165
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "만차비율"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   2115
         TabIndex        =   152
         Top             =   390
         Width           =   780
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   3525
         TabIndex        =   150
         Top             =   810
         Width           =   165
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "혼잡비율"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   2115
         TabIndex        =   149
         Top             =   810
         Width           =   780
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "여유문구"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   146
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "만차문구"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   144
         Top             =   390
         Width           =   780
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "혼잡문구"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   143
         Top             =   810
         Width           =   780
      End
   End
   Begin VB.TextBox txt_CertifyKey 
      Height          =   435
      IMEMode         =   3  '사용 못함
      Left            =   3990
      PasswordChar    =   "*"
      TabIndex        =   138
      ToolTipText     =   "인증키 입력하세요"
      Top             =   690
      Width           =   1650
   End
   Begin VB.CommandButton cmd_Certify 
      Caption         =   "인증필요"
      Height          =   435
      Left            =   2730
      TabIndex        =   137
      ToolTipText     =   "인증키있을 경우 인증받으세요"
      Top             =   675
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "사운드 및 전광판 긴급문구"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   136
      Top             =   675
      Width           =   2370
   End
   Begin VB.Frame frmGuestLogBackup 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 방문차량 기록 보유기간 설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Left            =   11325
      TabIndex        =   132
      Top             =   2160
      Width           =   3690
      Begin VB.TextBox txt_GuestLogBackup 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2010
         TabIndex        =   133
         Text            =   "60"
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "개월"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2820
         TabIndex        =   135
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "보유기간 설정"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   240
         TabIndex        =   134
         Top             =   390
         Width           =   1395
      End
   End
   Begin VB.Frame frmHomeNet 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 세대통보 설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Left            =   30
      TabIndex        =   113
      Top             =   2160
      Width           =   4080
      Begin VB.TextBox Text_HomeNet_IP 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   119
         Text            =   "192.168.0.200"
         Top             =   630
         Width           =   1400
      End
      Begin VB.TextBox txt_Ho 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1245
         TabIndex        =   118
         Text            =   "101"
         Top             =   1080
         Width           =   630
      End
      Begin VB.TextBox txt_Dong 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   255
         TabIndex        =   117
         Text            =   "102"
         Top             =   1080
         Width           =   630
      End
      Begin VB.CommandButton cmd_HomeTest 
         Caption         =   "세대통보 테스트"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2370
         TabIndex        =   116
         Top             =   1080
         Width           =   1440
      End
      Begin VB.TextBox Text_HomeNet_Port 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   115
         Text            =   "18497"
         Top             =   630
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.ComboBox cmb_HomeNet 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmTcpServer2.frx":0000
         Left            =   2370
         List            =   "FrmTcpServer2.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   114
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "포트"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   123
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "아이피"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   122
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   930
         TabIndex        =   121
         Top             =   1170
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   120
         Top             =   1170
         Width           =   255
      End
   End
   Begin VB.Timer DB_Connect_Timer 
      Enabled         =   0   'False
      Left            =   21180
      Top             =   300
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   200
      Left            =   8760
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   200
      Left            =   8340
      Top             =   0
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Enabled         =   0   'False
      Height          =   210
      Index           =   5
      Left            =   18630
      TabIndex        =   78
      Top             =   4065
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Enabled         =   0   'False
      Height          =   210
      Index           =   4
      Left            =   15420
      TabIndex        =   94
      Top             =   4065
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE5 "
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Index           =   4
      Left            =   12900
      TabIndex        =   95
      Top             =   4185
      Width           =   3195
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   960
         Style           =   2  '드롭다운 목록
         TabIndex        =   106
         Top             =   1005
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer2.frx":0004
         Left            =   960
         List            =   "FrmTcpServer2.frx":000E
         Style           =   2  '드롭다운 목록
         TabIndex        =   105
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "캡쳐"
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   75
         TabIndex        =   104
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer2.frx":001E
         Left            =   2520
         List            =   "FrmTcpServer2.frx":0020
         Style           =   2  '드롭다운 목록
         TabIndex        =   103
         Top             =   2730
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer2.frx":0022
         Left            =   2520
         List            =   "FrmTcpServer2.frx":0024
         Style           =   2  '드롭다운 목록
         TabIndex        =   102
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "전송"
         Height          =   330
         Index           =   4
         Left            =   2415
         TabIndex        =   101
         Top             =   3120
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "긴급"
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   1635
         TabIndex        =   100
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "차단기"
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   855
         TabIndex        =   99
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   98
         Text            =   "주차장내 절대 서행"
         Top             =   2730
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   97
         Text            =   "일단 정지..!!"
         Top             =   2400
         Width           =   2430
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   4
         Left            =   975
         TabIndex        =   96
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "위치"
         Height          =   210
         Index           =   23
         Left            =   300
         TabIndex        =   109
         Top             =   1065
         Width           =   840
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종류"
         Height          =   285
         Index           =   11
         Left            =   300
         TabIndex        =   108
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "명칭"
         Height          =   210
         Index           =   24
         Left            =   300
         TabIndex        =   107
         Top             =   645
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   90
         X2              =   3165
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE6 "
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Index           =   5
      Left            =   16110
      TabIndex        =   79
      Top             =   4185
      Width           =   3195
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   5
         Left            =   975
         TabIndex        =   90
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   89
         Text            =   "일단 정지..!!"
         Top             =   2400
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   88
         Text            =   "주차장내 절대 서행"
         Top             =   2730
         Width           =   2430
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "차단기"
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   855
         TabIndex        =   87
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "긴급"
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   1635
         TabIndex        =   86
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "전송"
         Height          =   330
         Index           =   5
         Left            =   2415
         TabIndex        =   85
         Top             =   3120
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer2.frx":0026
         Left            =   2520
         List            =   "FrmTcpServer2.frx":0028
         Style           =   2  '드롭다운 목록
         TabIndex        =   84
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer2.frx":002A
         Left            =   2520
         List            =   "FrmTcpServer2.frx":002C
         Style           =   2  '드롭다운 목록
         TabIndex        =   83
         Top             =   2730
         Width           =   615
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "캡쳐"
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   75
         TabIndex        =   82
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer2.frx":002E
         Left            =   960
         List            =   "FrmTcpServer2.frx":0038
         Style           =   2  '드롭다운 목록
         TabIndex        =   81
         Top             =   240
         Width           =   1725
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   960
         Style           =   2  '드롭다운 목록
         TabIndex        =   80
         Top             =   1005
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   90
         X2              =   3165
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "명칭"
         Height          =   210
         Index           =   25
         Left            =   300
         TabIndex        =   93
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종류"
         Height          =   285
         Index           =   12
         Left            =   300
         TabIndex        =   92
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "위치"
         Height          =   210
         Index           =   26
         Left            =   300
         TabIndex        =   91
         Top             =   1065
         Width           =   840
      End
   End
   Begin MSWinsockLib.Winsock Aps_UDP 
      Left            =   11070
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ApsS_sock 
      Left            =   10515
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chk_ApsYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Height          =   225
      Left            =   30990
      TabIndex        =   75
      ToolTipText     =   "호스트pc와 운영pc를 분리할 경우, 운영pc라면  ""사용"" 체크 하세요"
      Top             =   1440
      Width           =   660
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   200
      Left            =   7920
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   200
      Left            =   7485
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   200
      Left            =   7050
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   6630
      Top             =   0
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "무인정산기"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   29265
      TabIndex        =   66
      Top             =   1590
      Width           =   2475
      Begin VB.TextBox TxtAspIp 
         Height          =   315
         Left            =   180
         TabIndex        =   76
         Text            =   "255.255.255.255"
         ToolTipText     =   "관리pc와 운영pc를 분리할 경우  ""사용"" 체크 하세요. 관리pc ip주소를 넣으세요"
         Top             =   495
         Width           =   1395
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "아이피"
         Height          =   255
         Left            =   225
         TabIndex        =   77
         Top             =   285
         Width           =   885
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   0
      TabIndex        =   63
      Top             =   7905
      Width           =   19290
      _Version        =   65536
      _ExtentX        =   34025
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "시스템 로그창"
      BackColor       =   32896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check2 
         BackColor       =   &H00008080&
         Caption         =   "Refresh"
         Height          =   225
         Left            =   18240
         TabIndex        =   65
         Top             =   150
         Value           =   1  '확인
         Width           =   945
      End
      Begin VB.CommandButton Command8 
         Caption         =   "클리어"
         Height          =   315
         Left            =   17160
         TabIndex        =   64
         Top             =   105
         Width           =   975
      End
   End
   Begin VB.CheckBox chk_RemoteYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Height          =   195
      Index           =   0
      Left            =   29805
      TabIndex        =   49
      ToolTipText     =   "호스트pc와 운영pc를 분리할 경우, 관리pc라면  ""사용"" 체크 하세요"
      Top             =   255
      Width           =   690
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Enabled         =   0   'False
      Height          =   210
      Index           =   3
      Left            =   12210
      TabIndex        =   32
      Top             =   4065
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Enabled         =   0   'False
      Height          =   210
      Index           =   2
      Left            =   9000
      TabIndex        =   31
      Top             =   4065
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Enabled         =   0   'False
      Height          =   210
      Index           =   1
      Left            =   5760
      TabIndex        =   30
      Top             =   4065
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Enabled         =   0   'False
      Height          =   210
      Index           =   0
      Left            =   2550
      TabIndex        =   29
      Top             =   4065
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   20790
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   3
      Left            =   21555
      Top             =   1515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   2
      Left            =   21135
      Top             =   1515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   1
      Left            =   20715
      Top             =   1515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   3
      Left            =   21555
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   2
      Left            =   21135
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   1
      Left            =   20715
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   0
      Left            =   20310
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   0
      Left            =   20295
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   0
      Left            =   20295
      Top             =   1515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE4 "
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Index           =   3
      Left            =   9690
      TabIndex        =   10
      Top             =   4185
      Width           =   3195
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   960
         Style           =   2  '드롭다운 목록
         TabIndex        =   73
         Top             =   1005
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer2.frx":0048
         Left            =   960
         List            =   "FrmTcpServer2.frx":0052
         Style           =   2  '드롭다운 목록
         TabIndex        =   61
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "캡쳐"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   75
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer2.frx":0062
         Left            =   2520
         List            =   "FrmTcpServer2.frx":0064
         Style           =   2  '드롭다운 목록
         TabIndex        =   26
         Top             =   2730
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer2.frx":0066
         Left            =   2520
         List            =   "FrmTcpServer2.frx":0068
         Style           =   2  '드롭다운 목록
         TabIndex        =   25
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "일반"
         Height          =   330
         Index           =   3
         Left            =   2415
         TabIndex        =   24
         Top             =   3120
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "긴급"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1635
         TabIndex        =   23
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "차단기"
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   855
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Text            =   "주차장내 절대 서행"
         Top             =   2730
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Text            =   "일단 정지..!!"
         Top             =   2400
         Width           =   2430
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   3
         Left            =   975
         TabIndex        =   11
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "위치"
         Height          =   210
         Index           =   19
         Left            =   300
         TabIndex        =   74
         Top             =   1065
         Width           =   840
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종류"
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   62
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "명칭"
         Height          =   210
         Index           =   16
         Left            =   300
         TabIndex        =   12
         Top             =   645
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   90
         X2              =   3165
         Y1              =   1590
         Y2              =   1590
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE3 "
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Index           =   2
      Left            =   6480
      TabIndex        =   7
      Top             =   4185
      Width           =   3195
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   960
         Style           =   2  '드롭다운 목록
         TabIndex        =   71
         Top             =   990
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer2.frx":006A
         Left            =   960
         List            =   "FrmTcpServer2.frx":0074
         Style           =   2  '드롭다운 목록
         TabIndex        =   59
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "캡쳐"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   27
         Top             =   3135
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer2.frx":0084
         Left            =   2520
         List            =   "FrmTcpServer2.frx":0086
         Style           =   2  '드롭다운 목록
         TabIndex        =   19
         Top             =   2745
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer2.frx":0088
         Left            =   2520
         List            =   "FrmTcpServer2.frx":008A
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   2415
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "전송"
         Height          =   330
         Index           =   2
         Left            =   2400
         TabIndex        =   17
         Top             =   3135
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "긴급"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1620
         TabIndex        =   16
         Top             =   3135
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "차단기"
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   855
         TabIndex        =   15
         Top             =   3135
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Text            =   "주차장내 절대 서행"
         Top             =   2745
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Text            =   "일단 정지..!!"
         Top             =   2415
         Width           =   2430
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   2
         Left            =   975
         TabIndex        =   8
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "위치"
         Height          =   210
         Index           =   14
         Left            =   300
         TabIndex        =   72
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종류"
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   60
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "명칭"
         Height          =   210
         Index           =   11
         Left            =   300
         TabIndex        =   9
         Top             =   645
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   90
         X2              =   3165
         Y1              =   1605
         Y2              =   1605
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE2 "
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Index           =   1
      Left            =   3255
      TabIndex        =   6
      Top             =   4185
      Width           =   3195
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   960
         Style           =   2  '드롭다운 목록
         TabIndex        =   69
         Top             =   975
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer2.frx":008C
         Left            =   975
         List            =   "FrmTcpServer2.frx":0096
         Style           =   2  '드롭다운 목록
         TabIndex        =   57
         Top             =   210
         Width           =   1725
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   1
         Left            =   975
         TabIndex        =   47
         Text            =   "정문"
         Top             =   570
         Width           =   1725
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer2.frx":00A6
         Left            =   2520
         List            =   "FrmTcpServer2.frx":00A8
         Style           =   2  '드롭다운 목록
         TabIndex        =   46
         Top             =   2745
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer2.frx":00AA
         Left            =   2520
         List            =   "FrmTcpServer2.frx":00AC
         Style           =   2  '드롭다운 목록
         TabIndex        =   45
         Top             =   2415
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "전송"
         Height          =   330
         Index           =   1
         Left            =   2415
         TabIndex        =   44
         Top             =   3135
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "긴급"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1635
         TabIndex        =   43
         Top             =   3135
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "차단기"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   840
         TabIndex        =   42
         Top             =   3135
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   41
         Text            =   "주차장내 절대 서행"
         Top             =   2745
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   40
         Text            =   "일단 정지..!!"
         Top             =   2415
         Width           =   2430
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "캡쳐"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   39
         Top             =   3135
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "위치"
         Height          =   210
         Index           =   9
         Left            =   300
         TabIndex        =   70
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종류"
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   58
         Top             =   270
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   90
         X2              =   3165
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "명칭"
         Height          =   210
         Index           =   6
         Left            =   300
         TabIndex        =   48
         Top             =   615
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE1 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   4185
      Width           =   3195
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "FrmTcpServer2.frx":00AE
         Left            =   930
         List            =   "FrmTcpServer2.frx":00B0
         Style           =   2  '드롭다운 목록
         TabIndex        =   67
         Top             =   990
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer2.frx":00B2
         Left            =   945
         List            =   "FrmTcpServer2.frx":00BC
         Style           =   2  '드롭다운 목록
         TabIndex        =   55
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "차단기"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   865
         TabIndex        =   53
         Top             =   3105
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "긴급"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1640
         TabIndex        =   52
         Top             =   3105
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "전송"
         Height          =   330
         Index           =   0
         Left            =   2415
         TabIndex        =   51
         Top             =   3105
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "캡쳐"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   50
         Top             =   3105
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer2.frx":00CC
         Left            =   2520
         List            =   "FrmTcpServer2.frx":00CE
         Style           =   2  '드롭다운 목록
         TabIndex        =   37
         Top             =   2745
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer2.frx":00D0
         Left            =   2520
         List            =   "FrmTcpServer2.frx":00D2
         Style           =   2  '드롭다운 목록
         TabIndex        =   36
         Top             =   2415
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   35
         Text            =   "주차장내 절대 서행"
         Top             =   2745
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Text            =   "일단 정지..!!"
         Top             =   2415
         Width           =   2430
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   0
         Left            =   945
         TabIndex        =   33
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "위치"
         Height          =   210
         Index           =   4
         Left            =   270
         TabIndex        =   68
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "종류"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   56
         Top             =   300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   3135
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "명칭"
         Height          =   210
         Index           =   2
         Left            =   270
         TabIndex        =   38
         Top             =   645
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "데이터 수신"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Index           =   0
      Left            =   28665
      TabIndex        =   2
      Top             =   360
      Width           =   1755
      Begin VB.TextBox TxtSvrPort 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Text            =   "10000"
         ToolTipText     =   "호스트pc와 운영pc를 분리할 경우  ""사용"" 체크 하세요. 운영pc에서 수신할 포트번호 입니다."
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "포트"
         Height          =   255
         Left            =   210
         TabIndex        =   54
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   -12795
      TabIndex        =   1
      Top             =   4530
      Width           =   75
   End
   Begin VB.ListBox ListData 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1950
      Left            =   0
      TabIndex        =   0
      Top             =   8430
      Width           =   19320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   20310
      Top             =   300
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   1
      Left            =   20715
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   2
      Left            =   21135
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   3
      Left            =   21555
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock RemoteR_sock 
      Left            =   21750
      Top             =   3015
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock RemoteS_sock 
      Left            =   21330
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock MvrSock 
      Left            =   20310
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   4
      Left            =   21990
      Top             =   1515
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   4
      Left            =   21990
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   4
      Left            =   21990
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   5
      Left            =   22410
      Top             =   2430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   5
      Left            =   22440
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   5
      Left            =   22440
      Top             =   1530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   0
      Left            =   20310
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   1
      Left            =   20745
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   2
      Left            =   21180
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   3
      Left            =   21600
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   4
      Left            =   22020
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   5
      Left            =   22455
      Top             =   855
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   0
      Left            =   17010
      TabIndex        =   110
      Top             =   30
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "적 용"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "FrmTcpServer2.frx":00D4
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   1
      Left            =   18180
      TabIndex        =   111
      Top             =   30
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "닫 기"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "FrmTcpServer2.frx":0425
   End
   Begin LPR_PARKING_HOST.Server Server 
      Left            =   6000
      Top             =   30
      _extentx        =   741
      _extenty        =   741
   End
   Begin VB.Frame frmParkFull 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 만차 설정 "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1635
      Left            =   4155
      TabIndex        =   124
      Top             =   2160
      Width           =   3210
      Begin VB.CheckBox chk_RegCarPass_YN 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "만차시 등록차량 입차허용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   130
         Top             =   1140
         Width           =   2655
      End
      Begin VB.TextBox txt_max_park 
         Height          =   315
         Left            =   2010
         TabIndex        =   126
         Text            =   "102"
         Top             =   390
         Width           =   630
      End
      Begin VB.TextBox txt_now_park 
         Height          =   315
         Left            =   2010
         TabIndex        =   125
         Text            =   "101"
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "현재 주차대수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   128
         Top             =   810
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전체 주차대수"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   240
         TabIndex        =   127
         Top             =   390
         Width           =   1275
      End
   End
   Begin VB.Label lbl_CertifyLimitDate 
      BackColor       =   &H000000FF&
      Caption         =   "만료기간:2019-01-01"
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5670
      TabIndex        =   139
      ToolTipText     =   "만료기간 이내 인증하세요. 만료기간 이후에는 안정적인 서비스 운영이 불가능합니다."
      Top             =   720
      Width           =   2520
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BackStyle       =   0  '투명
      Caption         =   " # 주의 #  시스템 관리자 외에 절대 수정 금지..!!"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   150
      TabIndex        =   5
      Top             =   165
      Width           =   5145
   End
End
Attribute VB_Name = "FrmTcpServer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const white = &H80000005
Const grey = &H8000000F



Private Sub chk_ParkFull_YN_Click()
    
    If chk_ParkFull_YN.value = 1 Then
        frmParkFull.Enabled = True
        txt_max_park.Enabled = True
        txt_now_park.Enabled = True
        chk_RegCarPass_YN.Enabled = True
    
    Else
        frmParkFull.Enabled = False
        txt_max_park.Enabled = False
        txt_now_park.Enabled = False
        chk_RegCarPass_YN.Enabled = False
    End If
End Sub

Private Sub chk_HomeNet_YN_Click()
    If chk_HomeNet_YN.value = 1 Then
        frmHomeNet.Enabled = True
    Else
        frmHomeNet.Enabled = False
    End If
End Sub



Private Sub cmd_Button_Click(Index As Integer)
    ' 통합 저장 및 적용
    If (Index = 1) Then
        Me.Hide
        Exit Sub
    End If

    Dim i As Integer
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '홈넷 설정
    cmd_HomeTest.Enabled = True
    If (chk_HomeNet_YN.value = 1) Then
        HomeNet_YN = "Y"
    Else
        HomeNet_YN = "N"
    End If
    HomeNet_IP = Trim(Text_HomeNet_IP)
    HomeNet_Port = Val(Text_HomeNet_Port)
    Call Put_Ini2("System Config", "HomeNetMode", CStr(cmb_HomeNet.ListIndex + 1), "C:\HomeNet\HomeNet.ini")
    
    Shell ("taskkill /f /im HomeNet.exe")
    If (HomeNet_YN = "Y") Then
        If (IsFile("C:\HomeNet\HomeNet.exe") = True) Then
            
            Delay_Time (1)
            Shell ("C:\HomeNet\HomeNet.exe")
            Delay_Time (2)
            
            FrmTcpServer.HomeSock.Close
            FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
            FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
            FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
        End If
    End If
    Call Put_Ini("System Config", "HomeNet_YN", HomeNet_YN)
    Call Put_Ini("System Config", "HomeNet_IP", HomeNet_IP)
    Call Put_Ini("System Config", "HomeNet_Port", CStr(HomeNet_Port))
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' 만차 설정
    If (chk_ParkFull_YN.value = 1) Then
        Glo_ParkFull_YN = "Y"
    Else
        Glo_ParkFull_YN = "N"
    End If
    Glo_ParkFull_Count = Val(txt_max_park)
    Glo_ParkNow_Count = Val(txt_now_park)
    If (chk_RegCarPass_YN.value = 1) Then
        Glo_ParkRegIn_YN = "Y"
    Else
        Glo_ParkRegIn_YN = "N"
    End If
    Call Put_Ini("System Config", "ParkFull_YN", Glo_ParkFull_YN)
    Call Put_Ini("System Config", "ParkFull_Count", CStr(Glo_ParkFull_Count))
    Call Put_Ini("System Config", "ParkNow_Count", CStr(Glo_ParkNow_Count))
    Call Put_Ini("System Config", "ParkRegIn_YN", Glo_ParkRegIn_YN)
    
    '만차
    Call ParkFull_Set
    
    '만차등
    Call ParkFullLight_Set
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    Call frmLogin.ShowMenu(Glo_Login_ID, Glo_Login_PW)
    Me.Show 0

    FrmTcpServer2.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "[환경설정 저장]", 0
    Call DataLogger("[환경설정 저장]")

End Sub






Private Sub cmd_Certify_Click()

    Dim rs As ADODB.Recordset
    Dim qry As String
    Dim bQryResult As Boolean
    Dim sIP, sMac As String
    
    Call FrmTcpServer.GetClientIP(Glo_IPAddr)
    Call FrmTcpServer.GetClientMac(Glo_MacAddr)
    Call FrmTcpServer.GetClienKey(Glo_PhyHDDKey)

'    If (cmd_Certify.Caption = "인증") Then
    If (Glo_Certify = enumCertify.eCertNoTry) Then '미인증상태
        
        
        
        Set rs = New ADODB.Recordset
        qry = "SELECT LockDate, UnLockDate FROM tb_Certify WHERE HASHCODE = '" & Glo_PhyHDDKey & "' "
    
        If (DataBaseQuery(rs, adoConn, qry, NWERR_GATE_STAY) = False) Then
            Call DebugLogger("[CERTIFY]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        If rs.EOF Then
            bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO tb_certify (LockDate, UnLockDate, IP, Mac, Hashcode, Memo ) VALUES ('" & Format(Now, "yyyy-mm-dd") & "','', '" & Glo_IPAddr & "', '" & Glo_MacAddr & "', '" & Glo_PhyHDDKey & "', '')", NWERR_GATE_STAY, 0)
            If (bQryResult = False) Then
                DataLogger ("[CERTIFY]    " & "인증이 원할하지 않았습니다. 다시 시도해주세요.")
                Set rs = Nothing
                Exit Sub
            End If
            
        Else
            bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_certify SET LockDate='" & Format(Now, "yyyy-mm-dd") & "' WHERE HASHCODE = '" & Glo_PhyHDDKey & "' ", NWERR_GATE_STAY)
            If (bQryResult = False) Then
                DataLogger ("[CERTIFY]    " & "인증이 원할하지 않았습니다. 다시 시도해주세요.")
                Set rs = Nothing
                Exit Sub
            End If
            
            
            
        End If
    
    
        cmd_Certify.Caption = "인증필요"
        cmd_Certify.ToolTipText = "반드시 만료일 이전까지 인증받으세요. 만료일이후부터 차단기가 정상동작하지 않습니다"
        txt_CertifyKey.Visible = True
        lbl_CertifyLimitDate.Caption = "만료기간:" & DateAdd("m", Glo_Cert_Month, Format(Now, "yyyy-mm-dd"))
        lbl_CertifyLimitDate.Visible = True

        Glo_Cert_LimitDate = DateAdd("m", Glo_Cert_Month, Format(Now, "yyyy-mm-dd")) '만료일
        Glo_Cert_NoticeSDate = DateAdd("m", Glo_Cert_Month - 1, Format(Now, "yyyy-mm-dd")) '만료기간 안내 시작일(만료일 1개월 전)
        Glo_Certify = enumCertify.eCertTry
            

'    ElseIf (cmd_Certify.Caption = "인증필요") Then
    ElseIf (Glo_Certify = enumCertify.eCertTry) Then

        If (Len(txt_CertifyKey.text) = 0) Then
            Call DataLogger("인증키 입력해주세요")

        Else
            If (txt_CertifyKey.text = "admin" & Format(Now, "ddmm")) Then 'admin일월
                bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_certify SET UnLockDate='" & Format(Now, "yyyy-mm-dd") & "' WHERE HASHCODE = '" & Glo_PhyHDDKey & "' ", NWERR_GATE_STAY)
                If (bQryResult = False) Then
                    DataLogger ("[Certify_Click]    " & "인증진행이 원할하지 않았습니다. 다시 시도해주세요.")
                    Set rs = Nothing
                    Exit Sub
                End If

                cmd_Certify.Caption = "인증완료"
                cmd_Certify.ToolTipText = ""
                cmd_Certify.Visible = True
                cmd_Certify.Enabled = False
                txt_CertifyKey.Visible = False
                lbl_CertifyLimitDate.Visible = False
                txt_CertifyKey.text = ""
                Glo_Certify = enumCertify.eCertOK
                Call DataLogger("인증성공")

            ElseIf (txt_CertifyKey.text = "jawootek" & Format(Now, "ddmm")) Then 'admin일월
                 Glo_COMPANY = "(주)자우텍"
            Else
                txt_CertifyKey.text = ""
                Call DataLogger("인증실패!!")
            End If
        End If
    End If
    
End Sub

Private Sub Command1_Click()
    FormSound.Show 1
End Sub

Private Sub Form_Activate()
    Dim Port As Integer
    Dim i As Integer
    Dim bScrNoChk As Boolean
    Dim iHomeNetNo As Integer
    
    
On Error GoTo Err_Proc
    
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    Call Certify '인증
    
    
    
    '홈넷 설정
    '파일로드
    iHomeNetNo = Val(Get_Ini2("System Config", "HomeNetMode", "1", "C:\HomeNet\HomeNet.ini"))
    If HomeNet_YN = "Y" Then
        chk_HomeNet_YN.value = 1
        frmHomeNet.Enabled = True
    Else
        chk_HomeNet_YN.value = 0
        frmHomeNet.Enabled = False
    End If
    Text_HomeNet_IP = Trim(HomeNet_IP)
    Text_HomeNet_Port = Val(HomeNet_Port)
    
    cmb_HomeNet.Clear
    With cmb_HomeNet
        .AddItem "1.현대통신"
        .AddItem "2.서울통신(DB)"
        .AddItem "3.이지빌"
        .AddItem "4.코콤"
        .AddItem "5.코맥스"
        .AddItem "6.아이콘트롤스"
        .AddItem "7.경동 원"
        .AddItem "8.LG전자"
        .AddItem "9.서울통신(TCP)"
        .AddItem "10.현대통신(리눅스서버)"
        .AddItem "11.맥서러시(GS 네오텍)"
        .AddItem "12.홈클래버"
    End With
    cmb_HomeNet.text = cmb_HomeNet.List(iHomeNetNo - 1)
    
    
    
    If (Glo_ParkFull_YN = "Y") Then
        chk_ParkFull_YN.value = 1
        txt_max_park.Enabled = True
        txt_now_park.Enabled = True
        chk_RegCarPass_YN.Enabled = True
    Else
        chk_ParkFull_YN.value = 0
        txt_max_park.Enabled = False
        txt_now_park.Enabled = False
        chk_RegCarPass_YN.Enabled = False
    End If
    txt_max_park = CStr(Glo_ParkFull_Count)
    txt_now_park = CStr(Glo_ParkNow_Count)
    If (Glo_ParkRegIn_YN = "Y") Then
        chk_RegCarPass_YN.value = 1
    Else
        chk_RegCarPass_YN.value = 0
    End If
    
    If (Glo_GuestLogBackup_YN = "Y") Then
        txt_GuestLogBackup.Enabled = True
        frmGuestLogBackup.Enabled = True
    Else
        txt_GuestLogBackup.Enabled = False
        frmGuestLogBackup.Enabled = False
    End If
    
    Dim sClor1  As String
    Dim sClor2  As String
    Dim sClor3  As String
    If (Glo_Display = "전광판(풀컬러)") Then
        sClor1 = "녹"
        sClor2 = "적"
        sClor3 = "황"
    Else
        sClor1 = "적"
        sClor2 = "녹"
        sClor3 = "황"
    End If
    For i = 0 To 5
        cmb_Disp1(i).Clear
        cmb_Disp1(i).AddItem "녹"
        cmb_Disp1(i).AddItem "적"
        cmb_Disp1(i).AddItem "황"
        cmb_Disp2(i).Clear
        cmb_Disp2(i).AddItem "녹"
        cmb_Disp2(i).AddItem "적"
        cmb_Disp2(i).AddItem "황"
    Next i
    
    
    If (LANE1_YN = "Y") Then
        chk_UseYN(0).value = "1"
        Frame2(0).Enabled = True
    Else
        chk_UseYN(0).value = "0"
        Frame2(0).Enabled = False
    End If
    If (LANE1_Inout = "입구") Then
        cmb_Inout(0).ListIndex = 0
    Else
        cmb_Inout(0).ListIndex = 1
    End If
    txt_GateName(0).text = LANE1_Name
    txt_Disp1(0) = LANE1_Disp1Msg
    txt_Disp2(0) = LANE1_Disp2Msg
    cmb_Disp1(0).ListIndex = LANE1_Disp1Color
    cmb_Disp2(0).ListIndex = LANE1_Disp2Color
    
    If (LANE2_YN = "Y") Then
        chk_UseYN(1).value = "1"
        Frame2(1).Enabled = True
    Else
        chk_UseYN(1).value = "0"
        Frame2(1).Enabled = False
    End If
    If (LANE2_Inout = "입구") Then
        cmb_Inout(1).ListIndex = 0
    Else
        cmb_Inout(1).ListIndex = 1
    End If
    txt_GateName(1).text = LANE2_Name
    txt_Disp1(1) = LANE2_Disp1Msg
    txt_Disp2(1) = LANE2_Disp2Msg
    cmb_Disp1(1).ListIndex = LANE2_Disp1Color
    cmb_Disp2(1).ListIndex = LANE2_Disp2Color
    
    If (LANE3_YN = "Y") Then
        chk_UseYN(2).value = "1"
        Frame2(2).Enabled = True
    Else
        chk_UseYN(2).value = "0"
        Frame2(2).Enabled = False
    End If
    If (LANE3_Inout = "입구") Then
        cmb_Inout(2).ListIndex = 0
    Else
        cmb_Inout(2).ListIndex = 1
    End If
    txt_GateName(2).text = LANE3_Name
    txt_Disp1(2) = LANE3_Disp1Msg
    txt_Disp2(2) = LANE3_Disp2Msg
    cmb_Disp1(2).ListIndex = LANE3_Disp1Color
    cmb_Disp2(2).ListIndex = LANE3_Disp2Color
    
    If (LANE4_YN = "Y") Then
        chk_UseYN(3).value = "1"
        Frame2(3).Enabled = True
    Else
        chk_UseYN(3).value = "0"
        Frame2(3).Enabled = False
    End If
    If (LANE4_Inout = "입구") Then
        cmb_Inout(3).ListIndex = 0
    Else
        cmb_Inout(3).ListIndex = 1
    End If
    txt_GateName(3).text = LANE4_Name
    txt_Disp1(3) = LANE4_Disp1Msg
    txt_Disp2(3) = LANE4_Disp2Msg
    cmb_Disp1(3).ListIndex = LANE4_Disp1Color
    cmb_Disp2(3).ListIndex = LANE4_Disp2Color
    
    If (LANE5_YN = "Y") Then
        chk_UseYN(4).value = "1"
        Frame2(4).Enabled = True
    Else
        chk_UseYN(4).value = "0"
        Frame2(4).Enabled = False
    End If
    If (LANE5_Inout = "입구") Then
        cmb_Inout(4).ListIndex = 0
    Else
        cmb_Inout(4).ListIndex = 1
    End If
    txt_GateName(4).text = LANE5_Name
    txt_Disp1(4) = LANE5_Disp1Msg
    txt_Disp2(4) = LANE5_Disp2Msg
    cmb_Disp1(4).ListIndex = LANE5_Disp1Color
    cmb_Disp2(4).ListIndex = LANE5_Disp2Color
 
    If LANE6_YN = "Y" Then
        chk_UseYN(5).value = 1
        Frame2(5).Enabled = True
    Else
        chk_UseYN(5).value = 0
        Frame2(5).Enabled = False
    End If
    If (LANE6_Inout = "입구") Then
        cmb_Inout(5).ListIndex = 0
    Else
        cmb_Inout(5).ListIndex = 1
    End If
    txt_GateName(5).text = LANE6_Name
    txt_Disp1(5) = LANE6_Disp1Msg
    txt_Disp2(5) = LANE6_Disp2Msg
    cmb_Disp1(5).ListIndex = LANE6_Disp1Color
    cmb_Disp2(5).ListIndex = LANE6_Disp2Color
    
    
Exit Sub

Err_Proc:
    MsgBox ("[FormLoad_Proc]  " & Err.Description)
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [TCP Server Load Proc]  " & Err.Description)


End Sub
    


Private Sub Form_Load()
Call Form_Activate
End Sub

Private Sub Certify()

    Dim rs As ADODB.Recordset
    Dim qry As String
    Dim LockDate As String
    Dim UnLockDate As String

On Error GoTo Err_P


    Glo_Certify = enumCertify.eCertNoTry
    
    Set rs = New ADODB.Recordset
    qry = "SELECT LockDate, UnLockDate FROM tb_Certify "

    If (DataBaseQuery(rs, adoConn, qry, NWERR_GATE_STAY) = False) Then
        Call DebugLogger("[CERTIFY]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If

''''    If rs.EOF Then
''''        Glo_Certify = True   '초기값은 인증자체를 진행하지 않았으므로 정상으로 처리
''''        cmd_Certify.Caption = "인증"
''''        cmd_Certify.ToolTipText = "인증키를 가지고 있는 시스템관리자만 사용하세요"
''''        txt_CertifyKey.Visible = False
''''        lbl_CertifyLimitDate.Visible = False
''''
''''    Else
    

'''                LockDate = "" & rs!LockDate
'''                UnLockDate = "" & rs!UnLockDate
'''
'''                If (Len(UnLockDate) > 0) Then    '인증받았을 경우
'''                    Glo_Certify = True
'''
'''                    cmd_Certify.Caption = "인증완료"
'''                    cmd_Certify.ToolTipText = ""
'''                    cmd_Certify.Enabled = False
'''                    txt_CertifyKey.Visible = False
'''                    lbl_CertifyLimitDate.Visible = False
'''
'''                ElseIf (Len(LockDate) > 0) Then
'''                    Glo_Certify = False
'''                    Glo_Cert_LimitDate = DateAdd("m", Glo_Cert_Month, LockDate)
'''                    Glo_Cert_NoticeSDate = DateAdd("m", Glo_Cert_Month - 1, LockDate) '만료기간 안내 시작일(만료일 1개월 전)
'''
'''                    cmd_Certify.Caption = "인증필요"
'''                    cmd_Certify.ToolTipText = "반드시 만료일 이전까지 인증받으세요. 만료일이후부터 차단기가 정상동작하지 않습니다"
'''                    txt_CertifyKey.Visible = True
'''                    lbl_CertifyLimitDate.Caption = "만료기간:" & Glo_Cert_LimitDate
'''                    lbl_CertifyLimitDate.Visible = True
'''
'''                ElseIf (Len(LockDate) <= 0) Then
'''                    Glo_Certify = True
'''                End If
'''
'''        End If

        Do While Not (rs.EOF)

                LockDate = "" & rs!LockDate
                UnLockDate = "" & rs!UnLockDate
                
                
                If (Len(LockDate) > 0) Then    '인증버튼 눌렀을 경우
                    
                    If (Len(UnLockDate) > 0) Then '인증 완료한 경우
                        Glo_Certify = enumCertify.eCertOK

                    Else
                        Glo_Certify = enumCertify.eCertTry

                    End If

                    Exit Do
                End If
                
                rs.MoveNext
        Loop
        
        
        If (Glo_Certify = enumCertify.eCertNoTry) Then
            cmd_Certify.Caption = "인증"
            cmd_Certify.ToolTipText = "인증키를 가지고 있는 시스템관리자만 사용하세요"
            txt_CertifyKey.Visible = False
            lbl_CertifyLimitDate.Visible = False
        
        ElseIf (Glo_Certify = enumCertify.eCertTry) Then
            Glo_Cert_LimitDate = DateAdd("m", Glo_Cert_Month, LockDate)
            Glo_Cert_NoticeSDate = DateAdd("m", Glo_Cert_Month - 1, LockDate) '만료기간 안내 시작일(만료일 1개월 전)
            
            cmd_Certify.Caption = "인증필요"
            cmd_Certify.ToolTipText = "반드시 만료일 이전까지 인증받으세요. 만료일이후부터 차단기가 정상동작하지 않습니다"
            txt_CertifyKey.Visible = True
            lbl_CertifyLimitDate.Caption = "만료기간:" & Glo_Cert_LimitDate
            lbl_CertifyLimitDate.Visible = True
        
        ElseIf (Glo_Certify = enumCertify.eCertOK) Then
            cmd_Certify.Caption = "인증완료"
            cmd_Certify.ToolTipText = ""
            cmd_Certify.Enabled = False
            txt_CertifyKey.Visible = False
            lbl_CertifyLimitDate.Visible = False
        End If
    
    Set rs = Nothing
    
    
    Exit Sub
    
Err_P:
    Set rs = Nothing
    Call DebugLogger("[CERTIFY] Cert Res:" & Glo_Certify & ", Limit Date: " & Glo_Cert_LimitDate & ", Err: " & Err.Description)
    
End Sub


Private Sub chk_UseYN_Click(Index As Integer)
    If (Index > Glo_Screen_No - 1) Then
        chk_UseYN(Index).value = "0"
    Else
        If chk_UseYN(Index).value = "1" Then
            Frame2(Index).Enabled = True
        Else
            Frame2(Index).Enabled = False
        End If
    End If

End Sub


Private Sub cmb_Disp1_Click(Index As Integer)
    Dim cmbIndex As Integer
    Dim cmbColor As Long

    cmbIndex = cmb_Disp1(Index).ListIndex

    Select Case cmbIndex
        Case 0
            cmbColor = &HFF00& ' 녹색
        Case 1
            cmbColor = &HFF&   ' 적색
        Case 2
            cmbColor = &H80C0FF ' 황색
    End Select

    txt_Disp1(Index).ForeColor = cmbColor
End Sub

Private Sub cmb_Disp2_Click(Index As Integer)
    Dim cmbIndex As Integer
    Dim cmbColor As Long

    cmbIndex = cmb_Disp2(Index).ListIndex

    Select Case cmbIndex
        Case 0
            cmbColor = &HFF00&  ' 녹색
        Case 1
            cmbColor = &HFF&    ' 적색
        Case 2
            cmbColor = &H80C0FF ' 황색
    End Select

    txt_Disp2(Index).ForeColor = cmbColor
End Sub




Private Sub cmd_CapTest_Click(Index As Integer)
    
    'Capture Test
    Call DataLogger("[Get Frame TEST]  Target Gate = " & Index)
    Call Relay_Out(1, Index)
    'Call Delay_Time(1000)
    
End Sub

Private Sub cmd_EmgTest_Click(Index As Integer)
    
    'Display Emg Test
    Call DataLogger("[DISPLAY Emg TEST]  Target Gate = " & Index)
    Call GL_Emergency("System Test", "System Test", 0, 30, 10, 1, 2, 1, Index)

End Sub

Private Sub cmd_GateTest_Click(Index As Integer)
    
    'Gate Test
    Call DataLogger("[GATE TEST]  Target Gate = " & Index)
    Call Relay_Out(0, Index)

End Sub

Private Sub cmd_HomeTest_Click()

    HomeNet_Dong = txt_Dong.text
    HomeNet_Ho = txt_Ho.text
    HomeNet_CarNo = "서울01가1234"
    HomeNet_Str = HomeNet_Dong & HomeNet_Ho & HomeNet_CarNo

    If (FrmTcpServer.HomeSock.State = sckClosed) Then

        If (HomeNet_IP <> "" And HomeNet_Port > 0) Then
        
            FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
            FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
            FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
    
            FrmTcpServer.HomeSock.SendData (HomeNet_Str)
            Call DataLogger("[HomeNet UDP 전송]  IP = " & HomeNet_IP & "    PORT = " & HomeNet_Port & "      DATA = " & HomeNet_Str)
            
        Else
            Call DataLogger("[HomeNet UDP 전송]  HomeNet IP 와 HomeNet Port 확인 및 저장해주세요")
        End If
    Else
        FrmTcpServer.HomeSock.SendData (HomeNet_Str)
        Call DataLogger("[HomeNet UDP 전송]  IP = " & HomeNet_IP & "    PORT = " & HomeNet_Port & "      DATA = " & HomeNet_Str)
    End If
    
    
    FrmTcpServer2.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "[HomeNet UDP 전송]  DATA = " & HomeNet_Str, 0
End Sub

Private Sub cmd_NmlTest_Click(Index As Integer)
    
    Dim upColor As Byte
    Dim downColor As Byte
    
    'Display Nomal Save

    If (Glo_Display = "전광판(풀컬러)_FW7") Then
        Select Case cmb_Disp1(Index).text
            Case "적"
                upColor = enumDIS_COLORs.eRED
            Case "황"
                upColor = enumDIS_COLORs.eYellow
            Case "녹"
                upColor = enumDIS_COLORs.eGreen
            Case "파"
                upColor = enumDIS_COLORs.eBLUE
            Case "자"
                upColor = enumDIS_COLORs.eWINE
            Case "하"
                upColor = enumDIS_COLORs.eSKY
            Case "백"
                upColor = enumDIS_COLORs.eWHITE
        End Select
        Select Case cmb_Disp2(Index).text
            Case "적"
                downColor = enumDIS_COLORs.eRED
            Case "황"
                downColor = enumDIS_COLORs.eYellow
            Case "녹"
                downColor = enumDIS_COLORs.eGreen
            Case "파"
                downColor = enumDIS_COLORs.eBLUE
            Case "자"
                downColor = enumDIS_COLORs.eWINE
            Case "하"
                downColor = enumDIS_COLORs.eSKY
            Case "백"
                downColor = enumDIS_COLORs.eWHITE
        End Select

    ElseIf (Glo_Display = "전광판(풀컬러)") Then '황:2, 초:1, 적:0
        Select Case cmb_Disp1(Index).text
            Case "적"
                'upColor = enumDIS_COLORs.eRED
                upColor = 0
            Case "황"
                'upColor = enumDIS_COLORs.eYellow
                upColor = 2
            Case "녹"
                'upColor = enumDIS_COLORs.eGreen
                upColor = 1
        End Select
        Select Case cmb_Disp2(Index).text
            Case "적"
                'downColor = enumDIS_COLORs.eRED
                downColor = 0
            Case "황"
                'downColor = enumDIS_COLORs.eYellow
                downColor = 2
            Case "녹"
                'downColor = enumDIS_COLORs.eGreen
                downColor = 1
        End Select

    Else
        Select Case cmb_Disp1(Index).text
            Case "녹"
                upColor = 0
            Case "적"
                upColor = 1
            Case "황"
                upColor = 2
        End Select
        Select Case cmb_Disp2(Index).text
            Case "녹"
                downColor = 0
            Case "적"
                downColor = 1
            Case "황"
                downColor = 2
        End Select
    End If
    


    Call DataLogger("[DISPLAY Nomal Save]  Target Gate = " & Index)
    'Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, cmb_Disp1(Index).ListIndex, cmb_Disp2(Index).ListIndex, Index)
    Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, upColor, downColor, Index)

    Select Case Index
        Case 0
            LANE1_Disp1Msg = txt_Disp1(0)
            LANE1_Disp2Msg = txt_Disp2(0)
            Call Put_Ini("System Config", "LANE1_Disp1Msg", txt_Disp1(0))
            Call Put_Ini("System Config", "LANE1_Disp2Msg", txt_Disp2(0))
            Call Put_Ini("System Config", "LANE1_Disp1Color ", CStr(cmb_Disp1(0).ListIndex))
            Call Put_Ini("System Config", "LANE1_Disp2Color ", CStr(cmb_Disp2(0).ListIndex))

        Case 1
            LANE2_Disp1Msg = txt_Disp1(1)
            LANE2_Disp2Msg = txt_Disp2(1)
            Call Put_Ini("System Config", "LANE2_Disp1Msg", txt_Disp1(1))
            Call Put_Ini("System Config", "LANE2_Disp2Msg", txt_Disp2(1))
            Call Put_Ini("System Config", "LANE2_Disp1Color ", CStr(cmb_Disp1(1).ListIndex))
            Call Put_Ini("System Config", "LANE2_Disp2Color ", CStr(cmb_Disp2(1).ListIndex))

        Case 2
            LANE3_Disp1Msg = txt_Disp1(2)
            LANE3_Disp2Msg = txt_Disp2(2)
            Call Put_Ini("System Config", "LANE3_Disp1Msg", txt_Disp1(2))
            Call Put_Ini("System Config", "LANE3_Disp2Msg", txt_Disp2(2))
            Call Put_Ini("System Config", "LANE3_Disp1Color ", CStr(cmb_Disp1(2).ListIndex))
            Call Put_Ini("System Config", "LANE3_Disp2Color ", CStr(cmb_Disp2(2).ListIndex))

        Case 3
            LANE4_Disp1Msg = txt_Disp1(3)
            LANE4_Disp2Msg = txt_Disp2(3)
            Call Put_Ini("System Config", "LANE4_Disp1Msg", txt_Disp1(3))
            Call Put_Ini("System Config", "LANE4_Disp2Msg", txt_Disp2(3))
            Call Put_Ini("System Config", "LANE4_Disp1Color ", CStr(cmb_Disp1(3).ListIndex))
            Call Put_Ini("System Config", "LANE4_Disp2Color ", CStr(cmb_Disp2(3).ListIndex))

        Case 4
            LANE5_Disp1Msg = txt_Disp1(4)
            LANE5_Disp2Msg = txt_Disp2(4)
            Call Put_Ini("System Config", "LANE5_Disp1Msg", txt_Disp1(4))
            Call Put_Ini("System Config", "LANE5_Disp2Msg", txt_Disp2(4))
            Call Put_Ini("System Config", "LANE5_Disp1Color ", CStr(cmb_Disp1(4).ListIndex))
            Call Put_Ini("System Config", "LANE5_Disp2Color ", CStr(cmb_Disp2(4).ListIndex))

        Case 5
            LANE6_Disp1Msg = txt_Disp1(5)
            LANE6_Disp2Msg = txt_Disp2(5)
            Call Put_Ini("System Config", "LANE6_Disp1Msg", txt_Disp1(5))
            Call Put_Ini("System Config", "LANE6_Disp2Msg", txt_Disp2(5))
            Call Put_Ini("System Config", "LANE6_Disp1Color ", CStr(cmb_Disp1(5).ListIndex))
            Call Put_Ini("System Config", "LANE6_Disp2Color ", CStr(cmb_Disp2(5).ListIndex))
    End Select

End Sub

'LANE Config Save & Effect
Private Sub cmd_OK_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If chk_UseYN(0).value = "1" Then
                LANE1_YN = "Y"
            Else
                LANE1_YN = "N"
            End If
            LANE1_Inout = cmb_Inout(0).text
            LANE1_Name = Trim(txt_GateName(0))
            Call Put_Ini("System Config", "LANE1_YN ", LANE1_YN)
            Call Put_Ini("System Config", "LANE1_INOUT ", LANE1_Inout)
            Call Put_Ini("System Config", "LANE1_Name ", LANE1_Name)
    
        Case 1
            If chk_UseYN(1).value = "1" Then
                LANE2_YN = "Y"
            Else
                LANE2_YN = "N"
            End If
            LANE2_Inout = cmb_Inout(1).text
            LANE2_Name = Trim(txt_GateName(1))
            Call Put_Ini("System Config", "LANE2_YN ", LANE2_YN)
            Call Put_Ini("System Config", "LANE2_INOUT ", LANE2_Inout)
            Call Put_Ini("System Config", "LANE2_Name ", LANE2_Name)
    
        Case 2
            If chk_UseYN(2).value = "1" Then
                LANE3_YN = "Y"
            Else
                LANE3_YN = "N"
            End If
            LANE3_Inout = cmb_Inout(2).text
            LANE3_Name = Trim(txt_GateName(2))
            Call Put_Ini("System Config", "LANE3_YN ", LANE3_YN)
            Call Put_Ini("System Config", "LANE3_INOUT ", LANE3_Inout)
            Call Put_Ini("System Config", "LANE3_Name ", LANE3_Name)

        Case 3
            If chk_UseYN(3).value = "1" Then
                LANE4_YN = "Y"
            Else
                LANE4_YN = "N"
            End If
            LANE4_Inout = cmb_Inout(3).text
            LANE4_Name = Trim(txt_GateName(3))
            Call Put_Ini("System Config", "LANE4_YN ", LANE4_YN)
            Call Put_Ini("System Config", "LANE4_INOUT ", LANE4_Inout)
            Call Put_Ini("System Config", "LANE4_Name ", LANE4_Name)
            
        Case 4
            If chk_UseYN(4).value = "1" Then
                LANE5_YN = "Y"
            Else
                LANE5_YN = "N"
            End If
            LANE5_Inout = cmb_Inout(4).text
            LANE5_Name = Trim(txt_GateName(4))
            Call Put_Ini("System Config", "LANE5_YN ", LANE5_YN)
            Call Put_Ini("System Config", "LANE5_INOUT ", LANE5_Inout)
            Call Put_Ini("System Config", "LANE5_Name ", LANE5_Name)
            
        Case 5
            If chk_UseYN(5).value = "1" Then
                LANE6_YN = "Y"
            Else
                LANE6_YN = "N"
            End If
            LANE6_Inout = cmb_Inout(5).text
            LANE6_Name = Trim(txt_GateName(5))
            Call Put_Ini("System Config", "LANE6_YN ", LANE6_YN)
            Call Put_Ini("System Config", "LANE6_INOUT ", LANE6_Inout)
            Call Put_Ini("System Config", "LANE6_Name ", LANE6_Name)
    End Select

End Sub

Private Sub Command8_Click()
    ListData.Clear
End Sub

Private Sub Text_HomeNet_IP_Change()
    If (HomeNet_IP <> Text_HomeNet_IP.text) Then
        cmd_HomeTest.Enabled = False
    Else
        cmd_HomeTest.Enabled = True
    End If
End Sub

