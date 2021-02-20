VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmG6_23 
   Appearance      =   0  '평면
   BackColor       =   &H00404040&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   15420
   ClientLeft      =   -7080
   ClientTop       =   2550
   ClientWidth     =   28785
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmG6_23.frx":0000
   ScaleHeight     =   1028
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   1919
   Begin VB.TextBox txt_CarNo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1065
      TabIndex        =   99
      Text            =   "25구5401"
      Top             =   915
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2985
      TabIndex        =   98
      Top             =   945
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3465
      TabIndex        =   97
      Top             =   945
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   3945
      TabIndex        =   96
      Top             =   945
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   4425
      TabIndex        =   95
      Top             =   945
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   4905
      TabIndex        =   94
      Top             =   945
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   5385
      TabIndex        =   93
      Top             =   945
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   " 차단기 자동열림(방문차량) "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5280
      TabIndex        =   86
      ToolTipText     =   "방문차량(미등록차량) 차단기 열림"
      Top             =   0
      Width           =   6255
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인5"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   2130
         TabIndex        =   92
         Top             =   510
         Width           =   1935
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인2"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   2130
         TabIndex        =   91
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인1"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   90
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인3"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   4140
         TabIndex        =   89
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인4"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   88
         Top             =   510
         Width           =   1935
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인6"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4140
         TabIndex        =   87
         Top             =   510
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "차단기 자동열림(영업차량)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   855
      Left            =   11580
      TabIndex        =   79
      ToolTipText     =   "영업용차량(택배,화물) 차단기 열림"
      Top             =   0
      Width           =   6285
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인6"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   240
         Index           =   5
         Left            =   4230
         TabIndex        =   85
         Top             =   510
         Width           =   1935
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인5"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   240
         Index           =   4
         Left            =   2160
         TabIndex        =   84
         Top             =   510
         Width           =   1935
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인4"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   83
         Top             =   510
         Width           =   1935
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인3"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   240
         Index           =   2
         Left            =   4230
         TabIndex        =   82
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인2"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   81
         Top             =   270
         Width           =   1935
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인1"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   270
         Width           =   1935
      End
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   5
      Left            =   26235
      TabIndex        =   36
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "환경설정"
      ForeColor       =   16777215
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
      Picture         =   "FrmG6_23.frx":3E8CF
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   0
      Left            =   21435
      TabIndex        =   37
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "입출차조회"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":3EC20
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   1
      Left            =   22635
      TabIndex        =   38
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "보호해제"
      ForeColor       =   65280
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
      Picture         =   "FrmG6_23.frx":3EF71
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   6
      Left            =   27435
      TabIndex        =   39
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "시스템 종료"
      ForeColor       =   16777215
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
      Picture         =   "FrmG6_23.frx":3F2C2
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   2
      Left            =   23835
      TabIndex        =   40
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "정기권 관리"
      ForeColor       =   16777215
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
      Picture         =   "FrmG6_23.frx":3F613
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   3
      Left            =   28980
      TabIndex        =   41
      Top             =   2070
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "정기권 이력"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":3F964
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   4
      Left            =   25035
      TabIndex        =   42
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "근무자 관리"
      ForeColor       =   16777215
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
      Picture         =   "FrmG6_23.frx":3FCB5
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   435
      Left            =   25020
      TabIndex        =   76
      Top             =   420
      Width           =   4965
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Height          =   7125
      Index           =   5
      Left            =   19215
      TabIndex        =   71
      Top             =   8265
      Width           =   9495
      Begin VB.CommandButton NoWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "자리비움"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   8010
         Style           =   1  '그래픽
         TabIndex        =   106
         ToolTipText     =   "[자리비움]은 모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
         Top             =   510
         Width           =   1335
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   5
         Left            =   2580
         Picture         =   "FrmG6_23.frx":40006
         Top             =   2130
         Width           =   3690
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   5
         Left            =   5880
         TabIndex        =   75
         Top             =   90
         Width           =   3405
      End
      Begin VB.Label lbl_carno 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "서울00가1234"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   5
         Left            =   6270
         TabIndex        =   74
         Top             =   6330
         Width           =   3165
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   450
         Index           =   5
         Left            =   180
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lbl_GN 
         BackStyle       =   0  '투명
         Caption         =   "Lbl_Name"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   5
         Left            =   150
         TabIndex        =   73
         Top             =   6450
         Width           =   3735
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "준비중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   780
         Index           =   5
         Left            =   2520
         TabIndex        =   72
         Top             =   6330
         Width           =   3585
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   7140
         Index           =   5
         Left            =   0
         Picture         =   "FrmG6_23.frx":58684
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Height          =   7125
      Index           =   4
      Left            =   9615
      TabIndex        =   66
      Top             =   8265
      Width           =   9495
      Begin VB.CommandButton NoWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "자리비움"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   8010
         Style           =   1  '그래픽
         TabIndex        =   105
         ToolTipText     =   "[자리비움]은 모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
         Top             =   510
         Width           =   1335
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   4
         Left            =   2550
         Picture         =   "FrmG6_23.frx":7DDB7
         Top             =   2160
         Width           =   3690
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "준비중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   780
         Index           =   4
         Left            =   2520
         TabIndex        =   70
         Top             =   6330
         Width           =   3585
      End
      Begin VB.Label lbl_GN 
         BackStyle       =   0  '투명
         Caption         =   "Lbl_Name"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   4
         Left            =   150
         TabIndex        =   69
         Top             =   6450
         Width           =   3735
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   450
         Index           =   4
         Left            =   180
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lbl_carno 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "서울00가1234"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   4
         Left            =   6270
         TabIndex        =   68
         Top             =   6330
         Width           =   3165
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   4
         Left            =   5880
         TabIndex        =   67
         Top             =   90
         Width           =   3405
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   7140
         Index           =   4
         Left            =   0
         Picture         =   "FrmG6_23.frx":96435
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Height          =   7125
      Index           =   3
      Left            =   105
      TabIndex        =   61
      Top             =   8265
      Width           =   9495
      Begin VB.CommandButton NoWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "근무중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   8010
         Style           =   1  '그래픽
         TabIndex        =   104
         ToolTipText     =   "[자리비움]은 모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
         Top             =   510
         Width           =   1335
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   3
         Left            =   2550
         Picture         =   "FrmG6_23.frx":BBB68
         Top             =   2130
         Width           =   3690
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   5880
         TabIndex        =   65
         Top             =   90
         Width           =   3405
      End
      Begin VB.Label lbl_carno 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "서울00가1234"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   3
         Left            =   6270
         TabIndex        =   64
         Top             =   6330
         Width           =   3165
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   450
         Index           =   3
         Left            =   180
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lbl_GN 
         BackStyle       =   0  '투명
         Caption         =   "Lbl_Name"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   3
         Left            =   150
         TabIndex        =   63
         Top             =   6450
         Width           =   3735
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "준비중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   780
         Index           =   3
         Left            =   2520
         TabIndex        =   62
         Top             =   6330
         Width           =   3585
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   7140
         Index           =   3
         Left            =   0
         Picture         =   "FrmG6_23.frx":D41E6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Height          =   7125
      Index           =   2
      Left            =   19215
      TabIndex        =   56
      Top             =   1050
      Width           =   9495
      Begin VB.CommandButton NoWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "자리비움"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   8010
         Style           =   1  '그래픽
         TabIndex        =   103
         ToolTipText     =   "[자리비움]은 모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
         Top             =   510
         Width           =   1335
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   2
         Left            =   2580
         Picture         =   "FrmG6_23.frx":F9919
         Top             =   2130
         Width           =   3690
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "준비중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   780
         Index           =   2
         Left            =   2520
         TabIndex        =   60
         Top             =   6330
         Width           =   3585
      End
      Begin VB.Label lbl_GN 
         BackStyle       =   0  '투명
         Caption         =   "Lbl_Name"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   2
         Left            =   150
         TabIndex        =   59
         Top             =   6450
         Width           =   3735
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   450
         Index           =   2
         Left            =   180
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lbl_carno 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "서울00가1234"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   2
         Left            =   6270
         TabIndex        =   58
         Top             =   6330
         Width           =   3165
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   5880
         TabIndex        =   57
         Top             =   90
         Width           =   3405
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   7140
         Index           =   2
         Left            =   0
         Picture         =   "FrmG6_23.frx":111F97
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Height          =   7125
      Index           =   1
      Left            =   9630
      TabIndex        =   51
      Top             =   1050
      Width           =   9495
      Begin VB.CommandButton NoWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "근무중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   8010
         Style           =   1  '그래픽
         TabIndex        =   102
         ToolTipText     =   "[자리비움]은 모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
         Top             =   510
         Width           =   1335
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   1
         Left            =   2580
         Picture         =   "FrmG6_23.frx":1376CA
         Top             =   2130
         Width           =   3690
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   5880
         TabIndex        =   55
         Top             =   90
         Width           =   3405
      End
      Begin VB.Label lbl_carno 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "서울00가1234"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   1
         Left            =   6270
         TabIndex        =   54
         Top             =   6330
         Width           =   3165
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   450
         Index           =   1
         Left            =   180
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lbl_GN 
         BackStyle       =   0  '투명
         Caption         =   "Lbl_Name"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   1
         Left            =   150
         TabIndex        =   53
         Top             =   6450
         Width           =   3735
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "준비중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   780
         Index           =   1
         Left            =   2520
         TabIndex        =   52
         Top             =   6330
         Width           =   3585
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   7140
         Index           =   1
         Left            =   0
         Picture         =   "FrmG6_23.frx":14FD48
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '없음
      Height          =   7125
      Index           =   0
      Left            =   90
      TabIndex        =   46
      Top             =   1050
      Width           =   9495
      Begin VB.CommandButton NoWork 
         BackColor       =   &H00E0E0E0&
         Caption         =   "근무중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   8010
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   101
         ToolTipText     =   "[자리비움]은 모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
         Top             =   510
         Width           =   1335
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   0
         Left            =   2580
         Picture         =   "FrmG6_23.frx":17547B
         Top             =   2160
         Width           =   3690
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "준비중"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   26.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   780
         Index           =   0
         Left            =   2520
         TabIndex        =   50
         Top             =   6330
         Width           =   3585
      End
      Begin VB.Label lbl_GN 
         BackStyle       =   0  '투명
         Caption         =   "Lbl_Name"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   0
         Left            =   150
         TabIndex        =   49
         Top             =   6450
         Width           =   3735
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         Height          =   450
         Index           =   0
         Left            =   180
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lbl_carno 
         BackColor       =   &H00404040&
         BackStyle       =   0  '투명
         Caption         =   "서울00가1234"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   0
         Left            =   6270
         TabIndex        =   48
         Top             =   6330
         Width           =   3165
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   15.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   0
         Left            =   5880
         TabIndex        =   47
         Top             =   90
         Width           =   3405
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   7140
         Index           =   0
         Left            =   0
         Picture         =   "FrmG6_23.frx":18DAF9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9480
      End
   End
   Begin VB.CheckBox Chk_FreePass_old 
      BackColor       =   &H00000000&
      Caption         =   "출 구"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   31575
      TabIndex        =   35
      Top             =   3510
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CheckBox Chk_FreePass_old 
      BackColor       =   &H00000000&
      Caption         =   "입 구"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   30630
      TabIndex        =   34
      Top             =   3510
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CheckBox chk_Taxi_old 
      BackColor       =   &H00000000&
      Caption         =   "영업용 차량 자동 통과 ( 택시, 택배 )"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   32520
      TabIndex        =   33
      Top             =   3570
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.CommandButton Command3 
      Caption         =   "출차"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   35865
      TabIndex        =   16
      Top             =   180
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "입차"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   34875
      TabIndex        =   15
      Top             =   180
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   42990
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   645
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   510
      IMEMode         =   10  '한글 
      Left            =   32280
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30630
      Top             =   270
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   10  '한글 
      Left            =   40035
      TabIndex        =   0
      Top             =   630
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   27150
      TabIndex        =   4
      Top             =   15540
      Width           =   1155
   End
   Begin VB.ListBox ListView1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   90
      TabIndex        =   1
      Top             =   15480
      Width           =   19035
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   3
      Left            =   33420
      TabIndex        =   3
      Top             =   2715
      Visible         =   0   'False
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "종 료"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B322C
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   0
      Left            =   30390
      TabIndex        =   2
      Top             =   2715
      Visible         =   0   'False
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "입출차내역"
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
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B357D
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   31155
      Top             =   270
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   7
      Left            =   31905
      TabIndex        =   32
      Top             =   2715
      Visible         =   0   'False
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "환경설정"
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
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B38CE
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   55
      Left            =   30390
      TabIndex        =   45
      Top             =   2070
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "일괄등록"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B3C1F
   End
   Begin VB.PictureBox Server1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   29040
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   110
      Top             =   4305
      Width           =   1000
   End
   Begin MSWinsockLib.Winsock APS_Winsock 
      Left            =   30870
      Top             =   930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Remote_Winsock 
      Left            =   31290
      Top             =   930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   7
      Left            =   19035
      TabIndex        =   77
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "무인정산기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B3F70
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   8
      Left            =   28980
      TabIndex        =   107
      Top             =   2745
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "결제내역"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B42C1
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   9
      Left            =   28980
      TabIndex        =   108
      Top             =   1440
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "방문객관리"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "FrmG6_23.frx":1B4612
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   555
      Index           =   10
      Left            =   20235
      TabIndex        =   109
      Top             =   300
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "방문예약"
      ForeColor       =   16777215
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
      Picture         =   "FrmG6_23.frx":1B4963
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "주차관제 시스템"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   6
      Left            =   315
      TabIndex        =   44
      Top             =   105
      Width           =   3060
   End
   Begin VB.Label lbl_ParkFull 
      BackStyle       =   0  '투명
      Caption         =   "만차현황 : Now/Full"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   28980
      TabIndex        =   100
      Top             =   3705
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label LblDBInfo 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "DB오류 메시지"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   21930
      TabIndex        =   78
      Top             =   60
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Shape ShapeCamera 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   165
      Index           =   5
      Left            =   4050
      Top             =   555
      Width           =   360
   End
   Begin VB.Shape ShapeCamera 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   165
      Index           =   4
      Left            =   3555
      Top             =   555
      Width           =   360
   End
   Begin VB.Shape ShapeCamera 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   165
      Index           =   0
      Left            =   1635
      Top             =   555
      Width           =   360
   End
   Begin VB.Shape ShapeCamera 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   165
      Index           =   1
      Left            =   2115
      Top             =   555
      Width           =   360
   End
   Begin VB.Shape ShapeCamera 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   165
      Index           =   2
      Left            =   2610
      Top             =   555
      Width           =   360
   End
   Begin VB.Shape ShapeCamera 
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   165
      Index           =   3
      Left            =   3075
      Top             =   555
      Width           =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "카메라 상태 "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   300
      TabIndex        =   43
      Top             =   495
      Width           =   1590
   End
   Begin VB.Label LblRecStat 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   630
      Index           =   1
      Left            =   47610
      TabIndex        =   31
      Top             =   780
      Width           =   2760
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   46830
      TabIndex        =   30
      Top             =   8760
      Width           =   3585
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   46830
      TabIndex        =   29
      Top             =   9255
      Width           =   3615
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   46830
      TabIndex        =   28
      Top             =   9795
      Width           =   3585
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   46830
      TabIndex        =   27
      Top             =   10260
      Width           =   3615
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   46830
      TabIndex        =   26
      Top             =   10755
      Width           =   3615
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   46830
      TabIndex        =   25
      Top             =   11220
      Width           =   3615
   End
   Begin VB.Label lbl_info_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   46830
      TabIndex        =   24
      Top             =   11730
      Width           =   3615
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   44850
      TabIndex        =   23
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   44850
      TabIndex        =   22
      Top             =   9255
      Width           =   1815
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   44850
      TabIndex        =   21
      Top             =   9765
      Width           =   1815
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   44850
      TabIndex        =   20
      Top             =   10245
      Width           =   1815
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   44850
      TabIndex        =   19
      Top             =   10755
      Width           =   1815
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   44850
      TabIndex        =   18
      Top             =   11220
      Width           =   1815
   End
   Begin VB.Label lbl_title_Out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_Out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   44850
      TabIndex        =   17
      Top             =   11730
      Width           =   1800
   End
   Begin VB.Label LblGubun 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   14
      Top             =   5430
      Width           =   60
   End
   Begin VB.Label LblTime 
      BackColor       =   &H00000000&
      Caption         =   "현재시간"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   19020
      TabIndex        =   11
      Top             =   60
      Width           =   4800
   End
   Begin VB.Label LblName 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   10
      Top             =   3735
      Width           =   60
   End
   Begin VB.Label LblCar 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   9
      Top             =   3405
      Width           =   60
   End
   Begin VB.Label LblId 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   8
      Top             =   4080
      Width           =   60
   End
   Begin VB.Label LblCarType 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   7
      Top             =   4410
      Width           =   60
   End
   Begin VB.Label LblTel 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   6
      Top             =   4755
      Width           =   60
   End
   Begin VB.Label LblDate 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   40125
      TabIndex        =   5
      Top             =   5085
      Width           =   60
   End
   Begin VB.Menu MnuLoop 
      Caption         =   "LoopBack"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu MnuIlDelete 
      Caption         =   "일반권삭제"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu MnuIlOutDelete 
      Caption         =   "일반출차 삭제"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu MnuJungDelete 
      Caption         =   "정기권삭제"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FrmG6_23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Chk_FreePass_Click(Index As Integer)

    Dim sLaneName, sGuestUse, sAutoMode As String
    
    Select Case Index
        
        Case 0
                If Chk_FreePass(0).value = 1 Then
                    Glo_FreePassLane1_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane1_YN", "Y")
                Else
                    Glo_FreePassLane1_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane1_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane1_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane1_YN)
                End If
                
                sLaneName = LANE1_Name
        Case 1
                If Chk_FreePass(1).value = 1 Then
                    Glo_FreePassLane2_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane2_YN", "Y")
                Else
                    Glo_FreePassLane2_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane2_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane2_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane2_YN)
                End If
                
                sLaneName = LANE2_Name
        Case 2
                If Chk_FreePass(2).value = 1 Then
                    Glo_FreePassLane3_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane3_YN", "Y")
                Else
                    Glo_FreePassLane3_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane3_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane3_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane3_YN)
                End If
                
                sLaneName = LANE3_Name
        Case 3
                If Chk_FreePass(3).value = 1 Then
                    Glo_FreePassLane4_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane4_YN", "Y")
                Else
                    Glo_FreePassLane4_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane4_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane4_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane4_YN)
                End If
                
                sLaneName = LANE4_Name
        Case 4
                If Chk_FreePass(4).value = 1 Then
                    Glo_FreePassLane5_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane5_YN", "Y")
                Else
                    Glo_FreePassLane5_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane5_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane5_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane5_YN)
                End If
                
                sLaneName = LANE5_Name
        Case 5
                If Chk_FreePass(5).value = 1 Then
                    Glo_FreePassLane6_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane6_YN", "Y")
                Else
                    Glo_FreePassLane6_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane6_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane6_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane6_YN)
                End If
                
                sLaneName = LANE6_Name
    End Select

    '방문객 자동 처리유무
    If (Chk_FreePass(Index).value = 1) Then
        sGuestUse = "(자동처리)"
        sAutoMode = "Y"
    Else
        sGuestUse = ""
        sAutoMode = "N"
    End If
    If (Not Glo_FrmGuest(Index) Is Nothing) Then '만들어져 있다면
        'Call Glo_FrmGuest(Index).SetGuestName(sLaneName & sGuestUse)
        Call Glo_FrmGuest(Index).SetAutoMode(sAutoMode, sLaneName & sGuestUse)
    End If
    
    
    Dim sLog As String
    If (sAutoMode = "Y") Then
        sLog = "Lane" & Index + 1 & ":" & "방문차량자동열림"
    Else
        sLog = "Lane" & Index + 1 & ":" & "방문차량자동열림 해제"
    End If
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('차단기자동열림', 'HOST','" & sLog & "','방문차량'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
End Sub

Public Sub Chk_FreePassEnable(Index As Integer, bVal As Boolean)
    If (Index < Glo_Screen_No) Then
        Chk_FreePass(Index).Enabled = bVal
    End If
End Sub


Private Sub chk_Taxi_Click(Index As Integer)
    Select Case Index
        Case 0
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI1_YN = "Y"
            Else
                Glo_TAXI1_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI1_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI1_YN)
            End If
            
        Case 1
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI2_YN = "Y"
            Else
                Glo_TAXI2_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI2_YN", Glo_TAXI2_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI2_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI2_YN)
            End If
            
        Case 2
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI3_YN = "Y"
            Else
                Glo_TAXI3_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI3_YN", Glo_TAXI3_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI3_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI3_YN)
            End If
            
        Case 3
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI4_YN = "Y"
            Else
                Glo_TAXI4_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI4_YN", Glo_TAXI4_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI4_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI4_YN)
            End If
            
        Case 4
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI5_YN = "Y"
            Else
                Glo_TAXI5_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI5_YN", Glo_TAXI5_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI5_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI5_YN)
            End If
            
        Case 5
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI6_YN = "Y"
            Else
                Glo_TAXI6_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI6_YN", Glo_TAXI6_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI6_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI6_YN)
            End If
    End Select
    
    
    '영업차량 차단기 자동열림
    Dim sLog As String
    If (chk_Taxi(Index).value = 1) Then
        sLog = "Lane" & Index + 1 & ":" & "영업차량자동열림"
    Else
        sLog = "Lane" & Index + 1 & ":" & "영업차량자동열림 해제"
    End If
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('차단기자동열림', 'HOST','" & sLog & "','영업차량'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"

End Sub

Public Sub Set_GateName(ByVal iIndex As Integer, ByVal sGateName As String)
    If (iIndex < Glo_Screen_No) Then
        lbl_GN(iIndex).Caption = sGateName
        Chk_FreePass(iIndex).Caption = sGateName
    End If
End Sub

Public Sub ReDraw(sKind As String, iIndex As Integer, iValue As Integer)
    If sKind = "FreePass" Then
        Chk_FreePass(iIndex).value = iValue
        Call Chk_FreePass_Click(iIndex)
    ElseIf sKind = "Taxi" Then
        chk_Taxi(iIndex).value = iValue
        Call chk_Taxi_Click(iIndex)
    ElseIf sKind = "NOWORK" Then
        'chk_NoWork(iIndex).value = iValue
        If (iValue = 1) Then
            NoWork(iIndex).Caption = "근무중"
        Else
            NoWork(iIndex).Caption = "자리비움"
        End If
        Call NoWork_Click(iIndex)
    End If
End Sub

Private Sub cmd_Menu_Click(Index As Integer)
    Dim i As Integer

    Call GuestForm_WindowState(vbMinimized)

    Me.MousePointer = 11
    Select Case Index
        Case 6
            Call DataLogger("[HOST Button]    " & "주차관제 시스템 종료!!!")
            Unload Me
        Case 5
            If (Glo_Login_GUBUN = "총괄관리자") Then
                FrmTcpServer.Show 0
                Me.MousePointer = 0
                Call DataLogger("[HOST Button]    " & "TCP Server 화면 접근")
            'ElseIf (Glo_Login_GUBUN = "관리자") Then
            Else
                FrmTcpServer2.Show 0
                Me.MousePointer = 0
                Call DataLogger("[HOST Button]    " & "TCP Server2 화면 접근")
            End If
        Case 0
             'FrmInOut.Show 1
             FrmInOut.Show 0
             Me.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "입출차 보고서 화면 접근")
        Case 1
            If (cmd_menu(1).Caption = "보호모드") Then

                Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
                Call ProtectMainMenuButton6Form(Me)

                Glo_Login_ID = ""
                Glo_Login_PW = ""
                Glo_Login_GUBUN = ""
                Put_Ini "System Config", "보호모드", "True"

            Else
                Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
                frmLogin.Show 1
            End If
            Me.MousePointer = 0

        Case 2
             'FrmReg.Show 1
             FrmReg.Show 0
             Me.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "정기권관리 화면 접근")
'        Case 55
'             FrmCSV.Show 1
'             Me.MousePointer = 0
'             Call DataLogger("[HOST Button]    " & "일괄등록 화면 접근")
        Case 3
             'FrmRegHistory.Show 1
             FrmRegHistory.Show 0
             Me.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "정기권 이력 화면 접근")
        Case 4
            'FrmId.Show 1
            FrmId.Show 0
            Me.MousePointer = 0
            Call DataLogger("[HOST Button]    " & "아이디 관리 화면 접근")
'''        Case 8
'''            Me.MousePointer = 0
'''            frmResult.Show 0
'''            Call DataLogger("[HOST Button]    " & "결제내역 화면 접근")
'''
        Case 7
            Me.MousePointer = 0
            'FrmAccnt.Show 1
            If (cmd_menu(Index).Caption = "무인정산기") Then
                'FrmAccnt.Show 1
                FrmAccnt.Show 0
            ElseIf (cmd_menu(Index).Caption = "결제내역") Then
                'frmResult.Show 1
                frmResult.Show 0
            End If
            Call DataLogger("[HOST Button]    " & "무인정산기 관리 화면 접근")

        Case 8
            Me.MousePointer = 0
            'frmResult.Show 1
            frmResult.Show 0
            Call DataLogger("[HOST Button]    " & "결제내역 화면 접근")

        Case 9
            Me.MousePointer = 0
            'FrmGuestLog.Show 1
            FrmGuestLog.Show 0
            Call DataLogger("[HOST Button]    " & "방문객내역 화면 접근")

        Case 10  '방문차량 사전방문
            Me.MousePointer = 1
            FrmGuestRegLog.Show 0
            Call DataLogger("[HOST Button]    " & "방문예약 화면 접근")
            Exit Sub
    End Select
End Sub


Private Sub Form_Load()
    Dim Ret As Integer
    Dim fso As New FileSystemObject
    Dim i As Integer
    Dim Reg_Addr As String
    Dim Reg_HDD As String
    Dim sGuestUse, sAutoMode As String
    
'   Me.Caption = Me.Caption & "  " & "Version " & App.Major & "." & App.Minor & "." & App.Revision
    IniFileName$ = App.Path & "\Winpark.ini"
    Report_Path_Name$ = App.Path & "\Data\"
    Doc_Path_Name$ = App.Path & "\Doc\"
    
    If App.PrevInstance = True Then
        End
    End If
    

    Left = (Screen.width - width) / 2
    Top = 0
    
    
    
    If (Glo_ParkFull_YN = "Y") Then
        Call ParkFull_Show
    End If
    
    
    
    If (Glo_TestMode = "Y") Then
        txt_CarNo.Enabled = True
        Lane(0).Enabled = True
        Lane(1).Enabled = True
        Lane(2).Enabled = True
        Lane(3).Enabled = True
        Lane(4).Enabled = True
        Lane(5).Enabled = True
        txt_CarNo.Visible = True
        Lane(0).Visible = True
        Lane(1).Visible = True
        Lane(2).Visible = True
        Lane(3).Visible = True
        Lane(4).Visible = True
        Lane(5).Visible = True
    Else
        txt_CarNo.Enabled = False
        Lane(0).Enabled = False
        Lane(1).Enabled = False
        Lane(2).Enabled = False
        Lane(3).Enabled = False
        Lane(4).Enabled = False
        Lane(5).Enabled = False
        txt_CarNo.Visible = False
        Lane(0).Visible = False
        Lane(1).Visible = False
        Lane(2).Visible = False
        Lane(3).Visible = False
        Lane(4).Visible = False
        Lane(5).Visible = False
    End If


    For i = 0 To 5
        ImageIn(i).Picture = LoadPicture(App.Path & "\NoCar.jpg")
        lbl_GN(0).Caption = ""
        lbl_carno(i).Caption = ""
        lbl_time_now(i).Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")
        Shp_Rec(i).Visible = False
        
        Chk_FreePass(i).Caption = ""
    Next i


    lbl_GN(0).Caption = Trim(LANE1_Name)
    lbl_GN(1).Caption = Trim(LANE2_Name)
    lbl_GN(2).Caption = Trim(LANE3_Name)
    lbl_GN(3).Caption = Trim(LANE4_Name)
    lbl_GN(4).Caption = Trim(LANE5_Name)
    lbl_GN(5).Caption = Trim(LANE6_Name)

    
    
    
    
    
    
'''    If LANE1_YN = "Y" Then
'''        Chk_FreePass(0).Caption = LANE1_Name
'''        Call Chk_FreePassEnable(0, True)
'''    Else
'''        Chk_FreePass(0).Caption = "미사용"
'''        Call Chk_FreePassEnable(0, False)
'''    End If
'''    If LANE2_YN = "Y" Then
'''        Chk_FreePass(1).Caption = LANE2_Name
'''        Call Chk_FreePassEnable(1, True)
'''    Else
'''        Chk_FreePass(1).Caption = "미사용"
'''        Call Chk_FreePassEnable(1, False)
'''    End If
'''    If LANE3_YN = "Y" Then
'''        Chk_FreePass(2).Caption = LANE3_Name
'''        Call Chk_FreePassEnable(2, True)
'''    Else
'''        Chk_FreePass(2).Caption = "미사용"
'''        Call Chk_FreePassEnable(2, False)
'''    End If
'''    If LANE4_YN = "Y" Then
'''        Chk_FreePass(3).Caption = LANE4_Name
'''        Call Chk_FreePassEnable(3, True)
'''    Else
'''        Chk_FreePass(3).Caption = "미사용"
'''        Call Chk_FreePassEnable(3, False)
'''    End If
'''    If LANE5_YN = "Y" Then
'''        Chk_FreePass(4).Caption = LANE5_Name
'''        Call Chk_FreePassEnable(4, True)
'''    Else
'''        Chk_FreePass(4).Caption = "미사용"
'''        Call Chk_FreePassEnable(4, False)
'''    End If
'''    If LANE6_YN = "Y" Then
'''        Chk_FreePass(5).Caption = LANE6_Name
'''        Call Chk_FreePassEnable(5, True)
'''    Else
'''        Chk_FreePass(5).Caption = "미사용"
'''        Call Chk_FreePassEnable(5, False)
'''    End If
'''
'''
'''    If Glo_FreePassLane1_YN = "Y" Then
'''        Chk_FreePass(0).value = 1
'''    Else
'''        Chk_FreePass(0).value = 0
'''    End If
'''    If Glo_FreePassLane2_YN = "Y" Then
'''        Chk_FreePass(1).value = 1
'''    Else
'''        Chk_FreePass(1).value = 0
'''    End If
'''    If Glo_FreePassLane3_YN = "Y" Then
'''        Chk_FreePass(2).value = 1
'''    Else
'''        Chk_FreePass(2).value = 0
'''    End If
'''    If Glo_FreePassLane4_YN = "Y" Then
'''        Chk_FreePass(3).value = 1
'''    Else
'''        Chk_FreePass(3).value = 0
'''    End If
'''    If Glo_FreePassLane5_YN = "Y" Then
'''        Chk_FreePass(4).value = 1
'''    Else
'''        Chk_FreePass(4).value = 0
'''    End If
'''    If Glo_FreePassLane6_YN = "Y" Then
'''        Chk_FreePass(5).value = 1
'''    Else
'''        Chk_FreePass(5).value = 0
'''    End If

    ' 영업용차량 입출구 구분없애고, 레인별처리로 전환함 - 시작
    Call Chk_TaxiPassEnable(Me, LANE1_YN, Glo_TAXI1_YN, 0, LANE1_Name)
    Call Chk_TaxiPassEnable(Me, LANE2_YN, Glo_TAXI2_YN, 1, LANE2_Name)
    Call Chk_TaxiPassEnable(Me, LANE3_YN, Glo_TAXI3_YN, 2, LANE3_Name)
    Call Chk_TaxiPassEnable(Me, LANE4_YN, Glo_TAXI4_YN, 3, LANE4_Name)
    Call Chk_TaxiPassEnable(Me, LANE5_YN, Glo_TAXI5_YN, 4, LANE5_Name)
    Call Chk_TaxiPassEnable(Me, LANE6_YN, Glo_TAXI6_YN, 5, LANE6_Name)
    ' 영업용차량 입출구 구분없애고, 레인별처리로 전환함 - 끝
    
    ' 일반차량 입출구 구분없애고, 레인별처리로 전환함 - 시작
    Call Chk_NormalPassEnable(Me, LANE1_YN, Glo_FreePassLane1_YN, 0, LANE1_Name)
    Call Chk_NormalPassEnable(Me, LANE2_YN, Glo_FreePassLane2_YN, 1, LANE2_Name)
    Call Chk_NormalPassEnable(Me, LANE3_YN, Glo_FreePassLane3_YN, 2, LANE3_Name)
    Call Chk_NormalPassEnable(Me, LANE4_YN, Glo_FreePassLane4_YN, 3, LANE4_Name)
    Call Chk_NormalPassEnable(Me, LANE5_YN, Glo_FreePassLane5_YN, 4, LANE5_Name)
    Call Chk_NormalPassEnable(Me, LANE6_YN, Glo_FreePassLane6_YN, 5, LANE6_Name)
    ' 일반차량 입출구 구분없애고, 레인별처리로 전환함 - 끝
    
    
    If (Glo_Screen_No = 6) Then
        '방문객처리
        For i = 0 To MAX_LANE_COUNT
            If (Not Glo_FrmGuest(i) Is Nothing) Then '만들어져 있다면
                Unload Glo_FrmGuest(i)
                Set Glo_FrmGuest(i) = Nothing
            End If
        Next i
        
        If (LANE1_YN = "Y" And Glo_GUEST_LANE1_YN = "Y") Then
            Set Glo_FrmGuest(0) = New FormGuest1
            Glo_FrmGuest(0).Show 0
            Call Glo_FrmGuest(0).SetGateNo(0, Glo_Guest_Print_Model(0), Glo_Guest_Print_Port(0))
            
            If (Glo_FreePassLane1_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            Call Glo_FrmGuest(0).SetAutoMode(sAutoMode, LANE1_Name & sGuestUse)
             
        End If
        If (LANE2_YN = "Y" And Glo_GUEST_LANE2_YN = "Y") Then
            Set Glo_FrmGuest(1) = New FormGuest1
            Glo_FrmGuest(1).Show 0
            Call Glo_FrmGuest(1).SetGateNo(1, Glo_Guest_Print_Model(1), Glo_Guest_Print_Port(1))
            
            If (Glo_FreePassLane2_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            Call Glo_FrmGuest(1).SetAutoMode(sAutoMode, LANE2_Name & sGuestUse)
        End If
        
        If (LANE3_YN = "Y" And Glo_GUEST_LANE3_YN = "Y") Then
            Set Glo_FrmGuest(2) = New FormGuest1
            Glo_FrmGuest(2).Show 0
            Call Glo_FrmGuest(2).SetGateNo(2, Glo_Guest_Print_Model(2), Glo_Guest_Print_Port(2))
            
            If (Glo_FreePassLane3_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            Call Glo_FrmGuest(2).SetAutoMode(sAutoMode, LANE3_Name & sGuestUse)
        End If
        
        If (LANE4_YN = "Y" And Glo_GUEST_LANE4_YN = "Y") Then
            Set Glo_FrmGuest(3) = New FormGuest1
            Glo_FrmGuest(3).Show 0
            Call Glo_FrmGuest(3).SetGateNo(3, Glo_Guest_Print_Model(3), Glo_Guest_Print_Port(3))
            
            If (Glo_FreePassLane4_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            Call Glo_FrmGuest(3).SetAutoMode(sAutoMode, LANE4_Name & sGuestUse)
            
        End If
        If (Glo_GUEST_LANE5_YN = "Y") Then
            Set Glo_FrmGuest(4) = New FormGuest1
            Glo_FrmGuest(4).Show 0
            Call Glo_FrmGuest(4).SetGateNo(4, Glo_Guest_Print_Model(4), Glo_Guest_Print_Port(4))
            
            If (Glo_FreePassLane5_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            Call Glo_FrmGuest(4).SetAutoMode(sAutoMode, LANE5_Name & sGuestUse)
        End If
        If (Glo_GUEST_LANE6_YN = "Y") Then
            Set Glo_FrmGuest(5) = New FormGuest1
            Glo_FrmGuest(5).Show 0
            Call Glo_FrmGuest(5).SetGateNo(5, Glo_Guest_Print_Model(5), Glo_Guest_Print_Port(5))
            
            If (Glo_FreePassLane6_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            Call Glo_FrmGuest(5).SetAutoMode(sAutoMode, LANE6_Name & sGuestUse)
        End If
    End If
    
    For i = 0 To 5
        NoWork(i).Caption = "근무중"
        NoWork(i).ToolTipText = "자리비울경우 버튼을 눌러주세요"
    Next i
    
    
    
'''    If (Glo_Login_ID = "") Then
'''        cmd_menu(0).Enabled = False
'''        cmd_menu(1).Enabled = False
'''        'cmd_menu(2).Enabled = False
'''        'cmd_Menu(3).Enabled = False '종료
'''        cmd_menu(4).Enabled = False
'''        cmd_menu(5).Enabled = False '일괄등록
'''        cmd_menu(6).Enabled = False
'''        cmd_menu(7).Enabled = False
'''    Else
'''        Call frmLogin.ShowMenu(Glo_Login_ID, Glo_Login_PW)
'''    End If
    Call ProtectMainMenuButton6Form(Me)
    
    Call ShowTitlebarSiteCode
    
    
    Timer1.Enabled = True
    FrmTcpServer.Hide

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msg, Style, Title, Response
Dim Ret As Boolean
'msg = "차번인식기와의 접속이 해제되므로     " & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "시스템 운영이 중단됩니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "종료하시겠습니까?"
msg = "프로그램을 종료하시겠습니까?         "
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Parking Manager™ System"
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then
    Call DataLogger("[HOST Button]    " & "프로그램 종료")
    Call DataBaseClose(adoConn)
    
    Unload FrmTcpServer
    Unload FrmAccnt
    Unload FormIPCamera
    'Unload FormIPCameraPlayer
    Dim i As Integer
    For i = 0 To UBound(Glo_FrmIPCameraPlayer)
        If (Not Glo_FrmIPCameraPlayer(i) Is Nothing) Then
            Unload Glo_FrmIPCameraPlayer(i)
        End If
    Next i
    Unload frmApsCmd
    Unload FrmCSV
    Unload FrmDeviceCFG
    Unload FrmExtend
    'Unload FrmG1
    'Unload FrmG4Mini
    'Unload FrmG6_23
    Unload FrmId
    Unload FrmImg
    Unload FrmInOut
    Unload frmLogin
    'Unload FrmPhoto
    Unload FrmReg
    Unload FrmRegHistory
    Unload frmResult
    Unload frmSplash
    Unload FrmTcpServer2
    'Unload Jung
    Unload MBox
    Unload Msg_Box
    Unload Pwd
    
    
    If (Glo_Screen_No = 6) Then
        If (Glo_GUEST_LANE1_YN = "Y") Then
            If (Not Glo_FrmGuest(0) Is Nothing) Then
                Unload Glo_FrmGuest(0)
                Set Glo_FrmGuest(0) = Nothing
            End If
        End If
        If (Glo_GUEST_LANE2_YN = "Y") Then
            If (Not Glo_FrmGuest(1) Is Nothing) Then
                Unload Glo_FrmGuest(1)
                Set Glo_FrmGuest(1) = Nothing
            End If
        End If
        If (Glo_GUEST_LANE3_YN = "Y") Then
            If (Not Glo_FrmGuest(2) Is Nothing) Then
                Unload Glo_FrmGuest(2)
                Set Glo_FrmGuest(2) = Nothing
            End If
        End If
        If (Glo_GUEST_LANE4_YN = "Y") Then
            If (Not Glo_FrmGuest(3) Is Nothing) Then
                Unload Glo_FrmGuest(3)
                Set Glo_FrmGuest(3) = Nothing
            End If
        End If
        If (Glo_GUEST_LANE5_YN = "Y") Then
            If (Not Glo_FrmGuest(4) Is Nothing) Then
                Unload Glo_FrmGuest(4)
                Set Glo_FrmGuest(4) = Nothing
            End If
        End If
        If (Glo_GUEST_LANE6_YN = "Y") Then
            If (Not Glo_FrmGuest(5) Is Nothing) Then
                Unload Glo_FrmGuest(5)
                Set Glo_FrmGuest(5) = Nothing
            End If
        End If
    End If
    
    
    Call Unhook
    End
'    SetWindowLong hwnd, GWL_WNDPROC, g_addProcOld
'    Timer1.Enabled = False
'    Call Err_doc("호스트 : " & "프로그램 정상적으로 종료")
'    End
End If
Me.MousePointer = 0
Cancel = True
End Sub

Private Sub Lane_Click(Index As Integer)
    Dim CarnoEnc As String
    Dim carno As String
    
    carno = CStr(Index) & "_" & txt_CarNo & "_\\localhost\Lane1\image\20161229\20161229171357960_25구5401.jpg"
    CarnoEnc = EncodeNDE01(carno, "www.jawootek.com")
    carno = DecodeNDE01(CarnoEnc, "www.jawootek.com")

    If (Glo_TestMode = "Y") Then
        Call UDP_Proc(carno)
    End If
End Sub

Private Sub NoWork_Click(Index As Integer)
    Dim sSendValue As String
    Dim sLaneName, sGuestUse, sAutoMode As String
    
    If (NoWork(Index).Caption = "근무중") Then
        NoWork(Index).Caption = "자리비움"
        NoWork(Index).ToolTipText = "근무중이면 버튼을 눌러주세요"
        sSendValue = "Y"
        chk_Taxi(Index).Enabled = False
        Chk_FreePass(Index).Enabled = False
        NoWork(Index).BackColor = &HFF&
    Else
        NoWork(Index).Caption = "근무중"
        NoWork(Index).ToolTipText = "자리비울경우 버튼을 눌러주세요"
        sSendValue = "N"
        chk_Taxi(Index).Enabled = True
        Chk_FreePass(Index).Enabled = True
        NoWork(Index).BackColor = &HE0E0E0
    End If
    
    Select Case Index
        Case 0
            Glo_Lane1_NoWork = NoWork(Index).Caption
            sLaneName = LANE1_Name
        Case 1
            Glo_Lane2_NoWork = NoWork(Index).Caption
            sLaneName = LANE2_Name
        Case 2
            Glo_Lane3_NoWork = NoWork(Index).Caption
            sLaneName = LANE3_Name
        Case 3
            Glo_Lane4_NoWork = NoWork(Index).Caption
            sLaneName = LANE4_Name
        Case 4
            Glo_Lane5_NoWork = NoWork(Index).Caption
            sLaneName = LANE5_Name
        Case 5
            Glo_Lane6_NoWork = NoWork(Index).Caption
            sLaneName = LANE6_Name
    End Select
    
    
    
    '방문객 자동 처리유무
    If (NoWork(Index).Caption = "자리비움" Or Chk_FreePass(Index).value = 1) Then
        sGuestUse = "(자동처리)"
        sAutoMode = "Y"
    Else
        sGuestUse = ""
        sAutoMode = "N"
    End If
    If (Not Glo_FrmGuest(Index) Is Nothing) Then '만들어져 있다면
        'Call Glo_FrmGuest(Index).SetGuestName(sLaneName & sGuestUse)
        Call Glo_FrmGuest(Index).SetAutoMode(sAutoMode, sLaneName & sGuestUse)
    End If
    
    
    
    If (Glo_FreepassS_YN = "Y") Then
        FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_NOWORK_" & sSendValue)
        DataLogger ("FreePass Send : " & Index & "_NOWORK_" & sSendValue)
    End If
    
    Dim sLog As String
    sLog = "차단기 자동열림[자리비움] Lane:" & Index + 1 & ":" & NoWork(Index).Caption
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('자리비움제어', 'HOST','" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Dim qry As String
    Dim rs As Recordset
    Dim ECHO As ICMP_ECHO_REPLY
    Dim bQryResult As Boolean
    
    If (Glo_Certify = enumCertify.eCertTry And Glo_Cert_NoticeSDate < Format(Now, "yyyy-mm-dd")) Then
        LblTime(0).ForeColor = &HFF&
        LblTime(0).Caption = "[인증받으세요] " & "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")
    Else
        LblTime(0).ForeColor = &H0&
        LblTime(0).ToolTipText = ""
        LblTime(0).Caption = "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")
    End If

    If (Format(Now, "NNSS") = "0000") Then
        '게이트 카운트 초기화
    Else
    End If


    If (Abs(Glo_Mon_LastInTime - Timer) >= 5) Then
        Glo_MonStat_Lane(0) = "DEAD"
        Glo_MonStat_Lane(1) = "DEAD"
        Glo_MonStat_Lane(2) = "DEAD"
        Glo_MonStat_Lane(3) = "DEAD"
        Glo_MonStat_Lane(4) = "DEAD"
        Glo_MonStat_Lane(5) = "DEAD"
    End If
    
    If (LANE1_YN = "Y") Then
        If (Glo_Mon_Lane(0) = True) Then
            If Glo_MonStat_Lane(0) = "LIVE" Then
                Imgshutdown(0).Visible = False
                ShapeCamera(0).Visible = True
                ShapeCamera(0).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
            Else
                Imgshutdown(0).Visible = True
                ShapeCamera(0).Visible = True
                ShapeCamera(0).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
                'Call DataLogger("Lane1 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane1.exe")) Then
                Imgshutdown(0).Visible = False
                ShapeCamera(0).Visible = True
                ShapeCamera(0).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
            Else
                Imgshutdown(0).Visible = True
                ShapeCamera(0).Visible = True
                ShapeCamera(0).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
                'Call DataLogger("Lane1 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(0).Visible = False
        ShapeCamera(0).Visible = False
    End If
    
    If (LANE2_YN = "Y") Then
        If (Glo_Mon_Lane(1) = True) Then
            If Glo_MonStat_Lane(1) = "LIVE" Then
                Imgshutdown(1).Visible = False
                ShapeCamera(1).Visible = True
                ShapeCamera(1).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(1, "LIVE")
            Else
                Imgshutdown(1).Visible = True
                ShapeCamera(1).Visible = True
                ShapeCamera(1).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(1, "DEAD")
                'Call DataLogger("Lane2 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane2.exe")) Then
                Imgshutdown(1).Visible = False
                ShapeCamera(1).Visible = True
                ShapeCamera(1).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(1, "LIVE")
            Else
                Imgshutdown(1).Visible = True
                ShapeCamera(1).Visible = True
                ShapeCamera(1).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(1, "DEAD")
                'Call DataLogger("Lane2 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(1).Visible = False
        ShapeCamera(1).Visible = False
    End If

    If (LANE3_YN = "Y") Then
        If (Glo_Mon_Lane(2) = True) Then
            If Glo_MonStat_Lane(2) = "LIVE" Then
                Imgshutdown(2).Visible = False
                ShapeCamera(2).Visible = True
                ShapeCamera(2).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(2, "LIVE")
            Else
                Imgshutdown(2).Visible = True
                ShapeCamera(2).Visible = True
                ShapeCamera(2).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(2, "DEAD")
                'Call DataLogger("Lane3 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane3.exe")) Then
                Imgshutdown(2).Visible = False
                ShapeCamera(2).Visible = True
                ShapeCamera(2).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(2, "LIVE")
            Else
                Imgshutdown(2).Visible = True
                ShapeCamera(2).Visible = True
                ShapeCamera(2).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(2, "DEAD")
                'Call DataLogger("Lane3 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(2).Visible = False
        ShapeCamera(2).Visible = False
    End If
    
    If (LANE4_YN = "Y") Then
        If (Glo_Mon_Lane(3) = True) Then
            If Glo_MonStat_Lane(3) = "LIVE" Then
                Imgshutdown(3).Visible = False
                ShapeCamera(3).Visible = True
                ShapeCamera(3).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(3, "LIVE")
            Else
                Imgshutdown(3).Visible = True
                ShapeCamera(3).Visible = True
                ShapeCamera(3).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(3, "DEAD")
                'Call DataLogger("Lane4 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane4.exe")) Then
                Imgshutdown(3).Visible = False
                ShapeCamera(3).Visible = True
                ShapeCamera(3).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(3, "LIVE")
            Else
                Imgshutdown(3).Visible = True
                ShapeCamera(3).Visible = True
                ShapeCamera(3).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(3, "DEAD")
                'Call DataLogger("Lane4 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(3).Visible = False
        ShapeCamera(3).Visible = False
    End If
    
    If (LANE5_YN = "Y") Then
        If (Glo_Mon_Lane(4) = True) Then
            If Glo_MonStat_Lane(4) = "LIVE" Then
                Imgshutdown(4).Visible = False
                ShapeCamera(4).Visible = True
                ShapeCamera(4).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(4, "LIVE")
            Else
                Imgshutdown(4).Visible = True
                ShapeCamera(4).Visible = True
                ShapeCamera(4).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(4, "DEAD")
                'Call DataLogger("Lane5 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane5.exe")) Then
                Imgshutdown(4).Visible = False
                ShapeCamera(4).Visible = True
                ShapeCamera(4).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(4, "LIVE")
            Else
                Imgshutdown(4).Visible = True
                ShapeCamera(4).Visible = True
                ShapeCamera(4).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(4, "DEAD")
                'Call DataLogger("Lane5 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(4).Visible = False
        ShapeCamera(4).Visible = False
    End If
    
    If (LANE6_YN = "Y") Then
        If (Glo_Mon_Lane(5) = True) Then
            If Glo_MonStat_Lane(5) = "LIVE" Then
                Imgshutdown(5).Visible = False
                ShapeCamera(5).Visible = True
                ShapeCamera(5).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(5, "LIVE")
            Else
                Imgshutdown(5).Visible = True
                ShapeCamera(5).Visible = True
                ShapeCamera(5).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(5, "DEAD")
                'Call DataLogger("Lane6 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane6.exe")) Then
                Imgshutdown(5).Visible = False
                ShapeCamera(5).Visible = True
                ShapeCamera(5).FillColor = &HFF00&
                Call FrmTcpServer.LPR_Alive_State_Send(5, "LIVE")
            Else
                Imgshutdown(5).Visible = True
                ShapeCamera(5).Visible = True
                ShapeCamera(5).FillColor = &HFF&
                Call FrmTcpServer.LPR_Alive_State_Send(5, "DEAD")
                'Call DataLogger("Lane6 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(5).Visible = False
        ShapeCamera(5).Visible = False
    End If

End Sub


Public Sub ListView_Init1()
'    Dim Column_to_size As Integer
'
'    Call ListViewExtended(ListView1)
'    ListView1.View = lvwReport
'    ListView1.ListItems.Clear
'    ListView1.ColumnHeaders.Clear
'    ListView1.ColumnHeaders.Add , , " 차량번호    "
'    ListView1.ColumnHeaders.Add , , " 이    름    "
'    ListView1.ColumnHeaders.Add , , " 구    분        "
'    ListView1.ColumnHeaders.Add , , " 연 락 처        "

'    'ListView1.ColumnHeaders.Add , , " 차량모델   "
'    If (Glo_User_Type = "구분1/구분2") Then
'        ListView1.ColumnHeaders.Add , , " 소속, 직급   "
'    Else
'        ListView1.ColumnHeaders.Add , , " 동, 호수   "
'    End If
'
'    ListView1.ColumnHeaders.Add , , " 시 작 일  "
'    ListView1.ColumnHeaders.Add , , " 만 료 일  "
'    ListView1.ColumnHeaders.Add , , " 등록일시  "
'    ListView1.ColumnHeaders.Add , , "  "
'
'    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 2
'         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
'    Next
End Sub

