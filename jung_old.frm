VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Jung 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "  HOST Program"
   ClientHeight    =   14670
   ClientLeft      =   -28740
   ClientTop       =   1275
   ClientWidth     =   19185
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  '단색
   Icon            =   "jung.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "jung.frx":A4D2
   ScaleHeight     =   14670
   ScaleWidth      =   19185
   Begin LPR_PARKING_HOST.Server Server1 
      Left            =   6585
      Top             =   60
      _extentx        =   741
      _extenty        =   741
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4860
      Top             =   60
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   13830
      Width           =   18975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Refresh"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   19320
      TabIndex        =   14
      Top             =   13860
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   26.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   10  '한글 
      Left            =   8595
      TabIndex        =   6
      Top             =   2190
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11520
      Style           =   1  '그래픽
      TabIndex        =   12
      Top             =   2205
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   300
      Left            =   6405
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   11640
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   19410
      Top             =   330
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ParkHost"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DataJung 
      Height          =   375
      Left            =   19410
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ParkHost"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   0
      Left            =   22560
      TabIndex        =   82
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "영업현황"
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
      Picture         =   "jung.frx":58E2A
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   1
      Left            =   25260
      TabIndex        =   84
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "일반입차"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "jung.frx":5917B
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   2
      Left            =   26595
      TabIndex        =   85
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "일반출차"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "jung.frx":594CC
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   3
      Left            =   10575
      TabIndex        =   2
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
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
      RoundedCorners  =   0   'False
      Picture         =   "jung.frx":5981D
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   4
      Left            =   11895
      TabIndex        =   0
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "보호모드"
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
      Picture         =   "jung.frx":59B6E
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   5
      Left            =   22590
      TabIndex        =   80
      Top             =   1875
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "데이터 관리"
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
      Picture         =   "jung.frx":59EBF
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   6
      Left            =   22635
      TabIndex        =   4
      Top             =   3165
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "일반권 입차"
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
      Picture         =   "jung.frx":5A210
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   7
      Left            =   23925
      TabIndex        =   83
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "등록현황"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "jung.frx":5A561
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   8
      Left            =   17145
      TabIndex        =   5
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
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
      Picture         =   "jung.frx":5A8B2
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1500
      Left            =   6390
      TabIndex        =   13
      Top             =   3270
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   2646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   1485
      Left            =   6435
      TabIndex        =   16
      Top             =   12000
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   2619
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5385
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   9
      Left            =   13215
      TabIndex        =   1
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
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
      Picture         =   "jung.frx":5AC03
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   10
      Left            =   15840
      TabIndex        =   3
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
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
      RoundedCorners  =   0   'False
      Picture         =   "jung.frx":5AF54
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   11
      Left            =   14520
      TabIndex        =   78
      Top             =   870
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "TCP Server"
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
      Picture         =   "jung.frx":5B2A5
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   810
      Index           =   12
      Left            =   21225
      TabIndex        =   81
      Top             =   1530
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1429
      _StockProps     =   78
      Caption         =   "정기권 결산"
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
      Picture         =   "jung.frx":5B5F6
   End
   Begin MSWinsockLib.Winsock Host_sock 
      Left            =   5850
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.Label LblRecStat 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   630
      Index           =   1
      Left            =   16800
      TabIndex        =   75
      Top             =   2310
      Width           =   2055
   End
   Begin VB.Label LblRecStat 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   630
      Index           =   0
      Left            =   3810
      TabIndex        =   76
      Top             =   2310
      Width           =   2055
   End
   Begin VB.Label lbl_GN 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      Caption         =   "입구"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   21.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Index           =   1
      Left            =   13545
      TabIndex        =   87
      Top             =   2310
      Width           =   5175
   End
   Begin VB.Label lbl_GN 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      Caption         =   "입구"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   21.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Index           =   0
      Left            =   555
      TabIndex        =   86
      Top             =   2310
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "입출내역"
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
      Left            =   6630
      TabIndex        =   79
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   2220
      Index           =   2
      Left            =   6435
      Picture         =   "jung.frx":5B947
      Stretch         =   -1  'True
      Top             =   8115
      Width           =   2955
   End
   Begin VB.Label LblTime 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   14190
      TabIndex        =   77
      Top             =   150
      Width           =   4800
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   4320
      Index           =   0
      Left            =   240
      Picture         =   "jung.frx":8107A
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5730
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   4290
      Index           =   1
      Left            =   13230
      Picture         =   "jung.frx":A67AD
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5730
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   74
      Top             =   6645
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   73
      Top             =   6315
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   72
      Top             =   5970
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   71
      Top             =   5640
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   70
      Top             =   4965
      Width           =   60
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "기        간"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6765
      TabIndex        =   69
      Top             =   6660
      Width           =   840
   End
   Begin VB.Label Label5 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "차 량 모 델"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6750
      TabIndex        =   68
      Top             =   6330
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "연  락  처 "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6765
      TabIndex        =   67
      Top             =   5985
      Width           =   840
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "구        분"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6765
      TabIndex        =   66
      Top             =   5640
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "이        름"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6765
      TabIndex        =   65
      Top             =   5310
      Width           =   840
   End
   Begin VB.Label Lbl1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "차 량 번 호"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   6765
      TabIndex        =   64
      Top             =   4965
      Width           =   900
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   63
      Top             =   5295
      Width           =   60
   End
   Begin VB.Label LblSearch 
      BackColor       =   &H00000000&
      Caption         =   "검색결과 : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   6375
      TabIndex        =   62
      Top             =   2910
      Width           =   6435
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2370
      TabIndex        =   61
      Top             =   10320
      Width           =   3585
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2340
      TabIndex        =   60
      Top             =   10815
      Width           =   3615
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2370
      TabIndex        =   59
      Top             =   11325
      Width           =   3585
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2340
      TabIndex        =   58
      Top             =   11790
      Width           =   3615
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2340
      TabIndex        =   57
      Top             =   12285
      Width           =   3615
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2340
      TabIndex        =   56
      Top             =   12780
      Width           =   3615
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_in"
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
      Left            =   2340
      TabIndex        =   55
      Top             =   13260
      Width           =   3615
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   375
      TabIndex        =   54
      Top             =   10320
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   360
      TabIndex        =   53
      Top             =   10815
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   360
      TabIndex        =   52
      Top             =   11295
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   360
      TabIndex        =   51
      Top             =   11775
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   360
      TabIndex        =   50
      Top             =   12285
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   360
      TabIndex        =   49
      Top             =   12780
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_in"
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
      Left            =   360
      TabIndex        =   48
      Top             =   13260
      Width           =   1800
   End
   Begin VB.Label lbl_carno 
      BackStyle       =   0  '투명
      Caption         =   "경기00가0000"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Index           =   0
      Left            =   2550
      TabIndex        =   47
      Top             =   8400
      Width           =   3405
   End
   Begin VB.Label lbl_time_now 
      BackStyle       =   0  '투명
      Caption         =   "lbl_time_now"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Index           =   0
      Left            =   2550
      TabIndex        =   46
      Top             =   9120
      Width           =   3405
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Left            =   13350
      TabIndex        =   45
      Top             =   10320
      Width           =   1815
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Index           =   1
      Left            =   13335
      TabIndex        =   44
      Top             =   10800
      Width           =   1830
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Index           =   2
      Left            =   13350
      TabIndex        =   43
      Top             =   11280
      Width           =   1815
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Left            =   13350
      TabIndex        =   42
      Top             =   11775
      Width           =   1815
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Left            =   13350
      TabIndex        =   41
      Top             =   12285
      Width           =   1815
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Index           =   5
      Left            =   13335
      TabIndex        =   40
      Top             =   12765
      Width           =   1830
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_title_out"
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
      Index           =   6
      Left            =   13350
      TabIndex        =   39
      Top             =   13260
      Width           =   1815
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   15330
      TabIndex        =   38
      Top             =   10350
      Width           =   3615
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
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
      Left            =   15315
      TabIndex        =   37
      Top             =   10815
      Width           =   3630
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
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
      Left            =   15315
      TabIndex        =   36
      Top             =   11280
      Width           =   3630
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
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
      Left            =   15315
      TabIndex        =   35
      Top             =   11775
      Width           =   3630
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
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
      Left            =   15315
      TabIndex        =   34
      Top             =   12300
      Width           =   3630
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
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
      Index           =   5
      Left            =   15315
      TabIndex        =   33
      Top             =   12765
      Width           =   3630
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '투명
      Caption         =   "lbl_info_out"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   15315
      TabIndex        =   32
      Top             =   13260
      Width           =   3630
   End
   Begin VB.Label lbl_carno 
      BackStyle       =   0  '투명
      Caption         =   "경기00가0000"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   1
      Left            =   15510
      TabIndex        =   31
      Top             =   8400
      Width           =   3405
   End
   Begin VB.Label lbl_time_now 
      BackStyle       =   0  '투명
      Caption         =   "lbl_time_now"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   1
      Left            =   15510
      TabIndex        =   30
      Top             =   9120
      Width           =   3405
   End
   Begin VB.Label Proc_Type 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00404040&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   27.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   0
      Left            =   465
      TabIndex        =   29
      Top             =   7500
      Width           =   5355
   End
   Begin VB.Label Proc_Type 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00404040&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   27.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   1
      Left            =   13440
      TabIndex        =   28
      Top             =   7500
      Width           =   5355
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "수 정 일 시"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   1
      Left            =   6765
      TabIndex        =   27
      Top             =   7005
      Width           =   900
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   8655
      TabIndex        =   26
      Top             =   6990
      Width           =   60
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   9465
      TabIndex        =   25
      Top             =   8130
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   9465
      TabIndex        =   24
      Top             =   8535
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   9465
      TabIndex        =   23
      Top             =   8940
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   9465
      TabIndex        =   22
      Top             =   9345
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   9465
      TabIndex        =   21
      Top             =   9750
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   9465
      TabIndex        =   20
      Top             =   10155
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   9465
      TabIndex        =   19
      Top             =   10560
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   9465
      TabIndex        =   18
      Top             =   10965
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   9465
      TabIndex        =   17
      Top             =   11370
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   270
      Picture         =   "jung.frx":CBEE0
      Top             =   660
      Width           =   18600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 종 료 일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   12
      Left            =   31905
      TabIndex        =   10
      Top             =   5940
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 시 작 일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   11
      Left            =   31890
      TabIndex        =   9
      Top             =   5295
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 발 급 일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   10
      Left            =   31905
      TabIndex        =   8
      Top             =   4650
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   " 월정 요금"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   9
      Left            =   31905
      TabIndex        =   7
      Top             =   4005
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      Height          =   615
      Index           =   11
      Left            =   31830
      Top             =   5715
      Width           =   3645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      Height          =   615
      Index           =   10
      Left            =   31830
      Top             =   5070
      Width           =   3645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      Height          =   615
      Index           =   9
      Left            =   31830
      Top             =   4425
      Width           =   3645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      Height          =   615
      Index           =   8
      Left            =   31830
      Top             =   3780
      Width           =   3645
   End
End
Attribute VB_Name = "Jung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyText(1 To 8) As New clsText
Dim DataField_Enabled As Boolean
Dim Save_TagNum As String

Private Sub cmd_menu_Click(Index As Integer)
Dim i As Integer

Me.MousePointer = 11
 Select Case Index
'            Case 0
'                 Rptmenu.Show 1
'                 Me.MousePointer = 0
'                 Call DataLogger("[HOST Button]    " & "영업현황 보고서 화면 접근")
'                 'Call Err_doc("호스트 : 영업현황 보고서 화면 접근")
'            Case 1
'                 IlINList.Show 1
'                 Me.MousePointer = 0
'                 Call DataLogger("[HOST Button]    " & "일반권 입차 보고서 화면 접근")
'                 'Call Err_doc("호스트 : 일반권 입차 보고서 화면 접근")
'            Case 2
'                 IlIOList.Show 1
'                 Me.MousePointer = 0
'                 Call Err_doc("호스트 : 일반권 출차 보고서 화면 접근")
'            Case 3
'                 'JIOSch.Show 1
'                 FrmInOut.Show 1
'                 Me.MousePointer = 0
'                 Call DataLogger("[HOST Button]    " & "입출차 보고서 화면 접근")
'                 'Call Err_doc("호스트 : 입출차 보고서 화면 접근")
'            Case 4
'                 If (cmd_menu(4).Caption = "보호모드") Then
'                    Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
'                    'Call Err_doc("호스트 : 프로그램 보호모드로 전환")
'                    cmd_menu(4).Caption = "보호해제"
'                    For i = 0 To 12
'                        If (i <> 4) Then
'                           cmd_menu(i).Enabled = False
'                        End If
'                    Next i
'                    cmd_menu(8).Enabled = True
'                    cmd_menu(3).Enabled = True
'                    Put_Ini "System Config", "보호모드", "True"
'                 Else
'                    Protect.Show 1
'                    Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
'                    'Call Err_doc("호스트 : 프로그램 보호모드 해제")
'                 End If
'                 Me.MousePointer = 0
'            Case 5
'                 frmDbase.Show 1
'                 Me.MousePointer = 0
'                 Call Err_doc("호스트 : 데이터베이스 관리 화면 접근")
'            Case 6
'                 FrmTicketIn.Show 1
'                 Me.MousePointer = 0
'                 Call Err_doc("호스트 : 일반권 입차현황 화면 접근")
'            Case 7
'                 JungList2.Show 1
'                 Me.MousePointer = 0
'                 Call Err_doc("호스트 : 등록현황 화면 접근")
'            Case 8
'                 Call Server1.StopServer
'                 Unload Me
'            Case 9
'                 'Jung_New.Show 1
'                 FrmReg.Show 1
'                 Me.MousePointer = 0
'                 Call DataLogger("[HOST Button]    " & "정기권관리 화면 접근")
'                 'Call Err_doc("호스트 : 정기권관리 화면 접근")
'            Case 10
'                 'frmTicketSettle.Show 1
'                 FrmCSV.Show 1
'                 Me.MousePointer = 0
'                 Call DataLogger("[HOST Button]    " & "일괄등록 화면 접근")
'                 'Call Err_doc("호스트 : 일괄등록 화면 접근")
'            Case 11
'                 'FrmCouponSale.Show 1
'                 FrmTcpServer.Show 0
'                 Me.MousePointer = 0
'                 Call DataLogger("[HOST Button]    " & "TCP Server 화면 접근")
'                 'Call Err_doc("호스트 : TCP Server 화면 접근")
'            Case 12
'                 FrmFee.Show 1
'                 Me.MousePointer = 0
'                 Call Err_doc("호스트 : 정기권 결산 화면 접근")
'
       End Select
End Sub

Private Sub cmd_menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

If (Glo_cmd_menu_index = Index) Then
    If (Glo_cmd_menu_index <> 99) Then
        Exit Sub
    End If
End If
Glo_cmd_menu_index = Index

For i = 0 To 12
    cmd_menu(i).ForeColor = vbWhite
Next i
'cmd_menu(0).ForeColor = vbWhite

cmd_menu(Index).ForeColor = vbGreen

End Sub

Private Sub Command1_Click()
    LblCar(0).Caption = ""
    LblName(0).Caption = ""
    LblId(0).Caption = ""
    LblCarType(0).Caption = ""
    LblTel(0).Caption = ""
    LblDate(0).Caption = ""
    LblGubun(0).Caption = ""
    LblSearch = ""
    ListView1.ListItems.Clear
    Text1 = ""
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
ListView2.ListItems.Clear
'성훈
Lbl_inout(0).Caption = " 출입일시 : "
Lbl_inout(1).Caption = " 차량번호 : "
Lbl_inout(2).Caption = " 이    름 : "
Lbl_inout(3).Caption = " 구    분 : "
Lbl_inout(4).Caption = " 연 락 처 : "
Lbl_inout(5).Caption = " 인식번호 : "
Lbl_inout(6).Caption = " 종 료 일 : "
Lbl_inout(7).Caption = " 입출상태 : "
Lbl_inout(8).Caption = " 입출구분 : "

End Sub



Private Sub Form_Load()
Dim i As Integer
Dim SQL As String
Dim Reg_Addr As String

'Me.Caption = Me.Caption & "  " & "Version " & App.Major & "." & App.Minor & "." & App.Revision

Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
'Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
'Left = 0
Top = 0

Call ListView_Init1
Call ListView_Init2

'Call Server1.StartServer(Server_Port, Server1.ServerIP)
FrmTcpServer.Show 0
If (HostType = 0) Then
    lbl_GN(0).Caption = Trim(LANE1_Name)
    lbl_GN(1).Caption = Trim(LANE2_Name)
Else
    lbl_GN(0).Caption = Trim(LANE1_Name)
    lbl_GN(1).Caption = Trim(LANE3_Name)
End If

For i = 0 To 6
    lbl_title_in(i).Caption = ""
    lbl_info_in(i).Caption = ""
    lbl_title_out(i).Caption = ""
    lbl_info_out(i).Caption = ""
Next i
lbl_carno(0).Caption = ""
lbl_time_now(0).Caption = ""
lbl_carno(1).Caption = ""
lbl_time_now(1).Caption = ""

For i = 0 To 8
    Lbl_inout(i).BackStyle = 0
Next i
Lbl_inout(0).Caption = " 출입일시 : "
Lbl_inout(1).Caption = " 차량번호 : "
Lbl_inout(2).Caption = " 이    름 : "
Lbl_inout(3).Caption = " 구    분 : "
Lbl_inout(4).Caption = " 연 락 처 : "
Lbl_inout(5).Caption = " 인식번호 : "
Lbl_inout(6).Caption = " 종 료 일 : "
Lbl_inout(7).Caption = " 입출상태 : "
Lbl_inout(8).Caption = " 입출구분 : "

Call cmd_menu_Click(4)

Timer1.Enabled = True

FrmTcpServer.Hide

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

For i = 0 To 12
    cmd_menu(i).Visible = False
Next i
Glo_cmd_menu_index = 99
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

For i = 0 To 12
    cmd_menu(i).Visible = True
Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msg, Style, Title, Response
Dim Ret As Boolean
msg = "프로그램을 종료하시겠습니까?         "
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Parking Manager™  - JWT   "
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then
    Call DataLogger("[HOST Button]    " & "프로그램 종료")
    'Call Err_doc("호스트 : " & "프로그램 정상적으로 종료")
    Call DataBaseClose(adoConn)
    Call Unhook
    End
End If
Me.MousePointer = 0
Cancel = True
End Sub

Private Sub Server1_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
'Dim Qry As String
'Dim rs As ADODB.Recordset
'Dim Tmp_Path As String
'Dim itmX As ListItem
'Dim CarNum As String
'Dim GateNo As Integer
'Dim inout As String
'Dim Gubun As String
'Dim i, s As Integer
'Dim ECHO As ICMP_ECHO_REPLY
'
'On Error GoTo err_P
'
''Debug.Print Data
''Data = "0_인식실패_\\192.168.0.20\\Image\20130320\20130320194926671_인식실패.jpg"
''0_25구5401_\\192.168.0.200\\Image\20130320\20130320195911296_25구5401.jpg
'
'List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Data, 0
'
'If Len(Data) > 100 Then
'    Exit Sub
'End If
'
''기타 컨맨드 처리 후
''CMD
'
'
'
'
''LPR Protocol Proc
'i = InStr(1, Data, "_", 1)
'GateNo = Val(Left(Data, (i - 1)))
'Glo_GateNo = GateNo
'Select Case Glo_GateNo
'    Case 0
'        Glo_Lpr_IP = LANE1_IP
'    Case 1
'        Glo_Lpr_IP = LANE2_IP
'    Case 2
'        Glo_Lpr_IP = LANE3_IP
'    Case 3
'        Glo_Lpr_IP = LANE4_IP
'End Select
'If Glo_GateNo Mod 2 = 0 Then
'    Glo_GateGubun = 0
'Else
'    Glo_GateGubun = 1
'End If
'
's = InStr(4, Data, "_", 1)
'CarNum = Mid(Data, (i + 1), (s - i - 1))
'Glo_CarNum = CarNum
'i = Len(Data)
'Tmp_Path = Mid(Data, (s + 1), i)
'
''Debug.Print GateNo
''Debug.Print CarNum
''Debug.Print Tmp_Path
'Call LPRIn_Proc(CarNum, Tmp_Path)
'Call Jung_Show(CarNum)
'
'Server1.SendData Format(Now, "YYYYMMDDHHNNSS"), SckIndex
'
'Exit Sub
'
'err_P:
'        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & CarNum & "  " & Tmp_Path)
'        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Description)
End Sub

Private Sub Server1_Error(ByVal SckIndex As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
'Debug.Print Description

End Sub

Private Sub Timer1_Timer()
Dim Qry As String
Dim rs As ADODB.Recordset

LblTime(0).Caption = "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")

'If (Format(Now, "NNSS") = "0001") Then
'    '게이트 카운트 초기화
''    Qry = "show tables"
''    Set rs = New ADODB.Recordset
''    rs.Open Qry, adoConn
''    Set rs = Nothing
'    adoConn.Execute ""
'    List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "    MySQL Connection Test...!! ", 0
'Else
'    'List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss"), 0
'    'tmp = Format(Now, "HHNNSS")
'    'Debug.Print Format(Now, "NNSS")
'End If

End Sub

Private Sub ImageIn_DblClick(Index As Integer)
If (Index = 2) Then
    If (ImageIn(2).Height = 3780) Then
        ImageIn(2).Height = 2220
        ImageIn(2).Width = 2955
    Else
        ImageIn(2).Height = 3780
        ImageIn(2).Width = 6375
    End If
End If

End Sub


Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
ListView1.SetFocus

LblCar(0).Caption = ""
LblName(0).Caption = ""
LblId(0).Caption = ""
LblCarType(0).Caption = ""
LblTel(0).Caption = ""
LblDate(0).Caption = ""
LblGubun(0).Caption = ""

LblCar(0).Caption = ListView1.SelectedItem.Text
LblName(0).Caption = ListView1.SelectedItem.SubItems(1)
LblId(0).Caption = ListView1.SelectedItem.SubItems(2)
LblCarType(0).Caption = ListView1.SelectedItem.SubItems(3)
LblTel(0).Caption = ListView1.SelectedItem.SubItems(4)

If (ListView1.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView1.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
    LblDate(0).ForeColor = &H404040
    LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6)
Else
    LblDate(0).ForeColor = vbRed
    LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6) & "   " & "[기간에러]"
End If

LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)

End Sub


Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
ListView2.SetFocus

Lbl_inout(0).Caption = " 출입일시 : "
Lbl_inout(1).Caption = " 차량번호 : "
Lbl_inout(2).Caption = " 이    름 : "
Lbl_inout(3).Caption = " 구    분 : "
Lbl_inout(4).Caption = " 연 락 처 : "
Lbl_inout(5).Caption = " 인식번호 : "
Lbl_inout(6).Caption = " 종 료 일 : "
Lbl_inout(7).Caption = " 입출상태 : "
Lbl_inout(8).Caption = " 입출구분 : "

Lbl_inout(0).Caption = " 출입일시 : " & Format(ListView2.SelectedItem.SubItems(7), "####-##-## ##:##:##")
Lbl_inout(1).Caption = " 차량번호 : " & ListView2.SelectedItem.Text
Lbl_inout(2).Caption = " 이    름 : " & ListView2.SelectedItem.SubItems(2)
Lbl_inout(3).Caption = " 구    분 : " & ListView2.SelectedItem.SubItems(1)
Lbl_inout(4).Caption = " 연 락 처 : " & ListView2.SelectedItem.SubItems(3)
Lbl_inout(5).Caption = " 인식번호 : " & ListView2.SelectedItem.SubItems(4)
Lbl_inout(6).Caption = " 종 료 일 : " & ListView2.SelectedItem.SubItems(5)

If ((ListView2.SelectedItem.SubItems(6) = "정상입차") Or (ListView2.SelectedItem.SubItems(6) = "정상출차")) Then
    Lbl_inout(7).ForeColor = vbWhite
Else
    Lbl_inout(7).ForeColor = vbRed
End If
Lbl_inout(7).Caption = " 입출상태 : " & ListView2.SelectedItem.SubItems(6)

'성훈 추가
Lbl_inout(8).Caption = " 입출구분 : " & ListView2.SelectedItem.SubItems(8)

ImageIn(2).Picture = LoadPicture(ListView2.SelectedItem.SubItems(9))
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Car_Num_Str As String
Dim Qry As String
Dim rs As Recordset
Dim rs_Part As Recordset
Dim itmX As ListItem

'On Error GoTo erro_p

If (KeyAscii = 13) Then
        LblCar(0).Caption = ""
        LblName(0).Caption = ""
        LblId(0).Caption = ""
        LblCarType(0).Caption = ""
        LblTel(0).Caption = ""
        LblDate(0).Caption = ""
        LblGubun(0).Caption = ""
        If ((Len(Text1) <> 4) Or Not (IsNumeric(Text1))) Then
            MsgBox "차량번호 숫자 네지리를 정확하게 입력하세요!"
            Text1 = ""
            Exit Sub
        End If
        Qry = "Select * From tb_reg WHERE CAR_NO Like '%" & Text1 & "' ORDER BY CAR_NO"
        Set rs = New ADODB.Recordset
        rs.Open Qry, adoConn
        
        ListView1.ListItems.Clear
        
        If (rs.EOF) Then
            LblSearch.Caption = "검색결과 : 자료가 존재 하지않습니다.."
        Else
            LblSearch.Caption = "검색결과 : " & (rs.RecordCount) & " 건"
            
            Do While Not (rs.EOF)
                Set itmX = ListView1.ListItems.Add(, , "" & rs!CAR_NO)
                itmX.SubItems(1) = "" & rs!DRIVER_NAME
                itmX.SubItems(2) = "" & rs!CAR_GUBUN
                itmX.SubItems(3) = "" & rs!DRIVER_PHONE
                itmX.SubItems(4) = "" & rs!CAR_MODEL
                itmX.SubItems(5) = "" & rs!Start_Date
                itmX.SubItems(6) = "" & rs!End_Date
                itmX.SubItems(7) = "" & rs!REG_DATE
                rs.MoveNext
            Loop
            
            ListView1.ListItems.Item(1).Selected = True
            
            If (rs.RecordCount = 1) Then
            
            Else
                ListView1.SetFocus
            End If
            
            LblCar(0).Caption = ListView1.SelectedItem.Text
            LblName(0).Caption = ListView1.SelectedItem.SubItems(1)
            LblId(0).Caption = ListView1.SelectedItem.SubItems(2)
            LblCarType(0).Caption = ListView1.SelectedItem.SubItems(3)
            LblTel(0).Caption = ListView1.SelectedItem.SubItems(4)
            If (ListView1.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView1.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
                LblDate(0).ForeColor = &H404040
                LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6)
            Else
                LblDate(0).ForeColor = vbRed
                LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6) & "   " & "[기간만료]"
            End If
            
            '성훈
            LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)
        
        End If
        
        Set rs = Nothing
        KeyAscii = 0
        Exit Sub
End If

'erro_p:
'    MsgBox Err.Description
End Sub

Public Sub ListView_Init1()
Dim Column_to_size As Integer

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " 차량번호    "
    ListView1.ColumnHeaders.Add , , " 이    름    "
    ListView1.ColumnHeaders.Add , , " 구    분        "
    ListView1.ColumnHeaders.Add , , " 연 락 처        "
    ListView1.ColumnHeaders.Add , , " 차량모델   "
    ListView1.ColumnHeaders.Add , , " 시 작 일  "
    ListView1.ColumnHeaders.Add , , " 만 료 일  "
    ListView1.ColumnHeaders.Add , , " 수정일시  "
    ListView1.ColumnHeaders.Add , , "  "
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 2
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End Sub

Public Sub ListView_Init2()
Dim Column_to_size As Integer

    Call ListViewExtended(ListView2)
    ListView2.View = lvwReport
    ListView2.ListItems.Clear
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , " 차량번호     "      '0
    ListView2.ColumnHeaders.Add , , " 구    분         "  '1
    ListView2.ColumnHeaders.Add , , " 이    름  "       '2
    ListView2.ColumnHeaders.Add , , " 전화번호     "  '3
    ListView2.ColumnHeaders.Add , , " 인식번호     "   '4
    ListView2.ColumnHeaders.Add , , " 종 료 일     "        '5
    ListView2.ColumnHeaders.Add , , " 인식상태     "          '6
    ListView2.ColumnHeaders.Add , , " 처리일시     "         '7
    ListView2.ColumnHeaders.Add , , " 입출구분     "    '8
    ListView2.ColumnHeaders.Add , , " 이미지명                                            "    '9
    
    ListView2.ColumnHeaders.Add , , " "
    'ListView2.SortKey = 11
    ListView2.SortOrder = lvwDescending
    ListView2.Sorted = True
    
    For Column_to_size = 0 To ListView2.ColumnHeaders.Count - 2
         SendMessage ListView2.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Winsock1.GetData strData, , bytesTotal
Debug.Print strData
Host_sock.SendData (strData)
End Sub



