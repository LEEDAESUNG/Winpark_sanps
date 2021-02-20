VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmAccnt 
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   14865
   ClientLeft      =   5910
   ClientTop       =   7020
   ClientWidth     =   19395
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAccnt.frx":0000
   ScaleHeight     =   14865
   ScaleWidth      =   19395
   StartUpPosition =   1  '소유자 가운데
   Begin VB.ListBox List_SALE 
      Height          =   1860
      Left            =   180
      TabIndex        =   35
      Top             =   13020
      Width           =   11805
   End
   Begin VB.ListBox List_OP 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   180
      TabIndex        =   34
      Top             =   7560
      Width           =   11805
   End
   Begin VB.Timer Timer_APS_Monitor 
      Interval        =   300
      Left            =   2310
      Top             =   1110
   End
   Begin VB.TextBox TxtVal 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   14820
      MaxLength       =   6
      TabIndex        =   14
      Top             =   13275
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox cmb_Aps 
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
      ItemData        =   "FrmAccnt.frx":3ACAC2
      Left            =   12240
      List            =   "FrmAccnt.frx":3ACAC4
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   13275
      Width           =   2445
   End
   Begin VB.TextBox txt_100 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   14805
      MaxLength       =   4
      TabIndex        =   1
      Top             =   6585
      Width           =   1035
   End
   Begin VB.TextBox txt_500 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   14805
      MaxLength       =   4
      TabIndex        =   0
      Top             =   6090
      Width           =   1020
   End
   Begin VB.TextBox txt_1000 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   14685
      MaxLength       =   4
      TabIndex        =   3
      Top             =   8925
      Width           =   1110
   End
   Begin VB.TextBox txt_5000 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   14685
      MaxLength       =   4
      TabIndex        =   2
      Top             =   8280
      Width           =   1110
   End
   Begin MSWinsockLib.Winsock CmdS_Sock 
      Left            =   2940
      Top             =   1110
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ListView ListView_OP 
      Height          =   675
      Left            =   19560
      TabIndex        =   30
      Top             =   2700
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1191
      View            =   3
      Arrange         =   2
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView ListView_SL 
      Height          =   1875
      Left            =   19560
      TabIndex        =   31
      Top             =   3510
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   3307
      View            =   3
      Arrange         =   2
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   14820
      TabIndex        =   32
      Top             =   13770
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   208011264
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   14820
      TabIndex        =   33
      Top             =   14220
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   208011266
      CurrentDate     =   36927
   End
   Begin MSWinsockLib.Winsock CmdR_Sock 
      Left            =   3480
      Top             =   1110
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "IP CAMERA"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   4
      Left            =   12240
      TabIndex        =   36
      Top             =   1440
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "APS 명령"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   14010
      TabIndex        =   29
      Top             =   1530
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "결제내역"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   15810
      TabIndex        =   28
      Top             =   1530
      Width           =   1050
   End
   Begin VB.Image Img_Run 
      Height          =   720
      Left            =   17595
      Top             =   13140
      Width           =   1260
   End
   Begin VB.Label LblRun 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "실  행"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17700
      TabIndex        =   27
      Top             =   13380
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   4365
      Left            =   6315
      Picture         =   "FrmAccnt.frx":3ACAC6
      Stretch         =   -1  'True
      Top             =   2415
      Width           =   5670
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "서울12가1234"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   9
      Left            =   1125
      TabIndex        =   26
      Top             =   5595
      Width           =   2595
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "대기중..."
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Index           =   8
      Left            =   1125
      TabIndex        =   25
      Top             =   4800
      Width           =   2490
   End
   Begin VB.Label LblMsg 
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "LblMsg"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   1455
      TabIndex        =   24
      Top             =   6375
      Width           =   3600
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "경기12가1234"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   7
      Left            =   1845
      TabIndex        =   23
      Top             =   3180
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "잔   액 :"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   6
      Left            =   1830
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "지   불 :"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   5
      Left            =   1830
      TabIndex        =   21
      Top             =   4905
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "요   금 :"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   4
      Left            =   1830
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "할   인 :"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   3
      Left            =   1830
      TabIndex        =   19
      Top             =   4470
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "주차시간 : 12시간 30분"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   1830
      TabIndex        =   18
      Top             =   4215
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "입차일시 : 2016-10-30 12:30:45"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1830
      TabIndex        =   17
      Top             =   3990
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label LblAps 
      BackStyle       =   0  '투명
      Caption         =   "입차일시 : 2016-10-30 12:30:45"
      BeginProperty Font 
         Name            =   "휴먼모음T"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   1830
      TabIndex        =   16
      Top             =   3765
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   4365
      Left            =   180
      Picture         =   "FrmAccnt.frx":3B9E93
      Stretch         =   -1  'True
      Top             =   2415
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "분"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   16560
      TabIndex        =   15
      Top             =   13335
      Width           =   510
   End
   Begin VB.Label LblBillHopper 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "적  용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17655
      TabIndex        =   13
      ToolTipText     =   "지폐방출기내의 시제금을 재설정 합니다."
      Top             =   8595
      Width           =   1050
   End
   Begin VB.Label LblCoin 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "적  용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17670
      TabIndex        =   12
      ToolTipText     =   "동전수납통의 시제금을 재설정 합니다."
      Top             =   6330
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "종  료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   17625
      TabIndex        =   11
      Top             =   1530
      Width           =   1050
   End
   Begin VB.Label LblBill 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "출  금"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   17670
      TabIndex        =   10
      ToolTipText     =   "지폐인식기내의 금액을 모두 출금합니다"
      Top             =   3870
      Width           =   1050
   End
   Begin VB.Label LblBill1000 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   14355
      TabIndex        =   9
      Top             =   4560
      Width           =   1485
   End
   Begin VB.Label LblBill5000 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   14355
      TabIndex        =   8
      Top             =   4035
      Width           =   1485
   End
   Begin VB.Label LblBill10000 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   14355
      TabIndex        =   7
      Top             =   3525
      Width           =   1485
   End
   Begin VB.Label lbl_TotalAccnt 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "99999999 "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   450
      Left            =   14730
      TabIndex        =   6
      Top             =   9975
      Width           =   2265
   End
   Begin VB.Label lbl_Update 
      BackStyle       =   0  '투명
      Caption         =   "Update Date : "
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
      Height          =   240
      Left            =   14730
      TabIndex        =   5
      Top             =   2715
      Width           =   4350
   End
   Begin VB.Image Img_IN 
      Height          =   600
      Left            =   17580
      ToolTipText     =   "지폐방출기내의 시제금을 재설정 합니다."
      Top             =   8445
      Width           =   1185
   End
   Begin VB.Image Img_CoinOut 
      Height          =   600
      Left            =   17610
      ToolTipText     =   "동전수납통의 시제금을 재설정 합니다."
      Top             =   6180
      Width           =   1155
   End
   Begin VB.Image Img_BillOut 
      Height          =   600
      Left            =   17595
      ToolTipText     =   "지폐인식기내의 금액을 모두 출금합니다"
      Top             =   3765
      Width           =   1170
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   1
      Left            =   15480
      Picture         =   "FrmAccnt.frx":40E9FD
      Top             =   1110
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   3
      Left            =   13680
      Picture         =   "FrmAccnt.frx":413D2B
      Top             =   1110
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   0
      Left            =   17280
      Picture         =   "FrmAccnt.frx":419059
      Top             =   1110
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   2
      Left            =   17370
      Picture         =   "FrmAccnt.frx":41E387
      Top             =   13050
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Enabled         =   0   'False
      Height          =   915
      Index           =   4
      Left            =   11895
      Picture         =   "FrmAccnt.frx":4236B5
      Top             =   1110
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "FrmAccnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyText(1 To 5) As New clsText





Private Sub Form_Load()
    Dim i As Integer
    
    
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    Set MyText(1).MyText = Me.txt_5000
    Set MyText(2).MyText = Me.txt_1000
    Set MyText(3).MyText = Me.txt_500
    Set MyText(4).MyText = Me.txt_100
    Set MyText(5).MyText = Me.TxtVal
    Call Read_Account
    cmb_Aps.Clear
    cmb_Aps.AddItem "시간할인"
    cmb_Aps.AddItem "금액할인"
    cmb_Aps.AddItem "%   할인"
    cmb_Aps.AddItem "전액할인"
    cmb_Aps.AddItem "1일 주차요금"
    'cmb_Aps.AddItem "지정요금" '무인기 프로토콜 정의 후 사용권고
    cmb_Aps.AddItem "초기화면"
    cmb_Aps.AddItem "차단기 열림"
    cmb_Aps.AddItem "영수증"
    cmb_Aps.AddItem "지정시간"  '입차시간 재지정
    cmb_Aps.ListIndex = 0
    
    
    CmdS_Sock.Protocol = sckTCPProtocol
    CmdR_Sock.Protocol = sckTCPProtocol
    CmdR_Sock.RemotePort = Glo_APSCmdR_Port
    CmdR_Sock.Bind
    
    
    For i = 0 To 7
        LblAps(i).Visible = False
    Next i
    LblAps(9).Caption = ""
    
    
    Image1.Picture = LoadPicture(App.Path & "\Image\asp_small1.bmp")
    Image2.Picture = LoadPicture(App.Path & "\NoCar.jpg")
    LblMsg.Caption = ""
    
    Call ListView_OP_Init

End Sub

Public Sub ListView_OP_Init()
    Dim Column_to_size As Integer

    Call ListViewExtended(ListView_OP)
    ListView_OP.View = lvwReport
    ListView_OP.ListItems.Clear
    ListView_OP.ColumnHeaders.Clear
    ListView_OP.ColumnHeaders.Add , , " 처리일시         "         '7
    ListView_OP.ColumnHeaders.Add , , " 차량번호     "      '0
    ListView_OP.ColumnHeaders.Add , , " 구    분         "  '1
    ListView_OP.ColumnHeaders.Add , , " 결    과   "       '2
    ListView_OP.ColumnHeaders.Add , , " 이미지명                                            "    '9
    
    ListView_OP.ColumnHeaders.Add , , " "
    ListView_OP.SortOrder = lvwDescending
    ListView_OP.Sorted = True
    
    For Column_to_size = 0 To ListView_OP.ColumnHeaders.Count - 2
         SendMessage ListView_OP.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

End Sub

Private Sub cmb_Aps_Click()
'On Error Resume Next
    
    TxtVal.Visible = False
    Label1.Caption = ""
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    
    Select Case cmb_Aps.text
           Case "시간할인"
                TxtVal.Visible = True
                Label1.Caption = "분"
                'TxtVal.SetFocus
                DTPicker1.Visible = False
                DTPicker2.Visible = False
           
           Case "금액할인", "1일 주차요금", "지정요금"
                TxtVal.Visible = True
                Label1.Caption = "원"
                'TxtVal.SetFocus
                DTPicker1.Visible = False
                DTPicker2.Visible = False
           
           Case "%   할인"
                TxtVal.Visible = True
                Label1.Caption = "%"
                'TxtVal.SetFocus
                DTPicker1.Visible = False
                DTPicker2.Visible = False
                
           Case "지정시간"
                DTPicker1.Top = 13275
                DTPicker2.Top = 13770
                DTPicker1.Visible = True
                DTPicker2.Visible = True
                
                
'           Case Else
'                TxtVal.Visible = False
'                Label1.Caption = ""
'                DTPicker1.Visible = False
'                DTPicker2.Visible = False
    End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lblbutton(0).FontBold = False
    Lblbutton(1).FontBold = False
    Lblbutton(3).FontBold = False
    LblRun.FontBold = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set MyText(1).MyText = Nothing
    Set MyText(2).MyText = Nothing
    Set MyText(3).MyText = Nothing
    Set MyText(4).MyText = Nothing
    Set MyText(5).MyText = Nothing
    
    Timer_APS_Monitor.Enabled = False

End Sub

Private Sub Img_BillOut_Click()
    '지폐인식기에서 출금
    Dim tmp As Long
    Dim sQry As String
    Dim bQryResult As String
    
    tmp = Val(LblBill10000) * 10000 + Val(LblBill5000) * 5000 + Val(LblBill1000) * 1000
    If (tmp = 0) Then
        Call LISTBOX_PutString(List_OP, " 출금할 금액이 없습니다..!!")
    Else
        sQry = "UPDATE tb_account set BILL_S10000 = 0, BILL_S5000 = 0, BILL_S1000 = 0, UPDATE_DATE = sysdate(Now())"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        
        'MANAGER_LOG
        sQry = "INSERT INTO tb_manager_log(ID, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, INOUT_YN, REG_DATE) VALUES ('" & Glo_PartName & "', 'BILL STACK', 'Withdraw from Bill Stack', 'BILL_STACK', '" & tmp & "', 'OUT', SYSDATE(NOW()))"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        If (bQryResult = False) Then
            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If

        Read_Account
        Call LISTBOX_PutString(List_OP, "지폐인식기내 모든 금액을 출금하였습니다..!!")
    End If
End Sub



Private Sub Img_CoinOut_Click()
    Dim sQry As String
    Dim bQryResult As String
    
    If (IsNumeric(txt_500) And IsNumeric(txt_100)) Then
        '500원 동전수납통 설정
'''        sQry = "UPDATE tb_account set COIN_C500 = " & Val(txt_500) & ", UPDATE_DATE = sysdate(Now())"
'''        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
'''        If (bQryResult = False) Then
'''            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
'''            Exit Sub
'''        End If
'''
'''        sQry = "UPDATE tb_account set COIN_H500 = 0 , UPDATE_DATE = sysdate(Now())"
'''        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
'''        If (bQryResult = False) Then
'''            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
'''            Exit Sub
'''        End If
'''
'''
'''        '100원 동전수납통 설정
'''        sQry = "UPDATE tb_account set COIN_C100 = " & txt_100 & ", UPDATE_DATE = sysdate(Now())"
'''        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
'''        If (bQryResult = False) Then
'''            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
'''            Exit Sub
'''        End If
'''
'''        sQry = "UPDATE tb_account set COIN_H100 = 0 , UPDATE_DATE = sysdate(Now())"
'''        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
'''        If (bQryResult = False) Then
'''            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
'''            Exit Sub
'''        End If

        sQry = "UPDATE tb_account set COIN_C500 = " & Val(txt_500) & ", COIN_H500 = 0 , COIN_C100 = " & txt_100 & ", COIN_H100 = 0 , UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' "
        'Debug.Print sQry
        'sQry = "INSERT INTO tb_reg (CAR_NO,CAR_MODEL,CAR_GUBUN,CAR_FEE,DRIVER_NAME,DRIVER_PHONE,DRIVER_DEPT,DRIVER_CLASS,START_DATE,END_DATE,ETC,REG_DATE,UPDATE_DATE,FEE_DATE,DAY_ROTATION_YN,REG_PART,LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE) VALUES ('" & _
tmpCarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '" & sChkLane1 & "', '" & sChkLane2 & "', '" & sChkLane3 & "', '" & sChkLane4 & "', '" & sChkLane5 & "', '" & sChkLane6 & "', '" & sChkWeek1 & "', '" & sChkWeek2 & "', '" & sChkWeek3 & "', '" & sChkWeek4 & "', '" & sChkWeek5 & "', '" & sChkWeek6 & "', '" & sChkWeek7 & "', '" & sRotation & "', '" & sApp & "', '', Null, Null)"


        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        If (bQryResult = False) Then
            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        sQry = "INSERT INTO tb_manager_log(ID, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, INOUT_YN, REG_DATE) VALUES ('" & Glo_PartName & "', 'COIN CASE', 'Withdraw from Coin Case', 'COIN_C100', '" & txt_100 & "', 'OUT', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        sQry = "INSERT INTO tb_manager_log(ID, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, INOUT_YN, REG_DATE) VALUES ('" & Glo_PartName & "', 'COIN CASE', 'Withdraw from Coin Case', 'COIN_C500', '" & txt_500 & "', 'OUT', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        
        Read_Account
        'MsgBox "동전수납통의 수량이 설정 되었습니다..!!"
        Call LISTBOX_PutString(List_OP, " 동전수납통의 수량이 설정 되었습니다..!!")
        
    Else
        Call LISTBOX_PutString(List_OP, " 숫자만 입력해주세요")
        If (Not IsNumeric(txt_500)) Then
            txt_500.text = ""
            txt_500.SetFocus
        ElseIf (Not IsNumeric(txt_100)) Then
            txt_100.text = ""
            txt_100.SetFocus
        End If
    End If
End Sub


Private Sub Img_IN_Click()
    Dim sQry As String
    Dim bQryResult As String
    
    If (IsNumeric(txt_5000) And IsNumeric(txt_1000)) Then
        sQry = "UPDATE tb_account set BILL_H5000 = " & txt_5000 & ", BILL_H1000 = " & txt_1000 & ", UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "'"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        If (bQryResult = False) Then
            Call DataLogger("[FrmAccnt]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If

        'adoConn.Execute "INSERT INTO tb_manager_log(ID, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, INOUT_YN, REG_DATE) VALUES ('" & Glo_PartName & "', 'HOPPER', 'Deposit into Hopper', '" & txt_5000 & "_" & txt_1000 & "_" & txt_500 & "_" & txt_100 & "', '" & (txt_5000 * 5000) + (txt_1000 * 1000) + (txt_500 * 500) + (txt_100 * 100) & "', 'IN', SYSDATE(NOW()))"
        sQry = "INSERT INTO tb_manager_log(ID, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, INOUT_YN, REG_DATE) VALUES ('" & Glo_PartName & "', 'HOPPER', 'Deposit into Hopper', '" & txt_5000 & "_" & txt_1000 & "', '" & (txt_5000 * 5000) + (txt_1000 * 1000) & "', 'IN', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        
        Read_Account
        Call LISTBOX_PutString(List_OP, " 지폐방출기의 수량이 설정 되었습니다..!!")
    Else
        Call LISTBOX_PutString(List_OP, " 숫자만 입력해주세요")
        If (Not IsNumeric(txt_5000)) Then
            txt_5000.text = ""
            txt_5000.SetFocus
        ElseIf (Not IsNumeric(txt_1000)) Then
            txt_1000.text = ""
            txt_1000.SetFocus
        End If
    End If
End Sub



Private Sub Img_Run_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblRun.FontBold = True
End Sub

Private Sub LblBill_Click()
Call Img_BillOut_Click
End Sub

Private Sub LblBillHopper_Click()
Call Img_IN_Click
End Sub

Private Sub Lblbutton_Click(Index As Integer)
    Select Case Index
        Case 0  '종료
            Unload Me
        Case 1  '결제내역
            frmResult.Show 1
            Me.MousePointer = 0
            Call DataLogger("[HOST Button]    " & "결제내역 열림")
        Case 3  'APS명령
            frmApsCmd.Show 1
            Me.MousePointer = 0
            Call DataLogger("[HOST Button]    " & "APS CMD 화면 열림")
        
        Case 4  'CCTV(RTSP)
            Call Load_IPCamera
            Me.MousePointer = 0
            Call DataLogger("[HOST Button]    " & "CCTV 화면 열림")
    End Select
End Sub


Private Sub Load_IPCamera()
    Dim rs As Recordset
    Dim idx As Integer
    Dim i As Integer
    
    
    For i = 0 To UBound(Glo_FrmIPCameraPlayer) 'MAX_LANE_COUNT - 1
        If (Not Glo_FrmIPCameraPlayer(i) Is Nothing) Then '만들어져 있다면
            Unload Glo_FrmIPCameraPlayer(i)
            'Set Glo_FrmIPCameraPlayer(i) = Nothing
        End If
    Next i
    
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * From tb_CCTV WHERE HostStream_YN='Y' ", adoConn
    
    Do While Not (rs.EOF)
        If (Len(rs!url) > 0) Then
            If (idx < MAX_LANE_COUNT) Then
                Set Glo_FrmIPCameraPlayer(idx) = New FormIPCameraPlayer
                
                If (Not Glo_FrmIPCameraPlayer(idx) Is Nothing) Then '만들어져 있다면
                    Call Glo_FrmIPCameraPlayer(idx).Play(rs!url, rs!Comments)
                End If
                
                Glo_FrmIPCameraPlayer(idx).Show 0
                idx = idx + 1
            End If
        End If
        
        rs.MoveNext
        
    Loop
    Set rs = Nothing
    
    If (idx = 0) Then
        Lblbutton(4).Visible = False
        Lblbutton(4).Enabled = False
    Else
        Lblbutton(4).Visible = True
        Lblbutton(4).Enabled = True
    End If
    
End Sub

Private Sub Lblbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lblbutton(Index).FontBold = True
End Sub

Private Sub LblCoin_Click()
    Call Img_CoinOut_Click
End Sub

Private Sub LblRun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LblRun.FontBold = True
End Sub

Private Sub LblRun_Click()

    Dim sCmdName As String
    Dim iVal As Long
    
    sCmdName = cmb_Aps.text
    iVal = Val(TxtVal)
    
    If sCmdName = "시간할인" Or sCmdName = "금액할인" Or sCmdName = "%   할인" Or sCmdName = "지정요금" Then
        If (iVal = 0) Then
            Exit Sub
        End If
    End If
    
    Select Case sCmdName
        Case "시간할인"
            Glo_APSCMD_Str = CM_HOUR & CStr(iVal)
            
        Case "금액할인"
            Glo_APSCMD_Str = CM_WON & CStr(iVal)
            
        Case "%   할인"
            Glo_APSCMD_Str = CM_PER & CStr(iVal)
            
        Case "전액할인"
            Glo_APSCMD_Str = CM_PER & "100"
            
        Case "1일 주차요금"
            Glo_APSCMD_Str = CM_DAY
            
        Case "지정요금"
            Glo_APSCMD_Str = CM_DATE & CStr(iVal)
        
        Case "초기화면"
            Glo_APSCMD_Str = CM_INITAL
            
        Case "차단기 열림"
            Glo_APSCMD_Str = CM_GATE
            
        Case "영수증"
            Glo_APSCMD_Str = CM_PRINT
            
        Case "지정시간"
            Glo_APSCMD_Str = CM_DATE & Format(DTPicker1.value, "yyyymmdd") & Format(DTPicker2.value, "hhnnss")
    End Select
    
    Call CMD_Connect(sCmdName)
    Beep

End Sub

Private Sub Img_Run_Click()
    Call LblRun_Click
End Sub

Private Sub CMD_Connect(ByVal sCmdName As String)
Dim bData() As Byte

On Error GoTo Err_P

    If (CmdS_Sock.State <> sckClosed) Then
        CmdS_Sock.Close
    End If

    Select Case sCmdName
        Case "영수증", "지정시간" '초기화면 5888
            CmdS_Sock.Connect Glo_Aps_IP, Glo_Aps_PORT
        Case Else   '계산화면 5889
            CmdS_Sock.Connect Glo_Aps_IP, Glo_APSCMD_Port
    End Select
    
Exit Sub

Err_P:
    Call DataLogger("[CMD_Connect] Err_Msg : " & Err.Description)

End Sub

Private Sub CmdS_Sock_Connect()
    Dim sdata As String
    Dim bData() As Byte
    Dim i As Integer
    
On Error GoTo Err_P
    
    sdata = Glo_APSCMD_Str
    ReDim bData(Len(sdata) - 1) As Byte
    
    bData = StrConv(sdata, vbFromUnicode)
    CmdS_Sock.SendData bData
    
    Call DataLogger("[APS CMD SND]  SND : " & Glo_APSCMD_Str)
    'Call LISTBOX_PutString(List_OP, " SND : " & Glo_APSCMD_Str)
    Glo_APSCMD_Str = ""

Exit Sub

Err_P:
    Call DataLogger(" [CMDSock_Connect Proc] Err_Msg : " & Err.Description)

End Sub

Private Sub CmdS_Sock_DataArrival(ByVal bytesTotal As Long)
    Dim rMsg As String
    Dim B() As Byte
    Dim Ret As Integer
    Dim i As Integer
    Dim sdata As String

On Error GoTo Err_P

    ReDim B(bytesTotal - 1)
    
    CmdS_Sock.GetData B(), vbArray + vbByte, bytesTotal
    For i = 0 To bytesTotal - 1
        If (B(i) >= &H80) Then
            rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
            i = i + 1
        Else
            rMsg = rMsg & Chr$(B(i))
        End If
    Next i
    
    Call DataLogger("[APS CMD RCV]  RCV : " & rMsg)
    'Call LISTBOX_PutString(List_OP, " RCV : " & rMsg)
    
    CmdS_Sock.Close

Exit Sub

Err_P:
    Call DataLogger(" [APS CMD RCV] Err_Msg : " & Err.Description)

End Sub


Public Sub APS_PutImage(ByVal sCarNo As String, ByVal ImgFile As String)

    If (ImgFile <> "") Then
        If (IsFile(ImgFile) = True) Then
            Image2.Picture = LoadPicture(ImgFile)
        Else
            Image2.Picture = LoadPicture(App.Path & "\NoCar.jpg")
        End If
    End If
End Sub














