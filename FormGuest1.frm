VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FormGuest1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "FormGuest1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FormGuest1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ComboBox cmb_GHo 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   10  '한글 
      Left            =   1320
      TabIndex        =   3
      Text            =   "(주)자우텍"
      Top             =   3030
      Width           =   2025
   End
   Begin VB.TextBox txt_GHo 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   10  '한글 
      Left            =   8670
      TabIndex        =   7
      Text            =   "개발팀"
      Top             =   3240
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox txt_GObject 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   10  '한글 
      Left            =   1320
      TabIndex        =   1
      Text            =   "친척 방문"
      Top             =   1650
      Width           =   2025
   End
   Begin VB.TextBox txt_GName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   10  '한글 
      Left            =   1320
      TabIndex        =   4
      Text            =   "홍길동"
      Top             =   3690
      Width           =   2025
   End
   Begin VB.TextBox txt_GCarno 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   10  '한글 
      Left            =   1320
      TabIndex        =   0
      Text            =   "서울01가1234"
      Top             =   990
      Width           =   2025
   End
   Begin VB.ComboBox cmb_GDong 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   10  '한글 
      Left            =   1320
      TabIndex        =   2
      Text            =   "(주)자우텍"
      Top             =   2340
      Width           =   2025
   End
   Begin VB.TextBox txt_GTel 
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
      IMEMode         =   10  '한글 
      Left            =   1320
      TabIndex        =   5
      Text            =   "010-0000-4444"
      Top             =   4350
      Width           =   2025
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   765
      Left            =   1890
      TabIndex        =   6
      ToolTipText     =   "저장하고 차단기가 열립니다."
      Top             =   5070
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "입 력"
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormGuest1.frx":A506
   End
   Begin VB.Label lbl_GuestLaneName 
      BackColor       =   &H00404040&
      Caption         =   "Label1"
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
      Height          =   300
      Left            =   2340
      TabIndex        =   42
      Top             =   225
      Width           =   3645
   End
   Begin VB.Label lbl_TitleHo 
      BackStyle       =   0  '투명
      Caption         =   "부     서"
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
      Height          =   435
      Left            =   135
      TabIndex        =   40
      Top             =   3060
      Width           =   1080
   End
   Begin VB.Label Lbl_FuncTextF12 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   39
      Top             =   4110
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF11 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   38
      Top             =   3825
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF10 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   37
      Top             =   3540
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF9 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   36
      Top             =   3255
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF8 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   35
      Top             =   2970
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF7 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   34
      Top             =   2685
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncF12 
      BackStyle       =   0  '투명
      Caption         =   "F12 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   33
      Top             =   4110
      Width           =   525
   End
   Begin VB.Label Lbl_FuncF11 
      BackStyle       =   0  '투명
      Caption         =   "F11 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   32
      Top             =   3825
      Width           =   525
   End
   Begin VB.Label Lbl_FuncF10 
      BackStyle       =   0  '투명
      Caption         =   "F10 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   31
      Top             =   3540
      Width           =   525
   End
   Begin VB.Label Lbl_FuncF9 
      BackStyle       =   0  '투명
      Caption         =   "F9 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   30
      Top             =   3255
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF8 
      BackStyle       =   0  '투명
      Caption         =   "F8 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   29
      Top             =   2970
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF7 
      BackStyle       =   0  '투명
      Caption         =   "F7 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   28
      Top             =   2685
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF1 
      BackStyle       =   0  '투명
      Caption         =   "F1 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   27
      Top             =   975
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF2 
      BackStyle       =   0  '투명
      Caption         =   "F2 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   26
      Top             =   1260
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF3 
      BackStyle       =   0  '투명
      Caption         =   "F3 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   25
      Top             =   1545
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF4 
      BackStyle       =   0  '투명
      Caption         =   "F4 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   24
      Top             =   1830
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF5 
      BackStyle       =   0  '투명
      Caption         =   "F5 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   23
      Top             =   2115
      Width           =   435
   End
   Begin VB.Label Lbl_FuncF6 
      BackStyle       =   0  '투명
      Caption         =   "F6 :"
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
      Height          =   255
      Left            =   3495
      TabIndex        =   22
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Lbl_FuncTextF1 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   21
      Top             =   975
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF2 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   20
      Top             =   1260
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF3 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   19
      Top             =   1545
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF4 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   18
      Top             =   1830
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF5 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   17
      Top             =   2115
      Width           =   1605
   End
   Begin VB.Label Lbl_FuncTextF6 
      BackStyle       =   0  '투명
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
      Height          =   255
      Left            =   4095
      TabIndex        =   16
      Top             =   2400
      Width           =   1605
   End
   Begin VB.Label lbl_GuestImg 
      BackColor       =   &H00000000&
      Caption         =   "ImagePath"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   0
      TabIndex        =   15
      Top             =   5310
      Visible         =   0   'False
      Width           =   6240
   End
   Begin VB.Label lbl_TitleObject 
      BackStyle       =   0  '투명
      Caption         =   "방문목적"
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
      Height          =   435
      Left            =   135
      TabIndex        =   14
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Label lbl_TitleDong 
      BackStyle       =   0  '투명
      Caption         =   "회 사 명"
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
      Height          =   435
      Left            =   135
      TabIndex        =   13
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Label lbl_TitleName 
      BackStyle       =   0  '투명
      Caption         =   "방 문 자"
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
      Height          =   435
      Left            =   135
      TabIndex        =   12
      Top             =   3750
      Width           =   1080
   End
   Begin VB.Label lbl_TitleCarno 
      BackStyle       =   0  '투명
      Caption         =   "차량번호"
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
      Height          =   435
      Left            =   135
      TabIndex        =   11
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label lbl_TitleGuest 
      BackStyle       =   0  '투명
      Caption         =   " 방문차량 처리 : "
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
      Height          =   300
      Left            =   150
      TabIndex        =   10
      Top             =   225
      Width           =   2025
   End
   Begin VB.Line LineGuest 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   120
      X2              =   6100
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label lbl_TitleTel 
      BackStyle       =   0  '투명
      Caption         =   "연 락 처"
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
      Height          =   435
      Left            =   135
      TabIndex        =   9
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Label lbl_GuestPassDate 
      BackColor       =   &H00000000&
      Caption         =   "PassDate"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   6240
   End
   Begin VB.Label LabelGuest 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   6675
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   6240
   End
End
Attribute VB_Name = "FormGuest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GuestGateNo As Integer
Private PrintModel As String
Private PrintPort As String
Private F_Handle As Long



Private Sub cmd_Menu_Click()
    '방문증 출력
    If (Guest_Error_Check = True) Then
        Call Guest_In_Manual_Proc
    Else
        MsgBox "방문증 정보를 정확하게 입력하세요!"
        Me.MousePointer = 0
        Exit Sub
    End If
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()

    RemoveCancelMenuItem Me ' 닫기버튼 없애기
    
    Call FormOnTop(Me.hwnd, True) '최상위 폼
    
    Call InitFormField

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub InitFormField()
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_P
    
    If (Glo_User_Type = "구분1/구분2") Then
        lbl_TitleDong = "구    분1"
        lbl_TitleHo = "구    분2"
    Else
        lbl_TitleDong = "     동   "
        lbl_TitleHo = " 호    수"
    End If
    

    Call SetGuestFormFunc
   
    Call ClearField
    
    
    qry = "SELECT DONG From tb_guest_log Group By DONG"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    
    cmb_GDong.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GDong.AddItem "" & rs!Dong
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    qry = "SELECT HO From tb_guest_log Group By HO"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    
    cmb_GHo.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GHo.AddItem "" & rs!Ho
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    Exit Sub
    
Err_P:
    Set rs = Nothing
    'Call DataLogger("InitFormField:" & err.Description)
End Sub


Private Sub ClearField()
    txt_GCarno.text = ""
    txt_GObject.text = ""
    cmb_GDong.text = ""
    cmb_GHo.text = ""
    txt_GHo.text = ""
    txt_GName.text = ""
    txt_GTel.text = ""
    lbl_GuestPassDate = ""
    lbl_GuestImg = ""
End Sub


Private Sub SetGuestFormFunc()
    F_Key1 = Get_Ini("System Config", "F1", "")
    F_Key2 = Get_Ini("System Config", "F2", "")
    F_Key3 = Get_Ini("System Config", "F3", "")
    F_Key4 = Get_Ini("System Config", "F4", "")
    F_Key5 = Get_Ini("System Config", "F5", "")
    F_Key6 = Get_Ini("System Config", "F6", "")
    F_Key7 = Get_Ini("System Config", "F7", "")
    F_Key8 = Get_Ini("System Config", "F8", "")
    F_Key9 = Get_Ini("System Config", "F9", "")
    F_Key10 = Get_Ini("System Config", "F10", "")
    F_Key11 = Get_Ini("System Config", "F11", "")
    F_Key12 = Get_Ini("System Config", "F12", "")
    
    Lbl_FuncTextF1.Caption = F_Key1
    Lbl_FuncTextF2.Caption = F_Key2
    Lbl_FuncTextF3.Caption = F_Key3
    Lbl_FuncTextF4.Caption = F_Key4
    Lbl_FuncTextF5.Caption = F_Key5
    Lbl_FuncTextF6.Caption = F_Key6
    Lbl_FuncTextF7.Caption = F_Key7
    Lbl_FuncTextF8.Caption = F_Key8
    Lbl_FuncTextF9.Caption = F_Key9
    Lbl_FuncTextF10.Caption = F_Key10
    Lbl_FuncTextF11.Caption = F_Key11
    Lbl_FuncTextF12.Caption = F_Key12
End Sub


Public Sub SetGuestName(name As String)
    lbl_GuestLaneName.Caption = name
End Sub

Public Sub SetAutoMode(autoMode As String, name As String)
    
    lbl_GuestLaneName.Caption = name
    
    If (autoMode = "Y") Then
        cmd_menu.Enabled = False
    Else
        cmd_menu.Enabled = True
    End If
End Sub

Public Sub SetGateNo(gateNo As Integer, prtModel As String, prtPort)
    
On Error GoTo Err_P
    
    GuestGateNo = gateNo
            
    Select Case GuestGateNo
        Case 0
            lbl_GuestLaneName.Caption = LANE1_Name
        Case 1
            lbl_GuestLaneName.Caption = LANE2_Name
        Case 2
            lbl_GuestLaneName.Caption = LANE3_Name
        Case 3
            lbl_GuestLaneName.Caption = LANE4_Name
        Case 4
            lbl_GuestLaneName.Caption = LANE5_Name
        Case 5
            lbl_GuestLaneName.Caption = LANE6_Name
    End Select
    
    Select Case Glo_Screen_No
        Case 1
            Left = FrmG1.Left + FrmG1.ImageIn(GuestGateNo).Left + 200
            Top = FrmG1.Top + FrmG1.ImageIn(GuestGateNo).Top + 500
        Case 2
            Left = Jung.Left + Jung.Frame1(GuestGateNo).Left
            Top = Jung.Top + Jung.Frame1(GuestGateNo).Top + 6300
        Case 4
            Left = FrmG4Mini.Left + FrmG4Mini.ImageIn(GuestGateNo).Left + FrmG4Mini.ImageIn(GuestGateNo).width * GuestGateNo
            Top = FrmG4Mini.Top + FrmG4Mini.ImageIn(GuestGateNo).Top + 6300
        Case 6
            Left = FrmG6_23.Left + FrmG6_23.ImageIn(GuestGateNo).width * Int(GuestGateNo Mod 3) + 200
            Top = FrmG6_23.Top + (FrmG6_23.ImageIn(GuestGateNo).height * Int(GuestGateNo / 3)) + 1400
        
    End Select
    
    PrintModel = prtModel
    PrintPort = prtPort

        
    Exit Sub
Err_P:
    
    
End Sub



'방문증 필수 입력 데이터 확인
Private Function Guest_Error_Check()
    Dim Error_Flag
    Dim i As Integer

On Error GoTo Err_P

    Error_Flag = True

    Select Case LenH(txt_GCarno.text)
        Case 4, 8, 9, 11, 12
        
        Case Else
            Error_Flag = False
    End Select

    Guest_Error_Check = Error_Flag
    
    Exit Function
    
Err_P:
    DataLogger ("Guest Error Check:" & Err.Description)

End Function


Public Sub Guest_Inputdata(carno As String, passData As String, img_path As String)
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
    Set rs = New ADODB.Recordset
    qry = "SELECT * FROM tb_reg WHERE CAR_NO = '" & carno & "'"
    bQryResult = DataBaseQuery(rs, adoConn, qry, NWERR_GATE_STAY, 0)
    If (bQryResult = False) Then
        DataLogger ("[Guest_Inputdata]    " & "네트워크 및 DB 점검바랍니다, 입출차기록 저장실패_차단기 자동 열림")
        Exit Sub
    End If
    
    If (rs.EOF) Then
        txt_GCarno.text = carno
        lbl_GuestPassDate.Caption = passData
        lbl_GuestImg.Caption = img_path
    End If
    
    Set rs = Nothing
    
End Sub

'방문객 수동 버튼처리(입구)
Public Sub Guest_In_Manual_Proc()

On Error Resume Next

    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim GGate As Integer
    Dim GCarno, GObject, GName, GDong, GHo, GHo2, GTel, GImage, GPassDate, GEtc, GEtc2, GEtc3 As String

    Call FormOnTop(Me.hwnd, True) '최상위 폼
    
    GGate = GuestGateNo
    GCarno = MidH(Trim(txt_GCarno.text), 1, 16)
    GObject = MidH(Trim(txt_GObject.text), 1, 64)
    GName = MidH(Trim(txt_GName.text), 1, 32)
    GDong = MidH(Trim(cmb_GDong.text), 1, 32)
    'GHo = MidH(Trim(txt_GHo.Text), 1, 32)
    GHo = MidH(Trim(cmb_GHo.text), 1, 32)
    GHo2 = ""
    GEtc = "수동입차"
    GEtc2 = ""
    GEtc3 = ""
    GTel = MidH(Trim(txt_GTel.text), 1, 32)
    If (Len(lbl_GuestImg) > 0) Then
        GImage = Slash_Conv(Trim(lbl_GuestImg.Caption))
    Else
        GImage = ""
    End If
    
    
    If (Len(lbl_GuestPassDate.Caption) > 0) Then
        GPassDate = lbl_GuestPassDate.Caption
    Else
        'GPassDate = Format(Now, "yyyy-mm-dd hh:nn:ss") & Format(Timer * 1000 Mod 1000, " 000")
        GPassDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    End If
    
    'adoConn.Execute "INSERT INTO tb_guest_log (GUEST_NO, CAR_NO, OBJECT, DONG, HO, HO2, NAME,TEL,ETC,ETC2,ETC3,DT_IN,IN_GATE,IN_IMAGE_PATH,DT_OUT,OUT_GATE,OUT_IMAGE_PATH,REG_DATE,DT_UPDATE,PARK_TIME ) VALUES ('', '" & GCarno & "','" & GObject & "','" & GDong & "','" & GHo & "', '" & GHo2 & "', '" & GName & "','" & GTel & "','" & GEtc & "','" & GEtc2 & "','" & GEtc3 & "','" & GPassDate & "','" & GGate & "','" & GImage & "','','','','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "','', 0)"
    adoConn.Execute "INSERT INTO tb_guest_log (GUEST_NO, CAR_NO, OBJECT, DONG, HO, HO2, NAME,TEL,ETC,ETC2,ETC3,GUBUN,DT_IN,IN_GATE,IN_IMAGE_PATH,DT_OUT,OUT_GATE,OUT_IMAGE_PATH,REG_DATE,DT_UPDATE,PARK_TIME ) VALUES ('', '" & GCarno & "','" & GObject & "','" & GDong & "','" & GHo & "', '" & GHo2 & "', '" & GName & "','" & GTel & "','" & GEtc & "','" & GEtc2 & "','" & GEtc3 & "','입구','" & GPassDate & "','" & GGate & "','" & GImage & "','','','','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "','', 0)"
    
    Call DataLogger("[ 방문객 등록 ] 차량번호 : " & GCarno & "   방문지 : " & GDong)
    
    '영수증출력
    If (PrintModel <> "NONE") Then
    
        Call Visit_Print(PrintPort, GGate)
        
        '차단기오픈 시간지연
        If (Glo_Guest_Gate_OpenDelay(GGate) > 0) Then
            Call Delay_Time(Glo_Guest_Gate_OpenDelay(GGate))
        End If
        
    End If
    
    
    '차단기오픈
    Call Relay_Out(0, CInt(GGate))

    'Me.MousePointer = 0
    
    Call ClearField
    Call InitFormField
    
End Sub



'방문객 자동처리(입구)
Public Sub Guest_In_Auto_Proc(sCarNo As String, sPassDate As String, sImagePath As String, sInOut As String)
   
On Error GoTo Err_P
        
        Dim bQryResult As Boolean
        Dim rs As Recordset
        Dim qry As String
        Dim GGate, GCarno, GObject, GName, GDong, GHo, GHo2, GTel, GImage, GPassDate, GEtc, GEtc2, GEtc3 As String

        'Call FormOnTop(Me.hwnd, True) '최상위 폼
        


        GGate = GuestGateNo
        GCarno = MidH(Trim(txt_GCarno.text), 1, 16)
        GObject = MidH(Trim(txt_GObject.text), 1, 64)
        GName = MidH(Trim(txt_GName.text), 1, 32)
        GDong = MidH(Trim(cmb_GDong.text), 1, 32)
        GHo = MidH(Trim(txt_GHo.text), 1, 32)
        GHo2 = ""
        GEtc = "자동입차"
        GEtc2 = ""
        GEtc3 = ""
        GTel = MidH(Trim(txt_GTel.text), 1, 32)
        GImage = Slash_Conv(Trim(lbl_GuestImg.Caption))
        GPassDate = lbl_GuestPassDate.Caption


        If (sInOut = "입구") Then
            'If (sFreePass = "Y") Then
                adoConn.Execute "INSERT INTO tb_guest_log (GUEST_NO, CAR_NO, OBJECT, DONG, HO, HO2, NAME,TEL,ETC,ETC2,ETC3,GUBUN,DT_IN,IN_GATE,IN_IMAGE_PATH,DT_OUT,OUT_GATE,OUT_IMAGE_PATH,REG_DATE,DT_UPDATE,PARK_TIME ) VALUES ('', '" & GCarno & "','" & GObject & "','" & GDong & "','" & GHo & "', '" & GHo2 & "', '" & GName & "','" & GTel & "','" & GEtc & "','" & GEtc2 & "','" & GEtc3 & "','입구','" & GPassDate & "','" & GGate & "','" & GImage & "','','','','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "','', 0)"
                
                Call DataLogger("[ 방문객 자동등록 ] 차량번호 : " & GCarno & "   방문지 : " & GDong)
                
                'LPRIN_PRoc에서 차단기 자동열림 실행됐으므로 여기서는 주석처리함
                'Call Relay_Out(0, CInt(GGate))
                'Me.MousePointer = 0
        End If
        
    'Call ClearField

    Exit Sub

Err_P:
    Set rs = Nothing
    Call DataLogger("Guest Auto Proc:" & Err.Description)
End Sub



'방문객 자동처리(출구)
Public Sub Guest_Out_Auto_Proc(sGateNo As String, sCarNo As String, sPassDate As String, sImagePath As String, sInOut As String)

On Error GoTo Err_P
        
        Dim bQryResult As Boolean
        Dim rs As Recordset
        Dim qry As String
        Dim GGate, GCarno, GImage, GPassDate As String
        Dim GParkTime As Long
        
        GGate = sGateNo
        GCarno = Trim(sCarNo)
        GImage = Slash_Conv(Trim(sImagePath))
        GPassDate = sPassDate
        
        
        If (sInOut = "출구") Then
            'QRY = "SELECT IDX AS SEQ,DT_IN From tb_guest_log WHERE CAR_NO = '" & GCarno & "' AND  PARK_TIME = ''  ORDER BY IDX DESC limit 1" '느림
            qry = "SELECT IDX AS SEQ,DT_IN From tb_guest_log WHERE CAR_NO = '" & GCarno & "' AND  PARK_TIME = 0  ORDER BY REG_DATE DESC limit 1"
            Set rs = New ADODB.Recordset
            bQryResult = DataBaseQuery(rs, adoConn, qry, False)

            If Not rs.EOF Then
'''                GParkTime = DateDiff("n", Left(rs!DT_IN, 19), Left(GPassDate, 19))
'''                adoConn.Execute "UPDATE tb_guest_log set DT_OUT = '" & GPassDate & "', OUT_GATE = '" & GGate & "', OUT_IMAGE_PATH = '" & GImage & "', DT_UPDATE = '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "', PARK_TIME = " & GParkTime & " WHERE IDX = " & rs!SEQ & " "
                adoConn.Execute "INSERT INTO tb_guest_log (GUEST_NO, CAR_NO, OBJECT, DONG, HO, HO2, NAME,TEL,ETC,ETC2,ETC3,GUBUN,DT_IN,IN_GATE,IN_IMAGE_PATH,DT_OUT,OUT_GATE,OUT_IMAGE_PATH,REG_DATE,DT_UPDATE,PARK_TIME ) VALUES ('', '" & GCarno & "','','','', '', '','','','','','출구','','','','" & GPassDate & "','" & GGate & "','" & GImage & "','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "','', 0)"
            End If
            Set rs = Nothing
        
        End If

    Exit Sub

Err_P:
    Set rs = Nothing
    Call DataLogger("Guest Auto Proc:" & Err.Description)
End Sub




Public Sub Visit_Print(Port As String, Gate As Integer)
Dim tmp As Boolean
Dim Blank As String
Blank = "    "

tmp = Open_Printer(F_Handle, Gate)
If (tmp = False) Then
    Exit Sub
End If

With Jung
    'Call Print_String("----------------------------------------", 0)
    'Call Print_String("발급일시", 1)
    Call Print_String(Blank & "=====================================", 0, Gate, F_Handle)
    Call Print_String(Blank & "  [ 방문객 입차증 ]", 1, Gate, F_Handle)
    Call Print_String(Blank & "=====================================", 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    Call Print_String(Blank & "날 짜 : " & Format(Now, "yyyy-mm-dd"), 0, Gate, F_Handle)
    Call Print_String(Blank & "시 간 : " & Format(Now, "hh:nn"), 0, Gate, F_Handle)
    Call Print_String(Blank & "-------------------------------------", 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    'Call Print_String("번  호:" & ticket, 1)
    Call Print_String(Blank & "차량번호 : " & txt_GCarno, 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    
    If (Glo_User_Type = "구분1/구분2") Then
            Call Print_String(Blank & "방 문 지 : " & cmb_GDong & " ,  " & cmb_GHo, 0, Gate, F_Handle)
    Else
            Call Print_String(Blank & "방 문 지 : " & cmb_GDong & " 동  " & cmb_GHo & " 호", 0, Gate, F_Handle)
    End If
    Call Paper_Feed(1, Gate, F_Handle)
    
    Call Print_String(Blank & "방문목적 : " & txt_GObject, 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)

    Call Print_String(Blank & "연 락 처 : " & txt_GTel, 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    
    Call Print_String(Blank & "방문자명 : " & txt_GName, 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    
    'Call Print_String("비    고 : " & txt_GEtc, 0)
    'Call Paper_Feed(1)
    Call Print_String(Blank & "=====================================", 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    Call Print_String(Blank & "* 방문증은 차량 전면에 비치해", 0, Gate, F_Handle)
    Call Print_String(Blank & "     주시기 바랍니다.", 0, Gate, F_Handle)
    Call Print_String(Blank & "* 방문증이 없을시, 불법주차로 간주합니다.", 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    Call Print_String(Blank & "=====================================", 0, Gate, F_Handle)
    Call Paper_Feed(1, Gate, F_Handle)
    'Call Paper_Feed(1)
    'Call Print_String("일산자이 아파트", 1)
    Call Paper_Feed(1, Gate, F_Handle)
    Call Paper_Cut(1, Gate, F_Handle)
    'Call Paper_Feed(1)
    '========================================", 0)
    'Call Paper_Feed(1)
    '12345678901234567890123456789012345678901234567890
    '입차일시:yy-mm-dd hh:nn입차일시:yy-mm-dd hh:nn
    '========================================", 0)
    '----------------------------------------", 0)
'''    Call Print_String("========================================", 0, Gate, F_Handle)
'''    Call Print_String("  [ 방문객 입차증 ]", 1, Gate, F_Handle)
'''    Call Print_String("========================================", 0, Gate, F_Handle)
    tmp = CloseHandle(F_Handle)

End With

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call FormOnTop(Me.hwnd, False)
End Sub



Private Sub txt_GObject_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
        KeyCode = 0
    End If

    Select Case KeyCode
        Case 112
             txt_GObject = Trim(F_Key1)
        Case 113
             txt_GObject = Trim(F_Key2)
        Case 114
             txt_GObject = Trim(F_Key3)
        Case 115
             txt_GObject = Trim(F_Key4)
        Case 116
             txt_GObject = Trim(F_Key5)
        Case 117
             txt_GObject = Trim(F_Key6)
        Case 118
             txt_GObject = Trim(F_Key7)
        Case 119
             txt_GObject = Trim(F_Key8)
        Case 120
             txt_GObject = Trim(F_Key9)
        Case 121
             txt_GObject = Trim(F_Key10)
        Case 122
             txt_GObject = Trim(F_Key11)
        Case 123
             txt_GObject = Trim(F_Key12)
        Case Else
    End Select
    cmb_GDong.SetFocus

    KeyCode = 0

End Sub




