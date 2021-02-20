VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGuestRegCar 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   11250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11250
   ScaleWidth      =   13245
   Begin VB.CheckBox chk_GuestRegUse 
      BackColor       =   &H00404040&
      Caption         =   "사용여부"
      Enabled         =   0   'False
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   11235
      TabIndex        =   19
      Top             =   4470
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   " 방문객 정보"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4605
      Left            =   135
      TabIndex        =   16
      Top             =   6420
      Width           =   12975
      Begin VB.TextBox txt_Tel 
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
         Left            =   2190
         TabIndex        =   27
         Text            =   "010-0000-0000"
         Top             =   3105
         Width           =   2505
      End
      Begin VB.TextBox txt_Name 
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
         Left            =   2190
         TabIndex        =   25
         Text            =   "홍길동"
         Top             =   2535
         Width           =   2505
      End
      Begin VB.TextBox txt_Carno 
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
         Left            =   2190
         TabIndex        =   23
         Text            =   "서울12가3456"
         Top             =   825
         Width           =   2505
      End
      Begin VB.TextBox txt_Dong 
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
         Left            =   2190
         TabIndex        =   4
         Text            =   "101"
         Top             =   1395
         Width           =   2505
      End
      Begin VB.TextBox txt_Ho 
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
         Left            =   2190
         TabIndex        =   5
         Text            =   "202"
         Top             =   1965
         Width           =   2505
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   690
         Left            =   6135
         TabIndex        =   9
         Top             =   3735
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "초기화"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCar.frx":0000
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   690
         Left            =   7800
         TabIndex        =   6
         Top             =   3735
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "등록"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCar.frx":0351
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   690
         Left            =   9465
         TabIndex        =   7
         Top             =   3735
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "수정"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCar.frx":06A2
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   690
         Left            =   11130
         TabIndex        =   8
         Top             =   3735
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "삭제"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCar.frx":09F3
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   345
         Left            =   6225
         TabIndex        =   35
         Top             =   1395
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   9.75
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
         Format          =   138805248
         CurrentDate     =   36927
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   345
         Left            =   9690
         TabIndex        =   36
         Top             =   1395
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   9.75
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
         Format          =   138805248
         CurrentDate     =   36927
      End
      Begin VB.Label lbl_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   20.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   1
         Left            =   8865
         TabIndex        =   30
         Top             =   1395
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "방문예약종료날짜"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   9690
         TabIndex        =   29
         Top             =   885
         Width           =   2160
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "방문예약시작날짜"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6210
         TabIndex        =   28
         Top             =   885
         Width           =   2160
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "연락처"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1260
         TabIndex        =   26
         Top             =   3180
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이름"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1530
         TabIndex        =   24
         Top             =   2610
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "차량번호"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   990
         TabIndex        =   22
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label lbl_dong 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "동"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   1470
         Width           =   270
      End
      Begin VB.Label lbl_ho 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "호수"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1530
         TabIndex        =   17
         Top             =   2055
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   " 검색 "
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
      Height          =   1395
      Left            =   135
      TabIndex        =   13
      Top             =   4695
      Width           =   12975
      Begin VB.TextBox txt_SrchCarno 
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
         Left            =   2715
         TabIndex        =   21
         Text            =   "서울12가1234"
         Top             =   735
         Width           =   2010
      End
      Begin VB.ComboBox cmb_Ho 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8025
         TabIndex        =   2
         Text            =   "cmb_Ho"
         Top             =   810
         Width           =   1290
      End
      Begin VB.ComboBox cmb_Dong 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6180
         TabIndex        =   0
         Text            =   "cmb_Dong"
         Top             =   810
         Width           =   1290
      End
      Begin Threed.SSCommand cmd_Search 
         Height          =   690
         Left            =   11205
         TabIndex        =   3
         Top             =   585
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "검 색"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCar.frx":0D44
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   2715
         TabIndex        =   31
         Top             =   300
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   9.75
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
         Format          =   138805248
         CurrentDate     =   36927
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   6180
         TabIndex        =   32
         Top             =   300
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "나눔고딕"
            Size            =   9.75
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
         Format          =   138805248
         CurrentDate     =   36927
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~ 종료날짜"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   4905
         TabIndex        =   34
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "방문예약 시작날짜"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   600
         TabIndex        =   33
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "차량번호"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   1680
         TabIndex        =   20
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "동"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   7545
         TabIndex        =   15
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "호"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   9405
         TabIndex        =   14
         Top             =   840
         Width           =   270
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3000
      Left            =   135
      TabIndex        =   1
      Top             =   1185
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   0
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   570
      Left            =   11835
      TabIndex        =   11
      Top             =   150
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "닫기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestRegCar.frx":1095
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   570
      Left            =   10470
      TabIndex        =   10
      Top             =   150
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "저장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestRegCar.frx":13E6
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9150
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   13110
      Y1              =   780
      Y2              =   765
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '투명
      Caption         =   "사전방문신청 조회/등록"
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
      Index           =   0
      Left            =   165
      TabIndex        =   12
      Top             =   300
      Width           =   4470
   End
End
Attribute VB_Name = "FrmGuestRegCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nSeq As Long
Dim sPass_YN As String
'Const MAX_PARKDAY As Integer = 3 '오늘 제외하고 주차가능기간(ex: 3 = 10/17 ~ 10/20)
Dim MAX_PARKDAY As Integer

'Private Sub Calendar_Bgn_Click()
'    If (Format(Calendar_Bgn.value, "YYYY-MM-DD") < Format(Now, "YYYY-MM-DD")) Then
'        MsgBox "예약시작일은 오늘날짜부터 지정할 수 있습니다!!"
'        Calendar_Bgn.value = Format(Now, "YYYY-MM-DD")
'    End If
'End Sub
'Private Sub Calendar_End_Click()
'    If (Format(Calendar_End.value, "YYYY-MM-DD") < Format(Calendar_Bgn.value, "YYYY-MM-DD")) Then
'        MsgBox "예약종료일은 예약시작일보다 같거나 커야 합니다!!"
'        Calendar_End.value = Calendar_Bgn.value
'        Exit Sub
'    End If
'
'    If (Format(Calendar_End.value, "YYYY-MM-DD") > Format(DateAdd("d", (MAX_PARKDAY), Format(Calendar_Bgn.value, "yyyy-mm-dd")), "yyyy-mm-dd")) Then
'        MsgBox "예약종료일은 예약시작일로부터 최대 4일입니다!!"
'        Calendar_End.value = DateAdd("d", (MAX_PARKDAY), Format(Calendar_Bgn, "yyyy-mm-dd"))
'    End If
'End Sub
Private Sub Form_Load()

    Dim sSDate As String
    Dim sEDate As String
    Dim sQry As String
    
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
   
    DTPicker1.value = Now
    'DTPicker2.value = Format(DateAdd("m", (12), Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
    DTPicker2.value = "9999-12-31"
    
    'Calendar_Bgn.value = Format(Now, "YYYY-MM-DD")
    'Calendar_End.value = Format(Now, "YYYY-MM-DD")
    DTPicker3.value = Now
    DTPicker4.value = Now
    
    sSDate = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDate = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    sQry = "SELECT * From tb_guestReg WHERE START_DATE >= '" & sSDate & "' AND END_DATE <= '" & sEDate & "' "
    
    MAX_PARKDAY = Get_MaxParkDay
    Call Clear_Field
    Call ListView1_Draw
    Call ListView1_SQL("SELECT * From tb_guestReg WHERE START_DATE >= '" & sSDate & "' AND END_DATE <= '" & sEDate & "' ")
    
End Sub


Private Sub SSCommand2_Click()
    Unload Me
    'Me.Hide
End Sub

Private Function Get_MaxParkDay()
    
    Dim rs As Recordset
    Dim bQryResult As Boolean
    Dim nMaxParkDay As Integer
    
    nMaxParkDay = 0
    
    If (Glo_GuestReg_YN = "Y") Then
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'GuestCarReg_MaxParkDay' ", False): nMaxParkDay = rs!Content: Set rs = Nothing
    End If
    
    Get_MaxParkDay = nMaxParkDay
    
End Function
Private Sub Clear_Field()
    
    SSCommand4.Enabled = True '등록버튼
    SSCommand5.Enabled = True '수정버튼
    SSCommand6.Enabled = True '삭제버튼
    
    nSeq = -1
    sPass_YN = ""
    
    txt_CarNo = ""
    txt_Dong = ""
    txt_Ho = ""
    txt_Name = ""
    txt_Tel = ""

    'Calendar_Bgn.value = Format(Now, "yyyy-mm-dd")
    'Calendar_End.value = Format(DateAdd("d", (MAX_PARKDAY), Format(Calendar_Bgn.value, "yyyy-mm-dd")), "yyyy-mm-dd")
    DTPicker3.value = Format(Now, "yyyy-mm-dd")
    DTPicker4.value = Format(DateAdd("d", (MAX_PARKDAY), Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
    
    'chk_GuestRegUse.value = 1
    txt_SrchCarno = ""
    Call Set_cmbDong
    Call Set_cmbHo
End Sub
Private Sub ListView1_Draw()
    Dim Column_to_size As Integer

'On Error GoTo Err_p
On Error Resume Next

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " 차량번호      "
    ListView1.ColumnHeaders.Add , , " 동       "
    ListView1.ColumnHeaders.Add , , " 호수     "
    ListView1.ColumnHeaders.Add , , " 이름  "
    ListView1.ColumnHeaders.Add , , " 연락처    "
    ListView1.ColumnHeaders.Add , , " 방문예약시작날짜    "
    ListView1.ColumnHeaders.Add , , " 방문예약종료날짜    "
    ListView1.ColumnHeaders.Add , , " 입차유무    "
    ListView1.ColumnHeaders.Add , , " 등록날짜      "
    
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End Sub

Private Sub ListView1_SQL(qry As String)
    Dim bQryResult As Boolean
    Dim rs As Recordset
'    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Integer
    
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        Exit Sub
    End If

    INDEX_NO = 1
    Do While Not (rs.EOF)
        'Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        Set itmX = ListView1.ListItems.Add(, , "" & rs!SEQ)
        i = 1
        itmX.SubItems(i) = "" & rs!CAR_NO: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_DEPT: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_CLASS: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_NAME: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_PHONE: i = i + 1
        itmX.SubItems(i) = "" & Left(rs!START_DATE, 10): i = i + 1
        itmX.SubItems(i) = "" & Left(rs!END_DATE, 10): i = i + 1
        itmX.SubItems(i) = "" & rs!Pass_YN: i = i + 1
        itmX.SubItems(i) = "" & rs!REG_DATE: i = i + 1

        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    Set rs = Nothing
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim i As Integer
    With ListView1
        For i = 1 To .ColumnHeaders.Count
            If (.ColumnHeaders.Item(i) = ColumnHeader) Then
                .SortKey = i - 1
                .SortOrder = .SortOrder Xor 1
                '.SortOrder = lvwDescending
                .Sorted = True
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Error Resume Next
    
    ListView1.SetFocus
    
    nSeq = ListView1.SelectedItem
    
    txt_CarNo = ListView1.SelectedItem.SubItems(1)
    txt_Dong = ListView1.SelectedItem.SubItems(2)
    txt_Ho = ListView1.SelectedItem.SubItems(3)
    txt_Name = ListView1.SelectedItem.SubItems(4)
    txt_Tel = ListView1.SelectedItem.SubItems(5)
    
    'Calendar_Bgn.value = ListView1.SelectedItem.SubItems(6)
    'Calendar_End.value = ListView1.SelectedItem.SubItems(7)
    DTPicker3 = ListView1.SelectedItem.SubItems(6)
    DTPicker4 = ListView1.SelectedItem.SubItems(7)
    
    
    sPass_YN = ListView1.SelectedItem.SubItems(8)
    If (sPass_YN = "Y") Then '이미 입차차량
        SSCommand4.Enabled = False '등록버튼
        SSCommand5.Enabled = False '수정버튼
        SSCommand6.Enabled = False '삭제버튼
    Else
        SSCommand4.Enabled = True '등록버튼
        SSCommand5.Enabled = True '수정버튼
        SSCommand6.Enabled = True '삭제버튼
    End If
    
End Sub

'초기화
Private Sub SSCommand3_Click()
    Call Clear_Field
End Sub

'등록
Private Sub SSCommand4_Click()

    Dim sUse As String
    
    If (Check_Field = False) Then
        Msg_Box.Label1 = "방문예약차량 입력 오류입니다" & vbCrLf & vbCrLf & "재입력 바랍니다"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    If (Check_NewRecord = False) Then
        Msg_Box.Label1 = "기존 등록차량의 방문예약일과 중복됩니다" & vbCrLf & vbCrLf & "재입력 바랍니다"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    Dim sLog As String
    Dim sNowDT As String
    Dim sInputSDate As String
    Dim sInputEDate As String
    
    sNowDT = Format(Now, "yyyy-mm-dd hh:nn:ss")
    'sInputSDate = Format(Calendar_Bgn.value, "yyyy-mm-dd") & " 00:00:00"
    'sInputEDate = Format(Calendar_End.value, "yyyy-mm-dd") & " 23:59:59"
    sInputSDate = Format(DTPicker3, "yyyy-mm-dd") & " 00:00:00"
    sInputEDate = Format(DTPicker4, "yyyy-mm-dd") & " 23:59:59"
    
    sLog = "방문예약 등록:" & Glo_Login_ID & ":" & txt_CarNo & ""
    adoConn.Execute "insert into tb_guestReg (CAR_NO,CAR_GUBUN,CAR_FEE,DRIVER_NAME,DRIVER_PHONE,DRIVER_DEPT,DRIVER_CLASS,START_DATE,END_DATE,REG_DATE,DAY_ROTATION_YN,LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,PASS_YN) VALUES ( '" & txt_CarNo & "','방문예약','0','" & txt_Name & "','" & txt_Tel & "','" & txt_Dong & "','" & txt_Ho & "','" & sInputSDate & "','" & sInputEDate & "','" & sNowDT & "','적용','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','Y','N','N') "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('방문예약', '" & txt_Dong & "', '" & sLog & "', '" & txt_Ho & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Call DataLogger(sLog)

    Call Clear_Field
    Call ListView1_Draw
    Call ListView1_SQL("SELECT * From tb_guestReg")
End Sub

Private Function Check_Field() As Boolean
    Dim bCheck As Boolean
    bCheck = True
    
    If (IsNumeric(txt_Ho) = False) Then
        txt_Ho = "":     txt_Ho.SetFocus
        bCheck = False
    End If
    
    If (IsNumeric(txt_Dong) = False) Then
        txt_Dong = "":     txt_Dong.SetFocus
        bCheck = False
    End If
    
    If Not ((LenH(txt_CarNo.text) = 11) Or (LenH(txt_CarNo.text) = 12) Or (LenH(txt_CarNo.text) = 8) Or (LenH(txt_CarNo.text) = 9)) Then
        txt_CarNo = "":     txt_CarNo.SetFocus
        bCheck = False
    End If
    
    Check_Field = bCheck
End Function

Private Function Check_NewRecord() As Boolean
    Dim rs As Recordset
    Dim bCheck As Boolean
    Dim sInputSDate As String
    Dim sInputEDate As String
    Dim sDBSDate As String
    Dim sDBEDate As String
    
    bCheck = True
        
    'sInputSDate = Format(Calendar_Bgn.value, "yyyy-mm-dd") & " 00:00:00"
    'sInputEDate = Format(Calendar_End.value, "yyyy-mm-dd") & " 23:59:59"
    sInputSDate = Format(DTPicker3, "yyyy-mm-dd") & " 00:00:00"
    sInputEDate = Format(DTPicker4, "yyyy-mm-dd") & " 23:59:59"
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tb_guestreg WHERE CAR_NO = '" & txt_CarNo & "' ", adoConn
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            sDBSDate = "" & rs!START_DATE
            sDBEDate = "" & rs!END_DATE
            
            If (sInputEDate < sDBSDate Or sInputSDate > sDBEDate) Then
                '신규등록가능
            Else
                bCheck = False
                Exit Do
            End If
        Loop
    End If
    Set rs = Nothing
    
    Check_NewRecord = bCheck
End Function


'수정
Private Sub SSCommand5_Click()
    If (nSeq < 0) Then
        Msg_Box.Label1 = "삭제할 차량을 선택하세요"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    
    MBox.Label3.Caption = txt_CarNo.text
    MBox.Label1.Caption = "방문예약차량 정보를 수정합니다." & vbCrLf & " 수정하시겠습니까?"
    MBox.Label2.Caption = "방문예약차량 수정"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
        If (Check_InCAR(txt_CarNo) = True) Then
            Msg_Box.Label1 = "입차차량은 수정할 수 없습니다."
            Msg_Box.Show 1
            Exit Sub
        Else
            Dim sLog As String
            Dim sSDate, sEDate As String
            'sSDate = Format(Calendar_Bgn.value, "yyyy-mm-dd") & " 00:00:00"
            'sEDate = Format(Calendar_End.value, "yyyy-mm-dd") & " 23:59:59"
            sSDate = Format(DTPicker3, "yyyy-mm-dd") & " 00:00:00"
            sEDate = Format(DTPicker4, "yyyy-mm-dd") & " 23:59:59"
            
            'sLog = "방문예약 수정:" & Glo_Login_ID & ":" & txt_Carno & "(" & Calendar_Bgn.value & "~" & Calendar_End.value & ")"
            sLog = "방문예약 수정:" & Glo_Login_ID & ":" & txt_CarNo & "(" & Format(DTPicker3, "yyyy-mm-dd") & "~" & Format(DTPicker4, "yyyy-mm-dd") & ")"
            adoConn.Execute "UPDATE tb_guestReg SET CAR_NO = '" & txt_CarNo & "', DRIVER_DEPT = '" & txt_Dong & "', DRIVER_CLASS = '" & txt_Ho & "', DRIVER_NAME = '" & txt_Name & "', DRIVER_PHONE = '" & txt_Tel & "', START_DATE = '" & sSDate & "', END_DATE = '" & sEDate & "' WHERE SEQ = '" & nSeq & "' "
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('방문예약', '" & txt_Dong & "', '" & sLog & "', '" & txt_Ho & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            Call DataLogger(sLog)
            
            Call Clear_Field
            Call ListView1_Draw
            Call ListView1_SQL("SELECT * From tb_guestReg")
        End If
    End If
End Sub

'삭제
Private Sub SSCommand6_Click()
    If (nSeq < 0) Then
        Msg_Box.Label1 = "삭제할 차량을 선택하세요"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    
    MBox.Label3.Caption = txt_CarNo.text
    MBox.Label1.Caption = "방문예약차량 정보를 삭제합니다." & vbCrLf & " 삭제하시겠습니까?"
    MBox.Label2.Caption = "방문예약차량 삭제"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
        If (Check_InCAR(txt_CarNo) = True) Then
            Msg_Box.Label1 = "입차차량은 삭제할 수 없습니다."
            Msg_Box.Show 1
            Exit Sub
        Else
            Dim sLog As String
            'sLog = "방문예약 삭제:" & Glo_Login_ID & ":" & txt_Carno & "(" & Calendar_Bgn.value & "~" & Calendar_End.value & ")"
            sLog = "방문예약 삭제:" & Glo_Login_ID & ":" & txt_CarNo & "(" & Format(DTPicker3, "yyyy-mm-dd") & "~" & Format(DTPicker4, "yyyy-mm-dd") & ")"
            'sLog = "방문예약챠랑 삭제:" & txt_CarNo & "(" & Calendar_Bgn.value & "_" & Calendar_End.value & ")"
            adoConn.Execute "DELETE FROM tb_guestReg WHERE SEQ = '" & nSeq & "' "
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('방문예약', '" & txt_Dong & "', '" & sLog & "', '" & txt_Ho & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            Call DataLogger(sLog)
            
            Call Clear_Field
            Call ListView1_Draw
            Call ListView1_SQL("SELECT * From tb_guestReg")
        End If
    End If
End Sub

Private Function Check_InCAR(carno As String) As Boolean
    Dim bInCar  As Boolean
    Dim rs As Recordset
    
    bInCar = True
    Set rs = New ADODB.Recordset
    rs.Open "SELECT PASS_YN FROM tb_guestreg WHERE CAR_NO = '" & carno & "' ", adoConn
    Do While Not rs.EOF
        If (rs!Pass_YN = "N") Then
            bInCar = False '미입차
            Exit Do
        End If
    Loop
    Check_InCAR = bInCar
End Function



Private Sub cmd_Search_Click()
    Dim sDong, sHo As String
    Dim sQry As String
    Dim sSentence As String
    Dim sSDate As String
    Dim sEDate As String
    
    sDong = Trim(cmb_Dong.text)
    sHo = Trim(cmb_Ho.text)
    
    'sSDate = Calendar_Bgn.value & " 00:00:00"
    'sEDate = Calendar_End.value & " 23:59:59"
    sSDate = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDate = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    sQry = "SELECT * From tb_guestReg WHERE START_DATE >= '" & sSDate & "' AND END_DATE <= '" & sEDate & "' "
    
    
    If (Len(txt_SrchCarno) > 0) Then
        sQry = sQry & " AND CAR_NO LIKE '%" & txt_SrchCarno & "%' "
    End If
    
    If (cmb_Dong.text = "전체") Then
        If (cmb_Ho.text = "전체") Then
            sQry = sQry & " ORDER BY DRIVER_DEPT, DRIVER_CLASS "
        Else
            sQry = sQry & " AND DRIVER_CLASS = '" & cmb_Ho.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
        End If
    Else
        If (cmb_Ho.text = "전체") Then
            sQry = sQry & " AND DRIVER_DEPT = '" & cmb_Dong.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
        Else
            sQry = sQry & " AND DRIVER_DEPT = '" & cmb_Dong.text & "' AND DRIVER_CLASS = '" & cmb_Ho.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
        End If
    End If

    
    Call Clear_Field
    Call ListView1_Draw
    Call ListView1_SQL(sQry)
End Sub


Private Sub Set_cmbDong()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim nCount As Integer
On Error GoTo Err_P

    qry = "SELECT DRIVER_DEPT From tb_guestReg Group By DRIVER_DEPT ORDER BY DRIVER_DEPT"

    Set rs = New ADODB.Recordset
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        Call DataLogger("[FrmReg]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    cmb_Dong.Clear
    cmb_Dong.AddItem "전체"
    nCount = rs.RecordCount
    Do While Not (rs.EOF)
        cmb_Dong.AddItem "" & rs!DRIVER_DEPT
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    If (nCount > 0) Then
        cmb_Dong.ListIndex = 0
    End If

Exit Sub
Err_P:
    Call DataLogger("[FrmGuestRegCar Set_cmbDong]    " & Err.Description)
End Sub

Private Sub Set_cmbHo()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim nCount As Integer
On Error GoTo Err_P
    
    qry = "SELECT DRIVER_CLASS From tb_guestReg Group By DRIVER_CLASS ORDER BY DRIVER_CLASS"
    
    Set rs = New ADODB.Recordset
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        Call DataLogger("[FrmReg]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    cmb_Ho.Clear
    cmb_Ho.AddItem "전체"
    nCount = rs.RecordCount
    Do While Not (rs.EOF)
        cmb_Ho.AddItem "" & rs!DRIVER_CLASS
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    If (nCount > 0) Then
        cmb_Ho.ListIndex = 0
    End If
Exit Sub

Err_P:
    Call DataLogger("[FrmGuestRegCar Set_cmbHo]    " & Err.Description)
End Sub

Private Sub SSCommand7_Click()
'    Dim tmpFileName As String
'    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
'    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_방문차량 사전등록 설정" & ".xls"
'    Call MakeCSV(ListView1, tmpFileName)
    
    
    Dim tmpFileName As String
On Error GoTo Err_P
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_방문예약차량 등록내역"
        
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "엑셀파일(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave
    tmpFileName = CommonDialog1.fileName
    tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)

    Call MakeCSV(ListView1, tmpFileName)
    Exit Sub
Err_P:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
End Sub



Private Sub txt_Carno_LostFocus()
    If Not ((LenH(txt_CarNo.text) = 11) Or (LenH(txt_CarNo.text) = 12) Or (LenH(txt_CarNo.text) = 8) Or (LenH(txt_CarNo.text) = 9)) Then
        'MsgBox "차량번호 전체를 올바르게 입력하세요!!"
        'txt_Carno = ""
        'txt_Carno.SetFocus
    End If
End Sub

Private Sub txt_Dong_KeyUp(KeyCode As Integer, Shift As Integer)
    txt_Dong = Format(txt_Dong, "######")
    Debug.Print "KeyUp:" & txt_Dong
End Sub

Private Sub txt_Ho_KeyUp(KeyCode As Integer, Shift As Integer)
    txt_Ho = Format(txt_Ho, "######")
    Debug.Print "KeyUp:" & txt_Ho
End Sub


