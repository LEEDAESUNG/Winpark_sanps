VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmG1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   14670
   ClientLeft      =   2580
   ClientTop       =   1530
   ClientWidth     =   19365
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  '단색
   Icon            =   "FrmG1.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "FrmG1.frx":A4D2
   ScaleHeight     =   14670
   ScaleWidth      =   19365
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "자리비움"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   12540
      TabIndex        =   75
      ToolTipText     =   "모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
      Top             =   12990
      Width           =   3225
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "자리비움 레인1"
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
         Index           =   0
         Left            =   270
         TabIndex        =   76
         ToolTipText     =   "[자리비움]체크할 경우:미인식차량, 출입제한차량을 포함한 모든차량 통행을 허용힙니다."
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "방문차량"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   15840
      TabIndex        =   72
      ToolTipText     =   "방문차량(미등록차량) 차단기 열림"
      Top             =   13680
      Width           =   3225
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
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   73
         Top             =   210
         Width           =   2190
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "영업차량"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   12555
      TabIndex        =   70
      ToolTipText     =   "영업용차량(택배,화물) 차단기 열림"
      Top             =   13680
      Width           =   3225
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   71
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane1"
      Height          =   555
      Index           =   0
      Left            =   3180
      TabIndex        =   68
      Top             =   150
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txt_CarNo 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   570
      TabIndex        =   67
      Text            =   "25구5401"
      Top             =   150
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   870
      Index           =   0
      Left            =   10425
      TabIndex        =   62
      ToolTipText     =   "차단기를 개방합니다.."
      Top             =   9030
      Width           =   930
      _Version        =   65536
      _ExtentX        =   1640
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "열기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmG1.frx":3B5094
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   300
      Left            =   12540
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   11265
      Width           =   1320
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
      Left            =   13755
      TabIndex        =   0
      Top             =   2685
      Width           =   2775
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
      Left            =   300
      TabIndex        =   10
      Top             =   10125
      Width           =   11640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4860
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Refresh"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   20805
      TabIndex        =   5
      Top             =   9285
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5325
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin MSWinsockLib.Winsock Host_sock 
      Left            =   5790
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   930
      Left            =   12555
      TabIndex        =   9
      Top             =   11625
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   1640
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   16771534
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
   Begin ComctlLib.ListView ListView1 
      Height          =   1020
      Left            =   12525
      TabIndex        =   12
      Top             =   3795
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1799
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
   Begin Threed.SSCommand cmd_GateClose 
      Height          =   870
      Index           =   0
      Left            =   900
      TabIndex        =   78
      ToolTipText     =   "차단기를 개방합니다.."
      Top             =   9030
      Width           =   930
      _Version        =   65536
      _ExtentX        =   1640
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "닫기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmG1.frx":3B70EE
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "방문예약"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   7290
      TabIndex        =   79
      Top             =   1545
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   8850
      TabIndex        =   74
      Top             =   240
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "입출구분 : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   19710
      TabIndex        =   69
      Top             =   5010
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '투명
      Caption         =   "[차단기 자동열림]"
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
      Height          =   255
      Left            =   12600
      TabIndex        =   66
      Top             =   12720
      Width           =   1605
   End
   Begin VB.Label LblTime 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "시간"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   15150
      TabIndex        =   65
      Top             =   450
      Width           =   4020
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
      Left            =   12510
      TabIndex        =   63
      Top             =   60
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   15030
      Picture         =   "FrmG1.frx":3B77C2
      Top             =   210
      Width           =   4395
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "결제내역"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   20010
      TabIndex        =   64
      Top             =   4305
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "무인정산기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   5565
      TabIndex        =   61
      Top             =   1545
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Imgshutdown 
      Height          =   2325
      Left            =   2850
      Picture         =   "FrmG1.frx":3B7B58
      Top             =   5100
      Visible         =   0   'False
      Width           =   6840
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Left            =   3150
      Picture         =   "FrmG1.frx":3EB7E2
      Top             =   1635
      Width           =   480
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Left            =   3150
      Picture         =   "FrmG1.frx":3EBBCB
      Top             =   1635
      Width           =   465
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   17670
      TabIndex        =   60
      Top             =   2790
      Width           =   1170
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "시스템종료"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   17670
      TabIndex        =   59
      Top             =   1545
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "환경설정"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   15960
      TabIndex        =   58
      Top             =   1545
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "근무자관리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   14220
      TabIndex        =   57
      Top             =   1545
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "정기권관리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   12495
      TabIndex        =   55
      Top             =   1545
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "입출차조회"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9015
      TabIndex        =   53
      Top             =   1545
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   0
      Left            =   8655
      Picture         =   "FrmG1.frx":3EBFB2
      Top             =   1125
      Width           =   1725
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      Height          =   7305
      Left            =   885
      Top             =   2610
      Width           =   10485
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   7470
      Index           =   0
      Left            =   810
      Picture         =   "FrmG1.frx":3F12E0
      Stretch         =   -1  'True
      Top             =   2535
      Width           =   10650
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   15510
      TabIndex        =   51
      Top             =   9480
      Width           =   3630
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   15510
      TabIndex        =   50
      Top             =   10785
      Width           =   3630
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   20250
      TabIndex        =   49
      Top             =   9840
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   15510
      TabIndex        =   48
      Top             =   10485
      Width           =   3630
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   15510
      TabIndex        =   47
      Top             =   9810
      Width           =   3630
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   15510
      TabIndex        =   46
      Top             =   10140
      Width           =   3630
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   15510
      TabIndex        =   45
      Top             =   9150
      Width           =   3630
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   15510
      TabIndex        =   44
      Top             =   8805
      Width           =   3630
   End
   Begin VB.Label LblGubun 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   43
      Top             =   7995
      Width           =   3405
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "수 정 일 시"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   1
      Left            =   12780
      TabIndex        =   42
      Top             =   7995
      Width           =   2190
   End
   Begin VB.Label Proc_Type 
      BackColor       =   &H00404040&
      BackStyle       =   0  '투명
      Caption         =   "준비중"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   6615
      TabIndex        =   41
      Top             =   11550
      Width           =   4245
   End
   Begin VB.Label lbl_time_now 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   0
      Left            =   6630
      TabIndex        =   40
      Top             =   13830
      Width           =   4290
   End
   Begin VB.Label lbl_carno 
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
      Caption         =   "경기00가0000"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   39
      Top             =   12705
      Width           =   4155
   End
   Begin VB.Label lbl_title_in 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   5
      Left            =   450
      TabIndex        =   38
      Top             =   13845
      Width           =   1785
   End
   Begin VB.Label lbl_title_in 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   450
      TabIndex        =   37
      Top             =   13395
      Width           =   1785
   End
   Begin VB.Label lbl_title_in 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   450
      TabIndex        =   36
      Top             =   12960
      Width           =   1785
   End
   Begin VB.Label lbl_title_in 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   450
      TabIndex        =   35
      Top             =   12495
      Width           =   1785
   End
   Begin VB.Label lbl_title_in 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   450
      TabIndex        =   34
      Top             =   12030
      Width           =   1785
   End
   Begin VB.Label lbl_title_in 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   450
      TabIndex        =   33
      Top             =   11595
      Width           =   1785
   End
   Begin VB.Label lbl_info_in 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   5
      Left            =   2955
      TabIndex        =   32
      Top             =   13845
      Width           =   3270
   End
   Begin VB.Label lbl_info_in 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   2955
      TabIndex        =   31
      Top             =   13395
      Width           =   3270
   End
   Begin VB.Label lbl_info_in 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   2955
      TabIndex        =   30
      Top             =   12960
      Width           =   3270
   End
   Begin VB.Label lbl_info_in 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   2955
      TabIndex        =   29
      Top             =   12495
      Width           =   3270
   End
   Begin VB.Label lbl_info_in 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   1
      Left            =   2955
      TabIndex        =   28
      Top             =   12030
      Width           =   3270
   End
   Begin VB.Label lbl_info_in 
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   2955
      TabIndex        =   27
      Top             =   11595
      Width           =   3270
   End
   Begin VB.Label LblSearch 
      BackColor       =   &H00000000&
      Caption         =   "등록차량 검색결과 : "
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
      Left            =   12525
      TabIndex        =   26
      Top             =   3405
      Width           =   6435
   End
   Begin VB.Label LblName 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   25
      Top             =   5715
      Width           =   3405
   End
   Begin VB.Label Lbl1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "차 량 번 호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   12780
      TabIndex        =   24
      Top             =   5295
      Width           =   2190
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "이       름"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   12780
      TabIndex        =   23
      Top             =   5730
      Width           =   2130
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "구       분"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   12780
      TabIndex        =   22
      Top             =   6210
      Width           =   2130
   End
   Begin VB.Label Label4 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "연  락   처 "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   12780
      TabIndex        =   21
      Top             =   6645
      Width           =   2130
   End
   Begin VB.Label Label5 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "차 량 모 델"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   12780
      TabIndex        =   20
      Top             =   7110
      Width           =   2190
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "기       간"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Index           =   0
      Left            =   12780
      TabIndex        =   19
      Top             =   7545
      Width           =   2130
   End
   Begin VB.Label LblCar 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   18
      Top             =   5280
      Width           =   3405
   End
   Begin VB.Label LblId 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   17
      Top             =   6180
      Width           =   3405
   End
   Begin VB.Label LblCarType 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   16
      Top             =   6630
      Width           =   3405
   End
   Begin VB.Label LblTel 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   15
      Top             =   7095
      Width           =   3405
   End
   Begin VB.Label LblDate 
      Appearance      =   0  '평면
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
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
      Height          =   285
      Index           =   0
      Left            =   15555
      TabIndex        =   14
      Top             =   7515
      Width           =   3405
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   2220
      Index           =   2
      Left            =   12555
      Picture         =   "FrmG1.frx":3F7C14
      Stretch         =   -1  'True
      Top             =   8820
      Width           =   2955
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
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "주차관제 시스템"
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
      Height          =   420
      Left            =   1635
      TabIndex        =   8
      Top             =   1095
      Width           =   4110
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "카메라 상태 "
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
      Height          =   330
      Left            =   1635
      TabIndex        =   7
      Top             =   1710
      Width           =   1590
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
      Left            =   20130
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   5175
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
   Begin VB.Label Label7 
      BackColor       =   &H00404040&
      Height          =   7605
      Left            =   315
      TabIndex        =   52
      Top             =   2460
      Width           =   11655
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "보호해제"
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
      Height          =   255
      Index           =   1
      Left            =   10755
      TabIndex        =   54
      Top             =   1545
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   1
      Left            =   10395
      Picture         =   "FrmG1.frx":41D347
      Top             =   1125
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   2
      Left            =   12120
      Picture         =   "FrmG1.frx":422675
      Top             =   1125
      Width           =   1725
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "정기권이력"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   20040
      TabIndex        =   56
      Top             =   3045
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   3
      Left            =   19665
      Picture         =   "FrmG1.frx":4279A3
      Top             =   2625
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   4
      Left            =   13845
      Picture         =   "FrmG1.frx":42CCD1
      Top             =   1125
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   5
      Left            =   15585
      Picture         =   "FrmG1.frx":431FFF
      Top             =   1125
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   6
      Left            =   17295
      Picture         =   "FrmG1.frx":43732D
      Top             =   1125
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   7
      Left            =   5205
      Picture         =   "FrmG1.frx":43C65B
      Top             =   1125
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   8
      Left            =   19680
      Picture         =   "FrmG1.frx":441989
      Top             =   3885
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "방문객관리"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   20055
      TabIndex        =   77
      Top             =   1995
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   9
      Left            =   19680
      Picture         =   "FrmG1.frx":446CB7
      Top             =   1575
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   10
      Left            =   6930
      Picture         =   "FrmG1.frx":44BFE5
      Top             =   1125
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "FrmG1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyText(1 To 8) As New clsText
Dim DataField_Enabled As Boolean
Dim Save_TagNum As String


Private Sub chk_NoWork_Click(Index As Integer)

    Dim sNoWork As String
    Dim sSendValue As String
    Dim sLaneName, sGuestUse, sAutoMode As String
    

    If (chk_NoWork(Index).value = 1) Then
        sNoWork = "자리비움"
        sSendValue = "Y"
        chk_Taxi(Index).Enabled = False
        Chk_FreePass(Index).Enabled = False
        chk_NoWork(Index).ForeColor = &HFF&
    Else
        sNoWork = "근무중"
        sSendValue = "N"
        chk_Taxi(Index).Enabled = True
        Chk_FreePass(Index).Enabled = True
        chk_NoWork(Index).ForeColor = &HFFFFFF
    End If

    Select Case Index
        Case 0
            Glo_Lane1_NoWork = sNoWork
            sLaneName = LANE1_Name
        Case 1
            Glo_Lane2_NoWork = sNoWork
            sLaneName = LANE2_Name
        Case 2
            Glo_Lane3_NoWork = sNoWork
            sLaneName = LANE3_Name
        Case 3
            Glo_Lane4_NoWork = sNoWork
            sLaneName = LANE4_Name
        Case 4
            Glo_Lane5_NoWork = sNoWork
            sLaneName = LANE5_Name
        Case 5
            Glo_Lane6_NoWork = sNoWork
            sLaneName = LANE6_Name
    End Select
    
    
    '방문객 자동 처리유무
    If (chk_NoWork(Index).value = 1 Or Chk_FreePass(Index).value = 1) Then
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
    sLog = "차단기 자동열림[자리비움] Lane:" & Index + 1 & ":" & sNoWork
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('차단기자동열림', 'HOST','" & sLog & "','자리비움'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"

End Sub


Private Sub cmd_GateClose_Click(Index As Integer)
    On Error GoTo Err_P
    
    Call DataLogger("[GATE DOWN BTN]  Target Gate = Lane" & Index + 1)
    Call Relay_Close(0, Index)
    
    Exit Sub
    
Err_P:
    Call DataLogger("[GATE DOWN BTN]  Target Gate = Lane" & Index + 1 & ", " & Err.Description)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim SQL As String
    Dim Reg_Addr As String
    Dim sGuestUse, sAutoMode   As String

    IniFileName$ = App.Path & "\Winpark.ini"
    Report_Path_Name$ = App.Path & "\Data\"
    Doc_Path_Name$ = App.Path & "\Doc\"
    
    If App.PrevInstance = True Then
        End
    End If


    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = 0
    
    
    If (Glo_ParkFull_YN = "Y") Then
        Call ParkFull_Show
    End If
    
    
    
    If (Glo_TestMode = "Y") Then
        txt_CarNo.Enabled = True
        Lane(0).Enabled = True
        txt_CarNo.Visible = True
        Lane(0).Visible = True
    Else
        txt_CarNo.Enabled = False
        Lane(0).Enabled = False
        txt_CarNo.Visible = False
        Lane(0).Visible = False
    End If

    
    Call ListView_Init1
    Call ListView_Init2

    
    ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
    
    For i = 0 To 5
        lbl_title_in(i).Caption = ""
        lbl_info_in(i).Caption = ""
    Next i
    
    lbl_carno(0).Caption = ""
    lbl_time_now(0).Caption = ""
    
    For i = 0 To 8
        Lbl_inout(i).BackStyle = 0
    Next i
    

    ' 영업용차량 입출구 구분없애고, 레인별처리로 전환함 - 시작
    Call Chk_TaxiPassEnable(Me, LANE1_YN, Glo_TAXI1_YN, 0, LANE1_Name)
    ' 영업용차량 입출구 구분없애고, 레인별처리로 전환함 - 끝
        
    ' 일반차량 입출구 구분없애고, 레인별처리로 전환함 - 시작
    Call Chk_NormalPassEnable(Me, LANE1_YN, Glo_FreePassLane1_YN, 0, LANE1_Name)
    ' 일반차량 입출구 구분없애고, 레인별처리로 전환함 - 끝
    
    chk_NoWork(0).Caption = LANE1_Name
    
    
    
    If (Glo_Screen_No = 1) Then
        
        '방문차량 입력 폼 생성
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

            '방문객 자동 처리유무
            If (Glo_FreePassLane1_YN = "Y") Then
                sGuestUse = "(자동처리)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            If (Not Glo_FrmGuest(0) Is Nothing) Then '만들어져 있다면
                'Call Glo_FrmGuest(0).SetGuestName(LANE1_Name & sGuestUse)
                Call Glo_FrmGuest(0).SetAutoMode(sAutoMode, LANE1_Name & sGuestUse)
            End If


        End If
    End If


    If (Glo_User_Type = "구분1/구분2") Then
        Label5(0).Caption = "소 속, 직 급"
    Else
        Label5(0).Caption = "  동 / 호 수"
    End If
    
    
    Lbl_inout(0).Caption = " 출입일시 : "
    Lbl_inout(1).Caption = " 차량번호 : "
    Lbl_inout(2).Caption = " 이    름 : "
    Lbl_inout(3).Caption = " 게 이 트 : "
    Lbl_inout(4).Caption = " 연 락 처 : "
    Lbl_inout(5).Caption = " 인식번호 : "
    Lbl_inout(6).Caption = " 종 료 일 : "
    Lbl_inout(7).Caption = " 입출상태 : "
    Lbl_inout(8).Caption = " 입출구분 : "
       
    
    
'''    If (Glo_Login_ID = "") Then
'''        For i = 0 To 8
'''            Lblbutton(i).Enabled = False
'''            Imgbutton(i).Enabled = False
'''        Next i
'''        Lblbutton(1).Enabled = True
'''        Imgbutton(1).Enabled = True
'''        Lblbutton(6).Enabled = True
'''        Imgbutton(6).Enabled = True
'''    Else
'''        Call frmLogin.ShowMenu(Glo_Login_ID, Glo_Login_PW)
'''    End If

    Call ProtectMainMenuButton(Me)
    
    Call ShowTitlebarSiteCode
    
Timer1.Enabled = True
FrmTcpServer.Hide
gHW = Me.hwnd
Call Hook

End Sub

Public Sub ReDraw(sKind As String, iIndex As Integer, iValue As Integer)
    If sKind = "FreePass" Then
        Chk_FreePass(iIndex).value = iValue
        Call Chk_FreePass_Click(iIndex)
    ElseIf sKind = "Taxi" Then
        chk_Taxi(iIndex).value = iValue
        Call chk_Taxi_Click(iIndex)
    ElseIf sKind = "NOWORK" Then
        chk_NoWork(iIndex).value = iValue
        Call chk_NoWork_Click(iIndex)
    End If
End Sub
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
                'FrmTcpServer.FreepassS_sock.SendData ((Index + Glo_GateNo_StartNo) & "_FREEPASS_" & Glo_FreePassLane1_YN)
                'DataLogger ("Taxi Send : " & (Index + Glo_GateNo_StartNo) & "_FREEPASS_" & Glo_FreePassLane1_YN)
                FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane1_YN)
                DataLogger ("Taxi Send : " & Index & "_FREEPASS_" & Glo_FreePassLane1_YN)
            End If

            sLaneName = LANE1_Name
            
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

Public Sub Chk_FreePassEnable(ByVal Index As Integer, ByVal bVal As Boolean)
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
                DataLogger ("FreePass Send : " & Index & "_TAXI_" & Glo_TAXI1_YN)
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



Private Sub Command2_Click()

ListView2.ListItems.Clear
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


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontBold = False
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
    
    Call Unhook
    
    If (Glo_Screen_No = 1) Then
        '방문차량
        If (Glo_GUEST_LANE1_YN = "Y") Then
            If (Not Glo_FrmGuest(0) Is Nothing) Then
                'Call FormOnTop(Glo_FrmGuest(0).hwnd, False)
                Unload Glo_FrmGuest(0)
                Set Glo_FrmGuest(0) = Nothing
            End If
        End If
    End If
    
    End
End If
Me.MousePointer = 0
Cancel = True
End Sub



Private Sub Imgbutton_Click(Index As Integer)

    Call SelectMenuButton(Me, Index)
    Exit Sub
'
'Dim i As Integer
'
'Call GuestForm_WindowState(vbMinimized)
'
'Me.MousePointer = 11
'Select Case Index
'    Case 0
'         'FrmInOut.Show 1
'         FrmInOut.Show 0
'         Me.MousePointer = 0
'         Call DataLogger("[HOST Button]    " & "입출차 보고서 화면 접근")
'    Case 2
'         'FrmReg.Show 1
'         FrmReg.Show 0
'         Me.MousePointer = 0
'         Call DataLogger("[HOST Button]    " & "정기권관리 화면 접근")
'    Case 5
'         If (Glo_Login_GUBUN = "총괄관리자") Then
'            FrmTcpServer.Show 0
'            Me.MousePointer = 0
'            Call DataLogger("[HOST Button]    " & "TCP Server 화면 접근")
'         'ElseIf (Glo_Login_GUBUN = "관리자") Then
'         Else
'            FrmTcpServer2.Show 0
'            Me.MousePointer = 0
'            Call DataLogger("[HOST Button]    " & "TCP Server2 화면 접근")
'         End If
'    Case 6
'         Call DataLogger("[HOST Button]    " & "주차관제 시스템 종료!!!")
'         Unload Me
'    Case 1
''''         If (Lblbutton(1).Caption = "보호모드") Then
''''            Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
''''            Lblbutton(1).Caption = "보호해제"
''''            For i = 0 To 8
''''                Lblbutton(i).Enabled = False
''''                Imgbutton(i).Enabled = False
''''            Next i
''''            Lblbutton(6).Enabled = True '시스템종료
''''            Lblbutton(1).Enabled = True '보호해제
''''            Imgbutton(6).Enabled = True
''''            Imgbutton(1).Enabled = True
''''
''''            Lblbutton(7).Visible = False '무인정산기
''''            Imgbutton(7).Visible = False
''''
''''            Lblbutton(10).Visible = False '방문예약
''''            Imgbutton(10).Visible = False
''''            Lblbutton(10).Enabled = False '방문예약
''''            Imgbutton(10).Enabled = False
''''
''''            Put_Ini "System Config", "보호모드", "True"
''''
''''            Glo_Login_ID = ""
''''            Glo_Login_PW = ""
''''            Glo_Login_GUBUN = ""
''''         Else
''''            frmLogin.Show 1
''''            Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
''''            Lblbutton(1).Caption = "보호모드"
''''            ListView1.SetFocus
''''         End If
''''         Me.MousePointer = 0
'
'         If (Lblbutton(1).Caption = "보호모드") Then
'            'Call UnloadForms(Me) '모든 폼 제거(Jung, FrmTcpServer 폼은 제외)
'            Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
'            Call ProtectMainMenuButton(Me)
'
'            Glo_Login_ID = ""
'            Glo_Login_PW = ""
'            Glo_Login_GUBUN = ""
'            Put_Ini "System Config", "보호모드", "True"
'
'         Else
'            Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
'            frmLogin.Show 1
'            'Lblbutton(1).Caption = "보호모드"
'            ListView1.SetFocus
'         End If
'         Me.MousePointer = 0
'    Case 3
'         'FrmRegHistory.Show 1
'         FrmRegHistory.Show 0
'         Me.MousePointer = 0
'         Call DataLogger("[HOST Button]    " & "정기권 이력 화면 접근")
'    Case 4
'            'FrmId.Show 1
'            FrmId.Show 0
'            Me.MousePointer = 0
'            Call DataLogger("[HOST Button]    " & "아이디 관리 화면 접근")
'    Case 7
'        Me.MousePointer = 0
'        If (Lblbutton(Index).Caption = "무인정산기") Then
'            FrmAccnt.Show 0
'        ElseIf (Lblbutton(Index).Caption = "결제내역") Then
'            frmResult.Show 1
'        End If
'        Call DataLogger("[HOST Button]    " & "무인정산기 관리 화면 접근")
'    Case 8
'        Me.MousePointer = 1
'        frmResult.Show 0
'        Call DataLogger("[HOST Button]    " & "결제내역 화면 접근")
'    Case 9
'        Me.MousePointer = 1
'        'FrmGuestLog.Show 1
'        FrmGuestLog.Show 0
'        Call DataLogger("[HOST Button]    " & "방문객내역 화면 접근")
'
'    Case 10  '방문차량 사전방문
'        Me.MousePointer = 1
'        'FrmGuestRegLog.Show 1
'        FrmGuestRegLog.Show 0
'        Call DataLogger("[HOST Button]    " & "방문예약 화면 접근")
'        Exit Sub
'
'End Select

End Sub

Private Sub Imgbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

For i = 0 To 8
    Lblbutton(i).FontBold = False

Next i
Lblbutton(Index).FontBold = True

End Sub

Private Sub Label8_Click()
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

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label8.FontBold = True

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


Private Sub Lblbutton_Click(Index As Integer)
    Call GuestForm_WindowState(vbMinimized)
    Call Imgbutton_Click(Index)
End Sub

Private Sub Lblbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

For i = 0 To 8

    Lblbutton(i).FontBold = False

Next i
Lblbutton(Index).FontBold = True

End Sub

Private Sub SSCommand1_Click(Index As Integer)
On Error GoTo Err_Proc

    
    On Error GoTo Err_Proc

'    Select Case HostType
'        Case 0
'            If LANE1_YN = "Y" Then
'                Select Case Index
'                    Case 0
'                        Call DataLogger("[GATE OPEN BTN]  Target Gate = 0")
'                        Call Relay_Out(0, 0)
'                    Case 1
'                        Call DataLogger("[GATE OPEN BTN]  Target Gate = 1")
'                        Call Relay_Out(0, 1)
'                End Select
'            Else
'                Select Case Index
'                    Case 0
'                        Call DataLogger("[GATE OPEN BTN]  Target Gate = 2")
'                        Call Relay_Out(0, 0)
'                    Case 1
'                        Call DataLogger("[GATE OPEN BTN]  Target Gate = 3")
'                        Call Relay_Out(0, 1)
'                End Select
'            End If
'
'        Case 1
'                Select Case Index
'                    Case 0
'                        Call DataLogger("[GATE OPEN BTN]  Target Gate = 0")
'                        Call Relay_Out(0, 0)
'                    Case 1
'                        Call DataLogger("[GATE OPEN BTN]  Target Gate = 2")
'                        Call Relay_Out(0, 2)
'                End Select
'    End Select

'    Debug.Print "[GATE OPEN BTN]  Target Gate = Lane" & Index + 1
'    If (Glo_ParkFull_YN = "Y") Then
'        If (Glo_ParkFull_Count > Glo_ParkNow_Count) Then
'            Glo_ParkNow_Count = Glo_ParkNow_Count + 1
'
'            Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1)
'            Call Relay_Out(0, Index)
'
'        Else
'            Glo_ParkNow_Count = Glo_ParkFull_Count
'            Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1 & "만차:차단기 안열림")
'
'        End If
'    Else
'        Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1)
'        Call Relay_Out(0, Index)
'    End If

    Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1)
    Call Relay_Out(0, Index)


    If (Glo_ParkFull_YN = "Y") Then
        Dim sInOut As String
        Select Case Index
            Case 0
                sInOut = LANE1_Inout
            Case 1
                sInOut = LANE2_Inout
            Case 2
                sInOut = LANE3_Inout
            Case 3
                sInOut = LANE4_Inout
            Case 4
                sInOut = LANE5_Inout
            Case 5
                sInOut = LANE6_Inout
        End Select
        
        Call ParkFull_GetState(Index, sInOut)   '만차계산
        Call ParkFull_PutNMLDisplay(Index)  '전광판출력
        Call ParkFull_Show                  '화면출력
    End If
    
    Exit Sub
    
Err_Proc:
    Call DataLogger(" [cmd_GateOpen_Click]  " & Err.Description)
End Sub

Private Sub Timer1_Timer()
Dim qry As String
Dim rs As ADODB.Recordset

    If (Glo_Certify = enumCertify.eCertTry And Glo_Cert_NoticeSDate < Format(Now, "yyyy-mm-dd")) Then
        LblTime(0).ForeColor = &HFF&
        LblTime(0).Caption = "[인증받으세요] " & "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")
    Else
        LblTime(0).ForeColor = &H0&
        LblTime(0).ToolTipText = ""
        LblTime(0).Caption = "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")
    End If

'If (Format(Now, "NNSS") = "0001") Then
'    '게이트 카운트 초기화
'    Qry = "show tables"
'    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
'    Set rs = Nothing
'    adoConn.Execute ""
'    List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "    MySQL Connection Test...!! ", 0
'End If

    If (Abs(Glo_Mon_LastInTime - Timer) >= 5) Then
        Glo_MonStat_Lane(0) = "DEAD"
    End If
    
    If (LANE1_YN = "Y") Then
        If (Glo_Mon_Lane(0) = True) Then
                If Glo_MonStat_Lane(0) = "LIVE" Then
                    ImgGreen.Visible = True
                    ImgRed.Visible = False
                    Imgshutdown.Visible = False
                    Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
                Else
                    ImgGreen.Visible = False
                    ImgRed.Visible = True
                    Imgshutdown.Visible = True
                    Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
                    'Call DataLogger("Lane1 Monitor Stat : DEAD")
                End If
        Else
                If (Get_Process("Lane1.exe")) Then
                    ImgGreen.Visible = True
                    ImgRed.Visible = False
                    Imgshutdown.Visible = False
                    Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
                Else
                    ImgGreen.Visible = False
                    ImgRed.Visible = True
                    Imgshutdown.Visible = True
                    Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
                    'Call DataLogger("Lane1 Stat : DEAD")
                End If
        End If
    Else
        Imgshutdown.Visible = False
        ImgGreen.Visible = False
        ImgRed.Visible = False
    End If




End Sub

Private Sub ImageIn_DblClick(Index As Integer)
If (Index = 2) Then
    If (ImageIn(2).height = 3780) Then
        ImageIn(2).height = 2220
        ImageIn(2).width = 2955
    Else
        ImageIn(2).height = 3780
        ImageIn(2).width = 6375
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

LblCar(0).Caption = ListView1.SelectedItem.text
LblName(0).Caption = ListView1.SelectedItem.SubItems(1)
LblId(0).Caption = ListView1.SelectedItem.SubItems(2)
LblCarType(0).Caption = ListView1.SelectedItem.SubItems(3)
LblTel(0).Caption = ListView1.SelectedItem.SubItems(4)

If (ListView1.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView1.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
    LblDate(0).ForeColor = vbBlack
    LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6)
Else
    LblDate(0).ForeColor = vbRed
    LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6) & "   " & "[기간에러]"
End If

LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)

End Sub


Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
    Dim sGateName As String
    ListView2.SetFocus
    
    '''Lbl_inout(0).Caption = " 출입일시 : "
    '''Lbl_inout(1).Caption = " 차량번호 : "
    '''Lbl_inout(2).Caption = " 이    름 : "
    '''Lbl_inout(3).Caption = " GATE : "
    '''Lbl_inout(4).Caption = " 연 락 처 : "
    '''Lbl_inout(5).Caption = " 인식번호 : "
    '''Lbl_inout(6).Caption = " 종 료 일 : "
    '''Lbl_inout(7).Caption = " 입출상태 : "
    '''Lbl_inout(8).Caption = " 입출구분 : "
    
    Lbl_inout(0).Caption = " 출입일시:" & ListView2.SelectedItem.text
    Lbl_inout(1).Caption = " 차량번호:" & ListView2.SelectedItem.SubItems(1)
    Lbl_inout(2).Caption = " 이    름:" & ListView2.SelectedItem.SubItems(3)
    Lbl_inout(3).Caption = " 게 이 트:" & ListView2.SelectedItem.SubItems(2)
'    If (ListView2.SelectedItem.SubItems(2) = "0") Then
'            sGateName = LANE1_Name
'        ElseIf (ListView2.SelectedItem.SubItems(2) = "1") Then
'            sGateName = LANE2_Name
'        ElseIf (ListView2.SelectedItem.SubItems(2) = "2") Then
'            sGateName = LANE3_Name
'        ElseIf (ListView2.SelectedItem.SubItems(2) = "3") Then
'            sGateName = LANE4_Name
'        ElseIf (ListView2.SelectedItem.SubItems(2) = "4") Then
'            sGateName = LANE5_Name
'        ElseIf (ListView2.SelectedItem.SubItems(2) = "5") Then
'            sGateName = LANE6_Name
'        Else
'            sGateName = ""
'        End If
    'Lbl_inout(3).Caption = " 게 이 트 : " & sGateName
    Lbl_inout(4).Caption = " 연 락 처:" & ListView2.SelectedItem.SubItems(4)
    Lbl_inout(5).Caption = " 인식번호:" & ListView2.SelectedItem.SubItems(5)
    Lbl_inout(6).Caption = " 종 료 일:" & ListView2.SelectedItem.SubItems(6)
    
    If ((ListView2.SelectedItem.SubItems(7) = "정상입차") Or (ListView2.SelectedItem.SubItems(7) = "정상출차")) Then
        Lbl_inout(7).ForeColor = vbWhite
    Else
        Lbl_inout(7).ForeColor = vbRed
    End If
    Lbl_inout(7).Caption = " 입출상태:" & ListView2.SelectedItem.SubItems(7)
    
'    Lbl_inout(8).Caption = " 입출구분 : " & ListView2.SelectedItem.SubItems(8)
    
    'ImageIn(2).Picture = LoadPicture(ListView2.SelectedItem.SubItems(9))
    If (IsFile(ListView2.SelectedItem.SubItems(8)) = True) Then
        ImageIn(2).Picture = LoadPicture(ListView2.SelectedItem.SubItems(8))
    Else
        ImageIn(2).Picture = LoadPicture(App.Path & "\NoCar.jpg")
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim Car_Num_Str As String
Dim qry As String
Dim rs As Recordset
Dim rs_Part As Recordset
Dim itmX As ListItem
Dim bQryResult As Boolean

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
        qry = "Select * From tb_reg WHERE CAR_NO Like '%" & Text1 & "' ORDER BY CAR_NO"
        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            Call DataLogger("[KeyPress]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        
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
                'itmX.SubItems(4) = "" & rs!CAR_MODEL
                itmX.SubItems(4) = "" & rs!DRIVER_DEPT & " / " & rs!DRIVER_CLASS
                itmX.SubItems(5) = "" & Left(rs!START_DATE, 10)
                itmX.SubItems(6) = "" & Left(rs!END_DATE, 10)
                itmX.SubItems(7) = "" & Format(rs!REG_DATE, "yyyy-mm-dd hh:nn:ss")
                rs.MoveNext
            Loop
            
            ListView1.ListItems.Item(1).Selected = True
            
            If (rs.RecordCount = 1) Then
            
            Else
                ListView1.SetFocus
            End If
            
            LblCar(0).Caption = ListView1.SelectedItem.text
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
            LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)
        
        End If
        
        Set rs = Nothing
        KeyAscii = 0
        Exit Sub
End If

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
    
    'ListView1.ColumnHeaders.Add , , " 차량모델   "
    If (Glo_User_Type = "구분1/구분2") Then
        ListView1.ColumnHeaders.Add , , " 소속, 직급   "
    Else
        ListView1.ColumnHeaders.Add , , " 동, 호수   "
    End If
    
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
    ListView2.ColumnHeaders.Add , , " 처리일시                             "         '7
    
    ListView2.ColumnHeaders.Add , , " 차량번호           "      '0
    ListView2.ColumnHeaders.Add , , " 구    분         "  '1
    ListView2.ColumnHeaders.Add , , " 이    름  "       '2
    ListView2.ColumnHeaders.Add , , " 전화번호     "  '3
    ListView2.ColumnHeaders.Add , , " 인식번호     "   '4
    ListView2.ColumnHeaders.Add , , " 종 료 일     "        '5
    ListView2.ColumnHeaders.Add , , " 인식상태     "          '6
    
    'ListView2.ColumnHeaders.Add , , " 입출구분     "    '8
    ListView2.ColumnHeaders.Add , , " 이미지명"    '9
    
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
Host_sock.SendData (strData)
End Sub



