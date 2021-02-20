VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Jung 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   14775
   ClientLeft      =   2580
   ClientTop       =   1530
   ClientWidth     =   19395
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  '단색
   BeginProperty Font 
      Name            =   "나눔고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "jung.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "jung.frx":A4D2
   ScaleHeight     =   14775
   ScaleWidth      =   19395
   Begin VB.CommandButton Lane 
      Caption         =   "Lane8  (후방)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   7
      Left            =   7305
      TabIndex        =   109
      Top             =   90
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane7  (후방)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   6
      Left            =   6465
      TabIndex        =   108
      Top             =   90
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame Frame2 
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
      Left            =   1770
      TabIndex        =   105
      ToolTipText     =   "모든 차량(정기권,미등록,미인식,출입제한 차량) 차단기 열림"
      Top             =   14130
      Width           =   5835
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "자리비움 레인2"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   3090
         TabIndex        =   107
         ToolTipText     =   "[자리비움]체크할 경우:미인식차량, 출입제한차량을 포함한 모든차량 통행을 허용힙니다."
         Top             =   210
         Width           =   2655
      End
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "자리비움 레인1"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   106
         ToolTipText     =   "[자리비움]체크할 경우:미인식차량, 출입제한차량을 포함한 모든차량 통행을 허용힙니다."
         Top             =   210
         Width           =   2655
      End
   End
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
      Height          =   555
      Left            =   750
      TabIndex        =   101
      Text            =   "25구5401"
      Top             =   90
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   3330
      TabIndex        =   100
      Top             =   90
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   4170
      TabIndex        =   99
      Top             =   90
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CheckBox Chk_Zoom 
      BackColor       =   &H00000000&
      Caption         =   " 영상 확대"
      ForeColor       =   &H0080FFFF&
      Height          =   360
      Left            =   150
      TabIndex        =   97
      Top             =   14400
      Width           =   1380
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
      Left            =   13530
      TabIndex        =   94
      ToolTipText     =   "방문차량(미등록차량) 차단기 열림"
      Top             =   14130
      Width           =   5835
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인2"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   2670
         TabIndex        =   96
         Top             =   210
         Width           =   2190
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "일반차량 레인1"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   95
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
      Left            =   7650
      TabIndex        =   91
      ToolTipText     =   "영업용차량(택배,화물) 차단기 열림"
      Top             =   14130
      Width           =   5835
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인1"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   93
         Top             =   210
         Width           =   2655
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "영업차량 레인2"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   3090
         TabIndex        =   92
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11970
      Index           =   1
      Left            =   12945
      TabIndex        =   56
      Top             =   2070
      Width           =   6345
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   12015
         Left            =   0
         Picture         =   "jung.frx":3A4F54
         ScaleHeight     =   12015
         ScaleWidth      =   6375
         TabIndex        =   57
         Top             =   -15
         Width           =   6375
         Begin Threed.SSCommand SSCommand1 
            Height          =   870
            Index           =   1
            Left            =   5160
            TabIndex        =   87
            ToolTipText     =   "차단기를 개방합니다.."
            Top             =   4230
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
            Picture         =   "jung.frx":49EF6E
         End
         Begin Threed.SSCommand cmd_GateClose 
            Height          =   870
            Index           =   1
            Left            =   210
            TabIndex        =   110
            ToolTipText     =   "차단기를 개방합니다.."
            Top             =   4230
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
            Picture         =   "jung.frx":4A0FC8
         End
         Begin VB.Image Img_outcar 
            Height          =   510
            Index           =   1
            Left            =   390
            Picture         =   "jung.frx":4A169C
            Top             =   6960
            Width           =   1815
         End
         Begin VB.Image Imgshutdown 
            Height          =   2025
            Index           =   1
            Left            =   780
            Picture         =   "jung.frx":4A4736
            Top             =   2100
            Visible         =   0   'False
            Width           =   4740
         End
         Begin VB.Label Proc_Type 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00404040&
            BackStyle       =   0  '투명
            Caption         =   "미인식입차"
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
            Left            =   915
            TabIndex        =   59
            Top             =   5250
            Width           =   4500
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
            ForeColor       =   &H00FFFF00&
            Height          =   435
            Index           =   0
            Left            =   405
            TabIndex        =   75
            Top             =   8595
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   1
            Left            =   390
            TabIndex        =   74
            Top             =   9030
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
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   2
            Left            =   405
            TabIndex        =   73
            Top             =   9465
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
            ForeColor       =   &H00FFFF00&
            Height          =   480
            Index           =   3
            Left            =   405
            TabIndex        =   72
            Top             =   9915
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
            ForeColor       =   &H00FFFF00&
            Height          =   435
            Index           =   4
            Left            =   405
            TabIndex        =   71
            Top             =   10380
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   5
            Left            =   390
            TabIndex        =   70
            Top             =   10815
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
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   6
            Left            =   405
            TabIndex        =   69
            Top             =   11265
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
            ForeColor       =   &H00FFFF00&
            Height          =   405
            Index           =   0
            Left            =   2385
            TabIndex        =   68
            Top             =   8625
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
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   1
            Left            =   2370
            TabIndex        =   67
            Top             =   9045
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   2
            Left            =   2370
            TabIndex        =   66
            Top             =   9465
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
            ForeColor       =   &H00FFFF00&
            Height          =   480
            Index           =   3
            Left            =   2370
            TabIndex        =   65
            Top             =   9915
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
            ForeColor       =   &H00FFFF00&
            Height          =   435
            Index           =   4
            Left            =   2370
            TabIndex        =   64
            Top             =   10395
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   5
            Left            =   2370
            TabIndex        =   63
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
            ForeColor       =   &H00FFFF00&
            Height          =   420
            Index           =   6
            Left            =   2370
            TabIndex        =   62
            Top             =   11265
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
            ForeColor       =   &H00000000&
            Height          =   525
            Index           =   1
            Left            =   2685
            TabIndex        =   61
            Top             =   6165
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
            ForeColor       =   &H00FF0000&
            Height          =   525
            Index           =   1
            Left            =   2685
            TabIndex        =   60
            Top             =   6975
            Width           =   3405
         End
         Begin VB.Label lbl_GN 
            Appearance      =   0  '평면
            BackColor       =   &H00800000&
            BackStyle       =   0  '투명
            Caption         =   "입구"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   480
            Index           =   1
            Left            =   255
            TabIndex        =   58
            Top             =   135
            Width           =   5655
         End
         Begin VB.Image ImageIn 
            Appearance      =   0  '평면
            BorderStyle     =   1  '단일 고정
            Height          =   4440
            Index           =   1
            Left            =   210
            Picture         =   "jung.frx":4C3B64
            Stretch         =   -1  'True
            Top             =   660
            Width           =   5880
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE9CE&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12015
      Index           =   0
      Left            =   150
      TabIndex        =   35
      Top             =   2070
      Width           =   6345
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   12060
         Left            =   0
         Picture         =   "jung.frx":4E9297
         ScaleHeight     =   12060
         ScaleWidth      =   6375
         TabIndex        =   36
         Top             =   -30
         Width           =   6375
         Begin Threed.SSCommand SSCommand1 
            Height          =   870
            Index           =   0
            Left            =   5160
            TabIndex        =   86
            ToolTipText     =   "차단기를 개방합니다.."
            Top             =   4230
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
            Picture         =   "jung.frx":5E32B1
         End
         Begin Threed.SSCommand cmd_GateClose 
            Height          =   870
            Index           =   0
            Left            =   210
            TabIndex        =   111
            ToolTipText     =   "차단기를 개방합니다.."
            Top             =   4230
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
            Picture         =   "jung.frx":5E530B
         End
         Begin VB.Image Img_outcar 
            Height          =   510
            Index           =   0
            Left            =   405
            Picture         =   "jung.frx":5E59DF
            Top             =   6975
            Width           =   1815
         End
         Begin VB.Image Imgshutdown 
            Height          =   2025
            Index           =   0
            Left            =   795
            Picture         =   "jung.frx":5E8A79
            Top             =   1980
            Visible         =   0   'False
            Width           =   4740
         End
         Begin VB.Label Proc_Type 
            Alignment       =   2  '가운데 맞춤
            BackColor       =   &H00404040&
            BackStyle       =   0  '투명
            Caption         =   "미인식 입차"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   27.75
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   960
            TabIndex        =   55
            Top             =   5235
            Width           =   4500
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
            ForeColor       =   &H00FF0000&
            Height          =   525
            Index           =   0
            Left            =   2610
            TabIndex        =   54
            Top             =   7005
            Width           =   3405
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
            ForeColor       =   &H00000000&
            Height          =   525
            Index           =   0
            Left            =   2595
            TabIndex        =   53
            Top             =   6210
            Width           =   3405
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   6
            Left            =   450
            TabIndex        =   52
            Top             =   11280
            Width           =   1800
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
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   5
            Left            =   450
            TabIndex        =   51
            Top             =   10845
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   4
            Left            =   450
            TabIndex        =   50
            Top             =   10395
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
            ForeColor       =   &H00FFFF00&
            Height          =   480
            Index           =   3
            Left            =   450
            TabIndex        =   49
            Top             =   9930
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
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   2
            Left            =   450
            TabIndex        =   48
            Top             =   9495
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
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   1
            Left            =   450
            TabIndex        =   47
            Top             =   9060
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
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   0
            Left            =   465
            TabIndex        =   46
            Top             =   8610
            Width           =   1815
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-미사용"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   6
            Left            =   2430
            TabIndex        =   45
            Top             =   11280
            Width           =   3615
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-입출상태"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   5
            Left            =   2430
            TabIndex        =   44
            Top             =   10845
            Width           =   3615
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-종료일"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   4
            Left            =   2430
            TabIndex        =   43
            Top             =   10395
            Width           =   3615
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-인식번호"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   3
            Left            =   2430
            TabIndex        =   42
            Top             =   9945
            Width           =   3615
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-호수"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   435
            Index           =   2
            Left            =   2460
            TabIndex        =   41
            Top             =   9525
            Width           =   3585
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-동"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   450
            Index           =   1
            Left            =   2430
            TabIndex        =   40
            Top             =   9060
            Width           =   3615
         End
         Begin VB.Label lbl_info_in 
            BackStyle       =   0  '투명
            Caption         =   "lbl_info_in-게이트"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   12
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   465
            Index           =   0
            Left            =   2460
            TabIndex        =   39
            Top             =   8610
            Width           =   3585
         End
         Begin VB.Label lbl_GN 
            Appearance      =   0  '평면
            BackColor       =   &H00800000&
            BackStyle       =   0  '투명
            Caption         =   "입구"
            BeginProperty Font 
               Name            =   "나눔고딕"
               Size            =   20.25
               Charset         =   129
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Index           =   0
            Left            =   300
            TabIndex        =   38
            Top             =   105
            Width           =   5655
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
            Left            =   3720
            TabIndex        =   37
            Top             =   285
            Width           =   2055
         End
         Begin VB.Image ImageIn 
            Appearance      =   0  '평면
            BorderStyle     =   1  '단일 고정
            Height          =   4440
            Index           =   0
            Left            =   210
            Picture         =   "jung.frx":607EA7
            Stretch         =   -1  'True
            Top             =   660
            Width           =   5880
         End
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4935
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
      Left            =   19860
      TabIndex        =   8
      Top             =   2715
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   19920
      TabIndex        =   7
      Top             =   2430
      Visible         =   0   'False
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
      Left            =   8100
      TabIndex        =   0
      Top             =   2355
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6675
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   11325
      Width           =   1320
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1170
      Left            =   6735
      TabIndex        =   6
      Top             =   3555
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   2064
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
      Height          =   2205
      Left            =   6675
      TabIndex        =   9
      Top             =   11715
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3889
      View            =   3
      LabelEdit       =   1
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5385
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin MSWinsockLib.Winsock Host_sock 
      Left            =   5850
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
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
      Left            =   7275
      TabIndex        =   113
      Top             =   1350
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   20250
      TabIndex        =   112
      Top             =   7980
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
      Left            =   8820
      TabIndex        =   104
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(5)"
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
      Left            =   19770
      TabIndex        =   103
      Top             =   7050
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "입출구분:"
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
      Index           =   8
      Left            =   19830
      TabIndex        =   102
      Top             =   6540
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404040&
      BackStyle       =   0  '투명
      Caption         =   "차단기 자동 열림 : "
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   98
      Top             =   14190
      Width           =   1530
   End
   Begin VB.Label LblDBInfo 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "DB오류 메시지"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12540
      TabIndex        =   88
      Top             =   60
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label LblTime 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "시간"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   14400
      TabIndex        =   90
      Top             =   405
      Width           =   4800
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   14970
      Picture         =   "jung.frx":62D5DA
      Top             =   180
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
      Left            =   20130
      TabIndex        =   89
      Top             =   5940
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   8
      Left            =   19800
      Picture         =   "jung.frx":62D970
      Top             =   5520
      Visible         =   0   'False
      Width           =   1725
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
      Left            =   5535
      TabIndex        =   85
      Top             =   1350
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label18 
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
      Left            =   11340
      TabIndex        =   84
      Top             =   2490
      Width           =   1170
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Index           =   1
      Left            =   3645
      Picture         =   "jung.frx":632C9E
      Top             =   1440
      Width           =   465
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Index           =   1
      Left            =   3645
      Picture         =   "jung.frx":633085
      Top             =   1425
      Width           =   480
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Index           =   0
      Left            =   3090
      Picture         =   "jung.frx":63346E
      Top             =   1440
      Width           =   465
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Index           =   0
      Left            =   3090
      Picture         =   "jung.frx":633855
      Top             =   1425
      Width           =   480
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
      Height          =   435
      Left            =   1425
      TabIndex        =   34
      Top             =   960
      Width           =   2685
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "카메라 상태"
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
      Height          =   315
      Index           =   0
      Left            =   1425
      TabIndex        =   33
      Top             =   1485
      Width           =   1560
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
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   6720
      TabIndex        =   32
      Top             =   10755
      Width           =   2055
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   2160
      Index           =   2
      Left            =   6720
      Picture         =   "jung.frx":633C3E
      Stretch         =   -1  'True
      Top             =   8490
      Width           =   2610
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
      Height          =   330
      Index           =   0
      Left            =   9390
      TabIndex        =   31
      Top             =   7170
      Width           =   3150
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
      Height          =   330
      Index           =   0
      Left            =   9390
      TabIndex        =   30
      Top             =   6720
      Width           =   3150
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
      Height          =   330
      Index           =   0
      Left            =   9390
      TabIndex        =   29
      Top             =   6270
      Width           =   3150
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
      Height          =   330
      Index           =   0
      Left            =   9390
      TabIndex        =   28
      Top             =   5835
      Width           =   3150
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
      Height          =   330
      Index           =   0
      Left            =   9390
      TabIndex        =   27
      Top             =   4950
      Width           =   3150
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "기        간"
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
      Left            =   6915
      TabIndex        =   26
      Top             =   7185
      Width           =   930
   End
   Begin VB.Label Label5 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "차 량 모 델"
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
      Left            =   6900
      TabIndex        =   25
      Top             =   6735
      Width           =   1080
   End
   Begin VB.Label Label4 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "연  락  처 "
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
      Left            =   6915
      TabIndex        =   24
      Top             =   6285
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "구        분"
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
      Left            =   6915
      TabIndex        =   23
      Top             =   5835
      Width           =   930
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "이        름"
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
      Left            =   6915
      TabIndex        =   22
      Top             =   5370
      Width           =   930
   End
   Begin VB.Label Lbl1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "차 량 번 호"
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
      Left            =   6915
      TabIndex        =   21
      Top             =   4950
      Width           =   1080
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
      Height          =   330
      Index           =   0
      Left            =   9390
      TabIndex        =   20
      Top             =   5355
      Width           =   3150
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
      Left            =   6720
      TabIndex        =   19
      Top             =   3210
      Width           =   5940
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "수 정 일 시"
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
      Index           =   1
      Left            =   6915
      TabIndex        =   18
      Top             =   7650
      Width           =   1080
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
      Height          =   270
      Index           =   0
      Left            =   9390
      TabIndex        =   17
      Top             =   7635
      Width           =   3150
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(0) - 출입일시"
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
      Left            =   9375
      TabIndex        =   16
      Top             =   8520
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(1)-차량번호"
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
      Left            =   9375
      TabIndex        =   15
      Top             =   8820
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(2)-이름"
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
      Left            =   9375
      TabIndex        =   14
      Top             =   9795
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(3)-게이트"
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
      Left            =   9375
      TabIndex        =   13
      Top             =   9465
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(4)-연락처"
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
      Left            =   9375
      TabIndex        =   12
      Top             =   10125
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(6)-종료일"
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
      Left            =   9375
      TabIndex        =   11
      Top             =   10425
      Width           =   3330
   End
   Begin VB.Label Lbl_inout 
      BackStyle       =   0  '투명
      Caption         =   "Lbl_inout(7)-입출상태"
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
      Left            =   9375
      TabIndex        =   10
      Top             =   9135
      Width           =   3330
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
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   0
      TabIndex        =   76
      Top             =   14100
      Width           =   19410
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
      Left            =   20190
      TabIndex        =   83
      Top             =   4830
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   10725
      TabIndex        =   82
      Top             =   1350
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
      Left            =   9000
      TabIndex        =   81
      Top             =   1350
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
      Left            =   12480
      TabIndex        =   80
      Top             =   1350
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
      Left            =   14205
      TabIndex        =   79
      Top             =   1350
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
      Left            =   15945
      TabIndex        =   78
      Top             =   1350
      Width           =   1050
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
      TabIndex        =   77
      Top             =   1350
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   6
      Left            =   17295
      Picture         =   "jung.frx":659371
      Top             =   930
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   5
      Left            =   15570
      Picture         =   "jung.frx":65E69F
      Top             =   930
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   4
      Left            =   13830
      Picture         =   "jung.frx":6639CD
      Top             =   930
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   3
      Left            =   19815
      Picture         =   "jung.frx":668CFB
      Top             =   4410
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   2
      Left            =   12105
      Picture         =   "jung.frx":66E029
      Top             =   930
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   1
      Left            =   10365
      Picture         =   "jung.frx":673357
      Top             =   930
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   0
      Left            =   8640
      Picture         =   "jung.frx":678685
      Top             =   930
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Enabled         =   0   'False
      Height          =   915
      Index           =   7
      Left            =   5175
      Picture         =   "jung.frx":67D9B3
      Top             =   930
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   9
      Left            =   19875
      Picture         =   "jung.frx":682CE1
      Top             =   7560
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   10
      Left            =   6900
      Picture         =   "jung.frx":68800F
      Top             =   930
      Visible         =   0   'False
      Width           =   1725
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
Dim FrmImg_F As Boolean



Private Sub chk_NoWork_Click(Index As Integer)
    Dim sNoWork As String
    Dim sGuestUse, sLaneName, sAutoMode As String
    Dim sOpen As String
    Dim sLog As String
    
    If (chk_NoWork(Index).value = 1) Then
        sNoWork = "자리비움"
        chk_Taxi(Index).Enabled = False
        Chk_FreePass(Index).Enabled = False
        chk_NoWork(Index).ForeColor = &HFF&
    Else
        sNoWork = "근무중"
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
    
    
    Select Case Index
    
        Case 0
                If chk_NoWork(0).value = 1 Then
                    
                    Glo_NOWORK1_YN = "Y"
                    Call Put_Ini("System Config", "NOWORK1_YN", "Y")
                    sOpen = "열림"
                    sNoWork = "자리비움"
                Else
                    Glo_NOWORK1_YN = "N"
                    Call Put_Ini("System Config", "NOWORK1_YN", "N")
                    sOpen = "해제"
                    sNoWork = "근무중"
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_NOWORK_" & Glo_NOWORK1_YN)
                    DataLogger ("FreePass Send : " & Index & "_NOWORK_" & Glo_NOWORK1_YN)
                End If
                
                sLaneName = LANE1_Name
                Glo_Lane1_NoWork = sNoWork
        Case 1
                If chk_NoWork(1).value = 1 Then
                    Glo_NOWORK2_YN = "Y"
                    Call Put_Ini("System Config", "NOWORK2_YN", "Y")
                    sOpen = "열림"
                    sNoWork = "자리비움"
                Else
                    Glo_NOWORK2_YN = "N"
                    Call Put_Ini("System Config", "NOWORK2_YN", "N")
                    sOpen = "해제"
                    sNoWork = "근무중"
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_NOWORK_" & Glo_NOWORK2_YN)
                    DataLogger ("FreePass Send : " & Index & "_NOWORK_" & Glo_NOWORK2_YN)
                End If
                
                sLaneName = LANE2_Name
                Glo_Lane2_NoWork = sNoWork
        Case 2
                If chk_NoWork(2).value = 1 Then
                    Glo_NOWORK3_YN = "Y"
                    Call Put_Ini("System Config", "NOWORK3_YN", "Y")
                    sOpen = "열림"
                    sNoWork = "자리비움"
                Else
                    Glo_NOWORK3_YN = "N"
                    Call Put_Ini("System Config", "NOWORK3_YN", "N")
                    sOpen = "해제"
                    sNoWork = "근무중"
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_NOWORK_" & Glo_NOWORK3_YN)
                    DataLogger ("FreePass Send : " & Index & "_NOWORK_" & Glo_NOWORK3_YN)
                End If
                
                sLaneName = LANE3_Name
                Glo_Lane3_NoWork = sNoWork
        Case 3
                If chk_NoWork(3).value = 1 Then
                    Glo_NOWORK4_YN = "Y"
                    Call Put_Ini("System Config", "NOWORK4_YN", "Y")
                    sOpen = "열림"
                    sNoWork = "자리비움"
                Else
                    Glo_NOWORK4_YN = "N"
                    Call Put_Ini("System Config", "NOWORK4_YN", "N")
                    sOpen = "해제"
                    sNoWork = "근무중"
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_NOWORK_" & Glo_NOWORK4_YN)
                    DataLogger ("FreePass Send : " & Index & "_NOWORK_" & Glo_NOWORK4_YN)
                End If
                
                sLaneName = LANE4_Name
                Glo_Lane4_NoWork = sNoWork
        Case 4
                If chk_NoWork(4).value = 1 Then
                    Glo_NOWORK5_YN = "Y"
                    Call Put_Ini("System Config", "NOWORK5_YN", "Y")
                    sOpen = "열림"
                    sNoWork = "자리비움"
                Else
                    Glo_NOWORK5_YN = "N"
                    Call Put_Ini("System Config", "NOWORK5_YN", "N")
                    sOpen = "해제"
                    sNoWork = "근무중"
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_NOWORK_" & Glo_NOWORK5_YN)
                    DataLogger ("FreePass Send : " & Index & "_NOWORK_" & Glo_NOWORK5_YN)
                End If
                
                sLaneName = LANE5_Name
                Glo_Lane5_NoWork = sNoWork
        Case 5
                If chk_NoWork(5).value = 1 Then
                    Glo_NOWORK6_YN = "Y"
                    Call Put_Ini("System Config", "NOWORK6_YN", "Y")
                    sOpen = "열림"
                    sNoWork = "자리비움"
                Else
                    Glo_NOWORK6_YN = "N"
                    Call Put_Ini("System Config", "NOWORK6_YN", "N")
                    sOpen = "해제"
                    sNoWork = "근무중"
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_NOWORK_" & Glo_NOWORK6_YN)
                    DataLogger ("FreePass Send : " & Index & "_NOWORK_" & Glo_NOWORK6_YN)
                End If
                
                sLaneName = LANE6_Name
                Glo_Lane6_NoWork = sNoWork
    End Select
    If (chk_NoWork(Index).value = 1) Then
        chk_NoWork(Index).ForeColor = &HFF& '빨강
    Else
        chk_NoWork(Index).ForeColor = &HFFFFFF '흰색
    End If
    
    

    '방문객 자동 처리유무
    If (chk_NoWork(Index).value = 1) Then
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
    
    If (sAutoMode = "Y") Then
        sLog = "Lane" & Index + 1 & ":" & "자리비움 설정"
    Else
        sLog = "Lane" & Index + 1 & ":" & "자리비움 해제"
    End If
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('차단기자동열림', 'HOST','" & sLog & "','자리비움'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"


End Sub
'
'Private Sub cmd_Menu_Click()
'    '방문증 출력
'    If (Guest_Error_Check = True) Then
'        Call Guest_Manual_Proc
'        'Call Guest_Proc
'        'Call Insert_Record
'    Else
'        'MsgBox "방문증 정보를 정확하게 입력하세요!"
'        Me.MousePointer = 0
'        Exit Sub
'    End If
'    Me.MousePointer = 0
'End Sub


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
Dim sGuestUse, sAutoMode As String


    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = 0

    If (Glo_ParkFull_YN = "Y") Then
        Call ParkFull_Show
    End If
    
    
    
    If (Glo_TestMode = "Y") Then
        txt_CarNo.Enabled = True
        Lane(0).Enabled = True
        Lane(1).Enabled = True
        txt_CarNo.Visible = True
        Lane(0).Visible = True
        Lane(1).Visible = True
        Lane(6).Enabled = True
        Lane(7).Enabled = True
        Lane(6).Visible = True
        Lane(7).Visible = True
    Else
        txt_CarNo.Enabled = False
        Lane(0).Enabled = False
        Lane(1).Enabled = False
        txt_CarNo.Visible = False
        Lane(0).Visible = False
        Lane(1).Visible = False
        Lane(6).Enabled = False
        Lane(7).Enabled = False
        Lane(6).Visible = False
        Lane(7).Visible = False
    End If

    Call ListView_Init1
    Call ListView_Init2

    
    
    lbl_GN(0).Caption = Trim(LANE1_Name)
    lbl_GN(1).Caption = Trim(LANE2_Name)
    
    
    Proc_Type(0).Caption = "준비중"
    Proc_Type(1).Caption = "준비중"
    
    If (LANE1_Inout = "입구") Then
        Img_outcar(0).Visible = False
    Else
        Img_outcar(0).Visible = True
    End If
    If (LANE2_Inout = "입구") Then
        Img_outcar(1).Visible = False
    Else
        Img_outcar(1).Visible = True
    End If
    
    
    For i = 0 To 6
        lbl_title_in(i).Caption = ""
        lbl_info_in(i).Caption = ""
        lbl_title_Out(i).Caption = ""
        lbl_info_Out(i).Caption = ""
    Next i
    
    lbl_carno(0).Caption = ""
    lbl_time_now(0).Caption = ""
    lbl_carno(1).Caption = ""
    lbl_time_now(1).Caption = ""
    
    For i = 0 To 8
        Lbl_inout(i).BackStyle = 0
    Next i
    
    
    ' 영업용차량 입출구 구분없애고, 레인별처리로 전환함 - 시작
    Call Chk_TaxiPassEnable(Me, LANE1_YN, Glo_TAXI1_YN, 0, LANE1_Name)
    Call Chk_TaxiPassEnable(Me, LANE2_YN, Glo_TAXI2_YN, 1, LANE2_Name)
    ' 영업용차량 입출구 구분없애고, 레인별처리로 전환함 - 끝
    
    ' 일반차량 입출구 구분없애고, 레인별처리로 전환함 - 시작
    Call Chk_NormalPassEnable(Me, LANE1_YN, Glo_FreePassLane1_YN, 0, LANE1_Name)
    Call Chk_NormalPassEnable(Me, LANE2_YN, Glo_FreePassLane2_YN, 1, LANE2_Name)
    ' 일반차량 입출구 구분없애고, 레인별처리로 전환함 - 끝
    
    chk_NoWork(0).Caption = LANE1_Name
    chk_NoWork(1).Caption = LANE2_Name
    
    
    If (Glo_Screen_No = 2) Then
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
'    Lbl_inout(5).Caption = " 인식번호 : "
    Lbl_inout(6).Caption = " 종 료 일 : "
    Lbl_inout(7).Caption = " 입출상태 : "
    Lbl_inout(8).Caption = " 입출구분 : "
    
    
    
'    If (Glo_Login_ID = "") Then
'        For i = 0 To 8
'            Lblbutton(i).Enabled = False
'            Imgbutton(i).Enabled = False
'        Next i
'        Lblbutton(1).Enabled = True
'        Imgbutton(1).Enabled = True
'        Lblbutton(6).Enabled = True
'        Imgbutton(6).Enabled = True
'    Else
'        Call frmLogin.ShowMenu(Glo_Login_ID, Glo_Login_PW)
'    End If
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

    Dim sGuestUse, sLaneName, sAutoMode As String
    
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
                FrmTcpServer.FreepassS_sock.SendData (Index & "_TAXI_" & Glo_TAXI2_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI2_YN)
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
Label18.FontBold = False

If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If
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
    
    
    If (Glo_Screen_No = 2) Then
        If (LANE1_YN = "Y" And Glo_GUEST_LANE1_YN = "Y") Then
            If (Not Glo_FrmGuest(0) Is Nothing) Then
                Call FormOnTop(Glo_FrmGuest(0).hwnd, False)
                Unload Glo_FrmGuest(0)
                Set Glo_FrmGuest(0) = Nothing
            End If
        End If
        If (LANE2_YN = "Y" And Glo_GUEST_LANE2_YN = "Y") Then
            If (Not Glo_FrmGuest(1) Is Nothing) Then
                'Call FormOnTop(Glo_FrmGuest(1).hwnd, False)
                Unload Glo_FrmGuest(1)
                Set Glo_FrmGuest(1) = Nothing
            End If
        End If
    End If
    
    
    End
Else
    Me.MousePointer = 0
    Cancel = True
End If
End Sub


Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If

End Sub


Private Sub ImageIn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'If (Index = 2) Then
'    Exit Sub
'End If
If (Chk_Zoom.value = 0) Then
    Exit Sub
End If

FrmImg.Image1.Picture = ImageIn(Index).Picture
FrmImg.Show 0
FrmImg_F = True
End Sub

Private Sub Imgbutton_Click(Index As Integer)
    Call SelectMenuButton(Me, Index)
    Exit Sub
    
'    Dim i As Integer
'
'    Call GuestForm_WindowState(vbMinimized)
'
'    Me.MousePointer = 11
'    Select Case Index
'        Case 0
'             FrmInOut.Show 1
'             'FrmInOut.Show 0
'             Me.MousePointer = 0
'             Call DataLogger("[HOST Button]    " & "입출차 보고서 화면 접근")
'        Case 2
'             FrmReg.Show 1
'             'FrmReg.Show 0
'             Me.MousePointer = 0
'             Call DataLogger("[HOST Button]    " & "정기권관리 화면 접근")
'        Case 5
'             If (Glo_Login_GUBUN = "총괄관리자") Then
'                FrmTcpServer.Show 0
'                Me.MousePointer = 0
'                Call DataLogger("[HOST Button]    " & "TCP Server 화면 접근")
'             'ElseIf (Glo_Login_GUBUN = "관리자") Then
'             Else
'                FrmTcpServer2.Show 0
'                Me.MousePointer = 0
'                Call DataLogger("[HOST Button]    " & "TCP Server2 화면 접근")
'             End If
'        Case 6
'             Call DataLogger("[HOST Button]    " & "주차관제 시스템 종료!!!")
'             Unload Me
'        Case 1
'             If (Lblbutton(1).Caption = "보호모드") Then
'                 'Call UnloadForms(Me) '모든 폼 제거(Jung, FrmTcpServer 폼은 제외)
'                 Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
'                 Call ProtectMainMenuButton(Me)
'
'                 Glo_Login_ID = ""
'                 Glo_Login_PW = ""
'                 Glo_Login_GUBUN = ""
'                 Put_Ini "System Config", "보호모드", "True"
'
'             Else
'                 Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
'                 frmLogin.Show 1
'                 'Lblbutton(1).Caption = "보호모드"
'                 ListView1.SetFocus
'             End If
'             Me.MousePointer = 0
'        Case 3
'             FrmRegHistory.Show 1
'             'FrmRegHistory.Show 0
'             Me.MousePointer = 0
'             Call DataLogger("[HOST Button]    " & "정기권 이력 화면 접근")
'        Case 4
'             FrmId.Show 1
'             'FrmId.Show 0
'             Me.MousePointer = 0
'             Call DataLogger("[HOST Button]    " & "아이디 관리 화면 접근")
'        Case 7
'            Me.MousePointer = 0
'            If (Lblbutton(Index).Caption = "무인정산기") Then
'                FrmAccnt.Show 0
'            ElseIf (Lblbutton(Index).Caption = "결제내역") Then
'                frmResult.Show 0
'            End If
'            Call DataLogger("[HOST Button]    " & "무인정산기 관리 화면 접근")
'
'        Case 8
'            Me.MousePointer = 0
'            frmResult.Show 1
'            Call DataLogger("[HOST Button]    " & "결제내역 화면 접근")
'
'        Case 9
'            Me.MousePointer = 0
'            FrmGuestLog.Show 1
'            Call DataLogger("[HOST Button]    " & "방문객내역 화면 접근")
'
'        Case 10  '방문차량 사전방문
'            Me.MousePointer = 1
'            FrmGuestRegLog.Show 0
'            Call DataLogger("[HOST Button]    " & "방문예약 화면 접근")
'            Exit Sub
'    End Select

End Sub


Private Sub Imgbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

For i = 0 To 8

    Lblbutton(i).FontBold = False

Next i
Lblbutton(Index).FontBold = True

End Sub

Private Sub Label18_Click()
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

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label18.FontBold = True

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

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If

End Sub


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If

End Sub

Private Sub SSCommand1_Click(Index As Integer)
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
'        End If
'    Else
'
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
    
Err_Proc:
    Call DataLogger(" [cmd_GateOpen_Click]  " & Err.Description)
End Sub



'' 무인기쪽으로 닫힘명령 전송
Private Sub SSCommand2_Click()
'    MBox.Label3.Caption = "차단기 닫힘!!!"
'    MBox.Label1.Caption = "차단기를 닫으시겠습니까?"
'    MBox.Label2.Caption = "Gate"
'    MBox.Show 1
'    If (Glo_MsgRet = True) Then
'        Glo_APS_Str = "GATE_CLOSE"
'        Call APS_Connect
'    End If
End Sub

Private Sub Timer1_Timer()
Dim qry As String
Dim rs As ADODB.Recordset
Dim iViewGateNo As Integer


    If (Glo_Certify = enumCertify.eCertTry And Glo_Cert_NoticeSDate < Format(Now, "yyyy-mm-dd")) Then
        LblTime(0).ForeColor = &HFF&
        LblTime(0).Caption = "[인증받으세요] " & "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")
    Else
        LblTime(0).ForeColor = &H0&
        LblTime(0).ToolTipText = ""
        LblTime(0).Caption = "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")
    End If

    
    

    If (Abs(Glo_Mon_LastInTime - Timer) >= 5) Then
        Glo_MonStat_Lane(0) = "DEAD"
        Glo_MonStat_Lane(1) = "DEAD"
    End If
'    Debug.Print Abs(Glo_Mon_LastInTime - Timer)

'''
'''    If (LANE1_YN = "Y") Then
'''        If (Glo_Mon_Lane(0) = True) Then    ' 소켓으로 LIVE 또는 DEAD 상태패킷을 받았을 경우 처리
'''            If Glo_MonStat_Lane(0) = "LIVE" Then
'''                Imgshutdown(0).Visible = False
'''                ImgGreen(0).Visible = True
'''                ImgRed(0).Visible = False
'''                Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
'''            Else
'''                Imgshutdown(0).Visible = True
'''                ImgGreen(0).Visible = False
'''                ImgRed(0).Visible = True
'''                Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
'''            End If
'''        Else
'''        ' 로컬 LPR프로그램 LIVE 또는 DEAD 상태 체크 처리
'''
'''
'''            If (Get_Process("Lane1.exe")) Then
'''                Imgshutdown(0).Visible = False
'''                ImgGreen(0).Visible = True
'''                ImgRed(0).Visible = False
'''                Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
'''            Else
'''                Imgshutdown(0).Visible = True
'''                ImgGreen(0).Visible = False
'''                ImgRed(0).Visible = True
'''                Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
'''            End If
'''        End If
'''    Else
'''        Imgshutdown(0).Visible = False
'''        ImgGreen(0).Visible = False
'''        ImgRed(0).Visible = False
'''    End If
'''
'''    If (LANE2_YN = "Y") Then
'''        If (Glo_Mon_Lane(1) = True) Then
'''            If Glo_MonStat_Lane(1) = "LIVE" Then
'''                Imgshutdown(1).Visible = False
'''                ImgGreen(1).Visible = True
'''                ImgRed(1).Visible = False
'''                Call FrmTcpServer.LPR_Alive_State_Send(1, "LIVE")
'''            Else
'''                Imgshutdown(1).Visible = True
'''                ImgGreen(1).Visible = False
'''                ImgRed(1).Visible = True
'''                Call FrmTcpServer.LPR_Alive_State_Send(1, "DEAD")
'''            End If
'''        Else
'''            If (Get_Process("Lane2.exe")) Then
'''                Imgshutdown(1).Visible = False
'''                ImgGreen(1).Visible = True
'''                ImgRed(1).Visible = False
'''                Call FrmTcpServer.LPR_Alive_State_Send(1, "LIVE")
'''                'LANE2_Handle = FindWindow(vbNullString, "Lane2")
'''            Else
'''                Imgshutdown(1).Visible = True
'''                ImgGreen(1).Visible = False
'''                ImgRed(1).Visible = True
'''                Call FrmTcpServer.LPR_Alive_State_Send(1, "DEAD")
'''            End If
'''        End If
'''    Else
'''        Imgshutdown(1).Visible = False
'''        ImgGreen(1).Visible = False
'''        ImgRed(1).Visible = False
'''    End If

    Timer1.Enabled = False
    Dim i As Integer
    Dim sViewGateName As String
    Dim sProcName() As String
    ReDim sProcName(Glo_Screen_No) As String
    Dim iLaneCount As Integer
    iLaneCount = 0
    
    
    If (LANE1_YN = "Y") Then
        sProcName(iLaneCount) = "Lane1.exe":    iLaneCount = iLaneCount + 1
    End If
    If (LANE2_YN = "Y") Then
        sProcName(iLaneCount) = "Lane2.exe":    iLaneCount = iLaneCount + 1
    End If
'    If (LANE3_YN = "Y") Then
'        sProcName(iLaneCount) = "Lane3.exe":    iLaneCount = iLaneCount + 1
'    End If
'    If (LANE4_YN = "Y") Then
'        sProcName(iLaneCount) = "Lane4.exe":    iLaneCount = iLaneCount + 1
'    End If
'    If (LANE5_YN = "Y") Then
'        sProcName(iLaneCount) = "Lane5.exe":    iLaneCount = iLaneCount + 1
'    End If
'    If (LANE6_YN = "Y") Then
'        sProcName(iLaneCount) = "Lane6.exe":    iLaneCount = iLaneCount + 1
'    End If
    
    
    iViewGateNo = Glo_GateNo - Glo_GateNo_StartNo

    'For i = 0 To Glo_Screen_No - 1
    For i = 0 To 1
        If (Glo_Mon_Lane(i) = True) Then    ' 소켓으로 LIVE 또는 DEAD 상태패킷을 받았을 경우 처리
            If Glo_MonStat_Lane(i) = "LIVE" Then
                Imgshutdown(i).Visible = False
                ImgGreen(i).Visible = True
                ImgRed(i).Visible = False
                Call FrmTcpServer.LPR_Alive_State_Send(Glo_GateNo_StartNo + i, "LIVE")
            Else
                Imgshutdown(0).Visible = True
                ImgGreen(0).Visible = False
                ImgRed(0).Visible = True
                Call FrmTcpServer.LPR_Alive_State_Send(Glo_GateNo_StartNo + i, "DEAD")
                'Call DataLogger("Lane" & Glo_GateNo_StartNo + i + 1 & " Monitor Stat : DEAD")
            End If
        Else
        ' 로컬 LPR프로그램 LIVE 또는 DEAD 상태 체크 처리
            If (Get_Process(sProcName(i))) Then
                Imgshutdown(i).Visible = False
                ImgGreen(i).Visible = True
                ImgRed(i).Visible = False
                Call FrmTcpServer.LPR_Alive_State_Send(Glo_GateNo_StartNo + i, "LIVE")
            Else
                Imgshutdown(i).Visible = True
                ImgGreen(i).Visible = False
                ImgRed(i).Visible = True
                Call FrmTcpServer.LPR_Alive_State_Send(Glo_GateNo_StartNo + i, "DEAD")
                'Call DataLogger("Lane" & Glo_GateNo_StartNo + i + 1 & " Stat : DEAD")
            End If
        End If
    Next
    
    Timer1.Enabled = True
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
    
'''    Lbl_inout(0).Caption = " 출입일시 : "
'''    Lbl_inout(1).Caption = " 차량번호 : "
'''    Lbl_inout(7).Caption = " 입출상태 : "
'''    Lbl_inout(3).Caption = " 게 이 트 : "
'''    Lbl_inout(2).Caption = " 이    름 : "
'''    Lbl_inout(4).Caption = " 연 락 처 : "
'''    Lbl_inout(6).Caption = " 종 료 일 : "
''''    Lbl_inout(5).Caption = " 인식번호 : "
'''    Lbl_inout(8).Caption = " 입출구분 : "
    
    Lbl_inout(0).Caption = " 출입일시:" & Left(ListView2.SelectedItem.text, 16)
    Lbl_inout(1).Caption = " 차량번호:" & ListView2.SelectedItem.SubItems(1)
    If ((ListView2.SelectedItem.SubItems(7) = "정상입차") Or (ListView2.SelectedItem.SubItems(7) = "정상출차")) Then
        Lbl_inout(7).ForeColor = vbWhite
    Else
        Lbl_inout(7).ForeColor = vbRed
    End If
    
    sGateName = ListView2.SelectedItem.SubItems(2)
    'Lbl_inout(3).Caption = " 게 이 트 : " & ListView2.SelectedItem.SubItems(2)
    Lbl_inout(3).Caption = " 게 이 트:" & sGateName
    Lbl_inout(2).Caption = " 이    름:" & ListView2.SelectedItem.SubItems(3)
    Lbl_inout(4).Caption = " 연 락 처:" & ListView2.SelectedItem.SubItems(4)
'    Lbl_inout(5).Caption = " 인식번호 : " & ListView2.SelectedItem.SubItems(5)
    Lbl_inout(6).Caption = " 종 료 일:" & Format(ListView2.SelectedItem.SubItems(6), "yyyy-mm-dd")
    
    
    
    
'    If (ListView2.SelectedItem.SubItems(2) = "0") Then
'        sGateName = LANE1_Name
'    ElseIf (ListView2.SelectedItem.SubItems(2) = "1") Then
'        sGateName = LANE2_Name
'    ElseIf (ListView2.SelectedItem.SubItems(2) = "2") Then
'        sGateName = LANE3_Name
'    ElseIf (ListView2.SelectedItem.SubItems(2) = "3") Then
'        sGateName = LANE4_Name
'    ElseIf (ListView2.SelectedItem.SubItems(2) = "4") Then
'        sGateName = LANE5_Name
'    ElseIf (ListView2.SelectedItem.SubItems(2) = "5") Then
'        sGateName = LANE6_Name
'    Else
'        sGateName = ""
'    End If
    
    
    
    Lbl_inout(7).Caption = " 입출상태:" & ListView2.SelectedItem.SubItems(7)
    Lbl_inout(8).Caption = " 입출구분:" & ListView2.SelectedItem.SubItems(8)
    
    'ImageIn(2).Picture = LoadPicture(ListView2.SelectedItem.SubItems(9)) '실제 이미지없을 경우 프로그램 종료현상발생(아래 코드로 대체함)
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
            Call DataLogger("[Jung]    " & "네트워크 및 DB 점검바랍니다")
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
                itmX.SubItems(5) = "" & rs!START_DATE
                itmX.SubItems(6) = "" & rs!END_DATE
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
            
            '성훈
            LblGubun(0).Caption = Format(ListView1.SelectedItem.SubItems(7), "yyyy-mm-dd hh:nn:ss")
        
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






