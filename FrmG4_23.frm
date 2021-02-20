VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{580DB626-DAF8-439F-8781-62D158794504}#1.0#0"; "xIP.ocx"
Begin VB.Form FrmG4_23 
   Caption         =   " HOST Program"
   ClientHeight    =   15630
   ClientLeft      =   -2130
   ClientTop       =   1575
   ClientWidth     =   28770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmG4_23.frx":0000
   ScaleHeight     =   15630
   ScaleWidth      =   28770
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
      Height          =   645
      IMEMode         =   10  '한글 
      Left            =   15825
      TabIndex        =   54
      Top             =   9510
      Width           =   2085
   End
   Begin VB.CommandButton cmd_Clear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   18240
      Style           =   1  '그래픽
      TabIndex        =   53
      Top             =   9510
      Width           =   1080
   End
   Begin MSWinsockLib.Winsock LaneSnd_Sock 
      Index           =   0
      Left            =   3075
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneRcv_Sock 
      Index           =   0
      Left            =   15
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmd_GateOpen 
      Caption         =   "Gate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   3
      Left            =   27825
      TabIndex        =   42
      Top             =   7965
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmd_GateOpen 
      Caption         =   "Gate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   2
      Left            =   20655
      TabIndex        =   37
      Top             =   7965
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmd_GateOpen 
      Caption         =   "Gate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   13470
      TabIndex        =   32
      Top             =   7965
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   0
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   150
      TabIndex        =   5
      Top             =   13845
      Width           =   14175
   End
   Begin VB.CommandButton cmd_GateOpen 
      Caption         =   "Gate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   6285
      TabIndex        =   4
      Top             =   7965
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton cmd_Update 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update"
      Height          =   675
      Left            =   46125
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   3
      Top             =   13950
      Width           =   945
   End
   Begin VB.CommandButton cmd_Refresh 
      BackColor       =   &H0000FFFF&
      Caption         =   "Refresh"
      Height          =   675
      Left            =   48225
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   2
      Top             =   13950
      Width           =   945
   End
   Begin VB.CommandButton cmd_Delete 
      BackColor       =   &H0000FFFF&
      Caption         =   "Delete"
      Height          =   675
      Left            =   47190
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   0
      Top             =   13950
      Width           =   945
   End
   Begin xIPs.xIP xIP1 
      Left            =   6660
      Top             =   45
      _ExtentX        =   900
      _ExtentY        =   873
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1680
      Left            =   6300
      TabIndex        =   1
      Top             =   9285
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin LPR_PARKING_HOST.Server Server1 
      Left            =   5220
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   750
      Index           =   0
      Left            =   20415
      TabIndex        =   6
      Top             =   915
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1323
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
      Picture         =   "FrmG4_23.frx":321B6
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   750
      Index           =   1
      Left            =   21765
      TabIndex        =   7
      Top             =   915
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "보호모드"
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
      Picture         =   "FrmG4_23.frx":32507
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   750
      Index           =   4
      Left            =   27150
      TabIndex        =   8
      Top             =   915
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1323
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
      Picture         =   "FrmG4_23.frx":32858
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   750
      Index           =   2
      Left            =   23100
      TabIndex        =   9
      Top             =   915
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1323
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
      Picture         =   "FrmG4_23.frx":32BA9
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   750
      Index           =   3
      Left            =   24450
      TabIndex        =   10
      Top             =   915
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1323
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
      Picture         =   "FrmG4_23.frx":32EFA
   End
   Begin MSWinsockLib.Winsock REC_sock 
      Left            =   5640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SND_sock 
      Left            =   6060
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneRcv_Sock 
      Index           =   1
      Left            =   435
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneRcv_Sock 
      Index           =   2
      Left            =   855
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneRcv_Sock 
      Index           =   3
      Left            =   1275
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneSnd_Sock 
      Index           =   1
      Left            =   3495
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneSnd_Sock 
      Index           =   2
      Left            =   3915
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LaneSnd_Sock 
      Index           =   3
      Left            =   4335
      Top             =   16230
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   1500
      Left            =   14730
      TabIndex        =   52
      Top             =   10935
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   2646
      View            =   3
      LabelEdit       =   1
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Threed.SSCommand cmd_menu 
      Height          =   750
      Index           =   5
      Left            =   25800
      TabIndex        =   74
      Top             =   915
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1323
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
      Picture         =   "FrmG4_23.frx":3324B
   End
   Begin VB.Label lbl_Update 
      BackStyle       =   0  '투명
      Caption         =   "Car No."
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   14730
      TabIndex        =   72
      Top             =   9900
      Width           =   885
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '투명
      Caption         =   " 정기권 간편 검색"
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
      Height          =   300
      Index           =   0
      Left            =   14595
      TabIndex        =   71
      Top             =   8835
      Width           =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   14565
      X2              =   28530
      Y1              =   9180
      Y2              =   9180
   End
   Begin VB.Label lbl_Update 
      BackStyle       =   0  '투명
      Caption         =   "차량번호"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   5
      Left            =   14700
      TabIndex        =   70
      Top             =   9585
      Width           =   1185
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   69
      Top             =   14460
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   68
      Top             =   14130
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   67
      Top             =   13785
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   66
      Top             =   13440
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   65
      Top             =   12750
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   3
      Left            =   14760
      TabIndex        =   64
      Top             =   14460
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   14745
      TabIndex        =   63
      Top             =   14130
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   14760
      TabIndex        =   62
      Top             =   13785
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   14760
      TabIndex        =   61
      Top             =   13440
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   14760
      TabIndex        =   60
      Top             =   13110
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   14775
      TabIndex        =   59
      Top             =   12750
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   58
      Top             =   13110
      Width           =   60
   End
   Begin VB.Label LblSearch 
      BackColor       =   &H00404040&
      Caption         =   "검색결과 : "
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   14730
      TabIndex        =   57
      Top             =   10560
      Width           =   4725
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "등 록 일 시"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   2
      Left            =   14760
      TabIndex        =   56
      Top             =   14805
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   15885
      TabIndex        =   55
      Top             =   14805
      Width           =   60
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "차량관제 프로그램"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   24
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   885
      TabIndex        =   51
      Top             =   990
      Width           =   6945
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      Height          =   1080
      Left            =   345
      Top             =   735
      Width           =   6480
   End
   Begin VB.Label lbl_GN 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      BackStyle       =   0  '투명
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
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   3
      Left            =   21720
      TabIndex        =   50
      Top             =   2115
      Width           =   3375
   End
   Begin VB.Label lbl_GN 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      BackStyle       =   0  '투명
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
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   2
      Left            =   14565
      TabIndex        =   49
      Top             =   2115
      Width           =   3375
   End
   Begin VB.Label lbl_GN 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      BackStyle       =   0  '투명
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
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   1
      Left            =   7395
      TabIndex        =   48
      Top             =   2130
      Width           =   3375
   End
   Begin VB.Label lbl_GN 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      BackStyle       =   0  '투명
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
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   0
      Left            =   210
      TabIndex        =   47
      Top             =   2130
      Width           =   3375
   End
   Begin VB.Image ImageLog 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   4500
      Left            =   240
      Picture         =   "FrmG4_23.frx":3359C
      Stretch         =   -1  'True
      Top             =   9270
      Width           =   6000
   End
   Begin VB.Shape Shp_Rec 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '투명하지 않음
      Height          =   300
      Index           =   3
      Left            =   28335
      Top             =   2070
      Width           =   300
   End
   Begin VB.Label lbl_RecState 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "RecState"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   21.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   3
      Left            =   23445
      TabIndex        =   46
      Top             =   7800
      Width           =   3240
   End
   Begin VB.Label lbl_time_now 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   360
      Index           =   3
      Left            =   24705
      TabIndex        =   45
      Top             =   2130
      Width           =   3405
   End
   Begin VB.Label lbl_carno 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   390
      Index           =   3
      Left            =   26040
      TabIndex        =   44
      Top             =   6870
      Width           =   2565
   End
   Begin VB.Label lbl_LprIP 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "192.168.123.123"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   3
      Left            =   26910
      TabIndex        =   43
      Top             =   7545
      Width           =   1740
   End
   Begin VB.Shape Shp_Rec 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '투명하지 않음
      Height          =   300
      Index           =   2
      Left            =   21165
      Top             =   2070
      Width           =   300
   End
   Begin VB.Label lbl_RecState 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "RecState"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   21.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   2
      Left            =   16275
      TabIndex        =   41
      Top             =   7800
      Width           =   3240
   End
   Begin VB.Label lbl_time_now 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   360
      Index           =   2
      Left            =   17535
      TabIndex        =   40
      Top             =   2130
      Width           =   3405
   End
   Begin VB.Label lbl_carno 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   390
      Index           =   2
      Left            =   18870
      TabIndex        =   39
      Top             =   6870
      Width           =   2565
   End
   Begin VB.Label lbl_LprIP 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "192.168.123.123"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   2
      Left            =   19740
      TabIndex        =   38
      Top             =   7545
      Width           =   1740
   End
   Begin VB.Shape Shp_Rec 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '투명하지 않음
      Height          =   300
      Index           =   1
      Left            =   13980
      Top             =   2070
      Width           =   300
   End
   Begin VB.Label lbl_RecState 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "RecState"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   21.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   1
      Left            =   9090
      TabIndex        =   36
      Top             =   7800
      Width           =   3240
   End
   Begin VB.Label lbl_time_now 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   360
      Index           =   1
      Left            =   10350
      TabIndex        =   35
      Top             =   2130
      Width           =   3405
   End
   Begin VB.Label lbl_carno 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   390
      Index           =   1
      Left            =   11685
      TabIndex        =   34
      Top             =   6870
      Width           =   2565
   End
   Begin VB.Label lbl_LprIP 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "192.168.123.123"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   1
      Left            =   12555
      TabIndex        =   33
      Top             =   7545
      Width           =   1740
   End
   Begin VB.Label LblTime 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Index           =   0
      Left            =   24450
      TabIndex        =   31
      Top             =   135
      Width           =   3915
   End
   Begin VB.Shape Shp_Rec 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '투명하지 않음
      Height          =   300
      Index           =   0
      Left            =   6795
      Top             =   2070
      Width           =   300
   End
   Begin VB.Label lbl_RecState 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "RecState"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   21.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   0
      Left            =   1905
      TabIndex        =   30
      Top             =   7800
      Width           =   3240
   End
   Begin VB.Label lbl_time_now 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   360
      Index           =   0
      Left            =   3165
      TabIndex        =   29
      Top             =   2130
      Width           =   3405
   End
   Begin VB.Label lbl_carno 
      Alignment       =   1  '오른쪽 맞춤
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
      Height          =   390
      Index           =   0
      Left            =   4500
      TabIndex        =   28
      Top             =   6870
      Width           =   2565
   End
   Begin VB.Label lbl_LprIP 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "192.168.123.123"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   0
      Left            =   5370
      TabIndex        =   27
      Top             =   7545
      Width           =   1740
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '투명
      Caption         =   " 운영 현황"
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
      Height          =   300
      Index           =   3
      Left            =   270
      TabIndex        =   25
      Top             =   8835
      Width           =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   240
      X2              =   14250
      Y1              =   9180
      Y2              =   9180
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   24
      Top             =   12900
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   23
      Top             =   12570
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   22
      Top             =   12225
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   21
      Top             =   11895
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   20
      Top             =   11220
      Width           =   60
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "처 리 상 태"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   6495
      TabIndex        =   19
      Top             =   12900
      Width           =   900
   End
   Begin VB.Label Label5 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "종  료  일"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   6480
      TabIndex        =   18
      Top             =   12570
      Width           =   780
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   6495
      TabIndex        =   17
      Top             =   12225
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   6495
      TabIndex        =   16
      Top             =   11880
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   6495
      TabIndex        =   15
      Top             =   11550
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   6495
      TabIndex        =   14
      Top             =   11205
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   13
      Top             =   11550
      Width           =   60
   End
   Begin VB.Label Label6 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '투명
      Caption         =   "처 리 일 시"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   6495
      TabIndex        =   12
      Top             =   13245
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
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   0
      Left            =   7620
      TabIndex        =   11
      Top             =   13245
      Width           =   60
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   5415
      Index           =   0
      Left            =   60
      Picture         =   "FrmG4_23.frx":40969
      Stretch         =   -1  'True
      Top             =   1995
      Width           =   7140
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   5415
      Index           =   3
      Left            =   21600
      Picture         =   "FrmG4_23.frx":6609C
      Stretch         =   -1  'True
      Top             =   1995
      Width           =   7140
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   5415
      Index           =   2
      Left            =   14430
      Picture         =   "FrmG4_23.frx":8B7CF
      Stretch         =   -1  'True
      Top             =   1995
      Width           =   7140
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   5415
      Index           =   1
      Left            =   7245
      Picture         =   "FrmG4_23.frx":B0F02
      Stretch         =   -1  'True
      Top             =   1995
      Width           =   7140
   End
   Begin VB.Label Label3 
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
      Height          =   6870
      Index           =   4
      Left            =   105
      TabIndex        =   26
      Top             =   8670
      Width           =   14280
   End
   Begin VB.Label Label3 
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
      Height          =   6870
      Index           =   2
      Left            =   14430
      TabIndex        =   73
      Top             =   8655
      Width           =   14250
   End
End
Attribute VB_Name = "FrmG4_23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyText(1 To 8) As New clsText
Dim DataField_Enabled As Boolean
Dim Save_TagNum, APS_CMD As String

Private Sub sOutput(strText As String, strIP As String)
    List1.AddItem " " & Format(Now, "yyyy-mm-dd hh:nn:ss") & strText & "     " & strIP, 0
End Sub

Private Sub Socket_ConnectAPS(ByVal IP As String, ByVal Port As Long)
'    'Gate_Winsock.Close
'
'    If (APS_Winsock.State <> sckClosed) Then
'        APS_Winsock.Close
'        'DoEvents
'    End If
'    APS_Winsock.Connect IP, Port
'
'    Call sOutput("[Gate 접속]", IP)
'    List2.AddItem " [Gate 접속] " & IP, 0
'    'Call Err_doc("    [Gate 접속]  시도 IP = " & IP & "    PORT = " & Port)
End Sub

Private Sub APS_Winsock_Connect()
'    Dim bData() As Byte
'
'    ReDim bData(Len(APS_CMD) - 1) As Byte
'    bData = StrConv(APS_CMD, vbFromUnicode)
'    APS_Winsock.SendData bData
'
'    Call sOutput("[Gate 송신]", APS_CMD)
'    List2.AddItem " [Gate 송신] " & APS_CMD, 0
''    If (Check5.value = 1) Then
''        Call Err_doc("    [Gate 송신] " & CMD)
''    End If
'    'Fee_sock.Close
End Sub

Private Sub APS_Winsock_DataArrival(ByVal bytesTotal As Long)
'    Dim strData As String
'
'    APS_Winsock.GetData strData, , bytesTotal
'    Call sOutput("[Gate 수신]", strData)
'    List2.AddItem " [Gate 수신] " & strData, 0
'    APS_Winsock.Close
    
End Sub

Private Sub APS_Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    Call sOutput(Source, "[Gate 소켓] " & "에러 : " & Description)
'    List2.AddItem " [Gate Error] " & Description, 0
''    If (Check5.value = 1) Then
''        Call Err_doc("   [Gate 소켓] " & "에러 : " & Description)
''    End If
End Sub


'차단기 컨트롤
Private Sub cmd_GateOpen_Click(Index As Integer)
    Glo_GateBar_IP = ""
    Glo_GateBar_IP = lbl_LprIP(Index).Caption
    FrmGateBar.Show 1
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim SQL As String
    Dim Reg_Addr As String
    
    Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    'Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    'Left = 0
    Top = 0
    
    Call ListView_Init1
    Call ListView_Init2
    
    FrmTcpServer.Show 0
    'Call Server1.StartServer(Val(Get_Ini("System Config", "Server_Port", "")), Server1.ServerIP)
    For i = 0 To 3
        lbl_GN(i).Caption = ""
        lbl_carno(i).Caption = ""
        lbl_time_now(i).Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")
        lbl_RecState(i).Caption = ""
        Shp_Rec(i).Visible = False
    Next i
    
    lbl_GN(0).Caption = Trim(LANE1_Name)
    lbl_GN(1).Caption = Trim(LANE2_Name)
    lbl_GN(2).Caption = Trim(LANE3_Name)
    lbl_GN(3).Caption = Trim(LANE4_Name)
    
    'Glo_PartName = Get_Ini("System Config", "Server_IP", "None")
    lbl_LprIP(0).Caption = Get_Ini("System Config", "LPR1_IP", "None")
    lbl_LprIP(1).Caption = Get_Ini("System Config", "LPR2_IP", "None")
    lbl_LprIP(2).Caption = Get_Ini("System Config", "LPR3_IP", "None")
    lbl_LprIP(3).Caption = Get_Ini("System Config", "LPR4_IP", "None")
    
    Call cmd_menu_Click(1)
    Timer1.Enabled = True
    FrmTcpServer.Hide
    'Call APS_Check
End Sub

Private Sub cmd_menu_Click(Index As Integer)
    Dim i As Integer

    Me.MousePointer = 11
    Select Case Index
        Case 0
             FrmInOut.Show 1
             Me.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "입출차 보고서 화면 접근")
        Case 1  '보호모드
            If (cmd_menu(1).Caption = "보호모드") Then
               Call DataLogger("[HOST Button]    " & "프로그램 보호모드로 전환")
               cmd_menu(1).Caption = "보호해제"
               For i = 0 To 5
                   If (i <> 1) Then
                      cmd_menu(i).Enabled = False
                   End If
               Next i
               cmd_menu(0).Enabled = True
               cmd_menu(4).Enabled = True
               Put_Ini "System Config", "보호모드", "True"
            Else
               Protect.Show 1
               Call DataLogger("[HOST Button]    " & "프로그램 보호모드 해제")
            End If
            Me.MousePointer = 0
        Case 2
             FrmReg.Show 1
             Me.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "정기권관리 화면 접근")
        Case 3
            FrmTcpServer.Show 0
            Me.MousePointer = 0
            Call DataLogger("[HOST Button]    " & "TCP Server 화면 접근")
        Case 4
            Call Server1.StopServer
            Unload Me
        Case 5
             FrmCSV.Show 1
             Me.MousePointer = 0
             Call DataLogger("[HOST Button]    " & "일괄등록 화면 접근")
    End Select

End Sub

'운영현황 처리===================================================================================================================================
Public Sub ListView_Init1()
Dim Column_to_size As Integer

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " 차량번호     "      '0
    ListView1.ColumnHeaders.Add , , " 구    분         "  '1
    ListView1.ColumnHeaders.Add , , " 이    름  "       '2
    ListView1.ColumnHeaders.Add , , " 전화번호     "  '3
    ListView1.ColumnHeaders.Add , , " 인식번호     "   '4
    ListView1.ColumnHeaders.Add , , " 종 료 일     "        '5
    ListView1.ColumnHeaders.Add , , " 인식상태     "          '6
    ListView1.ColumnHeaders.Add , , " 처리일시     "         '7
    ListView1.ColumnHeaders.Add , , " 입출구분     "    '8
    ListView1.ColumnHeaders.Add , , " 이미지명                                            "    '9
    
    ListView2.ColumnHeaders.Add , , " "
    'ListView2.SortKey = 11
    ListView2.SortOrder = lvwDescending
    ListView2.Sorted = True
    
    For Column_to_size = 0 To ListView2.ColumnHeaders.Count - 2
         SendMessage ListView2.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim Tmp_File As String
    
    ListView1.SetFocus
    LblCar(0).Caption = ""
    LblName(0).Caption = ""
    LblId(0).Caption = ""
    LblCarType(0).Caption = ""
    LblTel(0).Caption = ""
    LblDate(0).Caption = ""
    LblGubun(0).Caption = ""
    LblCar(0).Caption = ListView1.SelectedItem.Text
    LblName(0).Caption = ListView1.SelectedItem.SubItems(2)
    LblId(0).Caption = ListView1.SelectedItem.SubItems(1)
    LblCarType(0).Caption = ListView1.SelectedItem.SubItems(3)
    LblTel(0).Caption = Format(ListView1.SelectedItem.SubItems(5), "0000-00-00")
    LblDate(0).Caption = ListView1.SelectedItem.SubItems(6)
    LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)
        
    Tmp_File = Dir(Trim(ListView1.SelectedItem.SubItems(8)))
    If (Tmp_File <> "") Then
        ImageLog.Picture = LoadPicture(Trim(ListView1.SelectedItem.SubItems(8)))
    Else
        ImageLog.Picture = LoadPicture(App.Path & "\NoCar.jpg")
    End If

End Sub
'운영현황 처리 END ===============================================================================================================================

Public Sub ListView_Init2()
Dim Column_to_size As Integer

    Call ListViewExtended(ListView2)
    ListView2.View = lvwReport
    ListView2.ListItems.Clear
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , " 차량번호      "
    ListView2.ColumnHeaders.Add , , " 이    름      "
    ListView2.ColumnHeaders.Add , , " 구    분        "
    ListView2.ColumnHeaders.Add , , " 연 락 처                  "
    ListView2.ColumnHeaders.Add , , " 차량모델   "
    ListView2.ColumnHeaders.Add , , " 시 작 일          "
    ListView2.ColumnHeaders.Add , , " 만 료 일          "
    ListView2.ColumnHeaders.Add , , " 등록일시                       "
    ListView2.ColumnHeaders.Add , , "  "
    
    For Column_to_size = 0 To ListView2.ColumnHeaders.Count - 2
         SendMessage ListView2.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End Sub
Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
    ListView2.SetFocus
    LblCar(1).Caption = ""
    LblName(1).Caption = ""
    LblId(1).Caption = ""
    LblCarType(1).Caption = ""
    LblTel(1).Caption = ""
    LblDate(1).Caption = ""
    LblGubun(1).Caption = ""
    LblCar(1).Caption = ListView2.SelectedItem.Text
    LblName(1).Caption = ListView2.SelectedItem.SubItems(1)
    LblId(1).Caption = ListView2.SelectedItem.SubItems(2)
    LblCarType(1).Caption = ListView2.SelectedItem.SubItems(3)
    LblTel(1).Caption = ListView2.SelectedItem.SubItems(4)
    If (ListView2.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView2.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
        LblDate(1).ForeColor = vbWhite
        LblDate(1).Caption = ListView2.SelectedItem.SubItems(5) & " - " & ListView2.SelectedItem.SubItems(6)
    Else
        LblDate(1).ForeColor = vbRed
        LblDate(1).Caption = ListView2.SelectedItem.SubItems(5) & " - " & ListView2.SelectedItem.SubItems(6) & "   " & "[기간 에러]"
    End If
    LblGubun(1).Caption = ListView2.SelectedItem.SubItems(7)
End Sub

'정기권 간편검색 Start  ===================================================================================================================================
Private Sub cmd_Clear_Click()
   LblCar(1).Caption = ""
    LblName(1).Caption = ""
    LblId(1).Caption = ""
    LblCarType(1).Caption = ""
    LblTel(1).Caption = ""
    LblDate(1).Caption = ""
    LblGubun(1).Caption = ""
    LblSearch = ""
    ListView2.ListItems.Clear
    Text1 = ""
    Text1.SetFocus
End Sub
Private Sub ListView3_ItemClick(ByVal Item As ComctlLib.ListItem)
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
        LblDate(0).ForeColor = vbWhite
        LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6)
    Else
        LblDate(0).ForeColor = vbRed
        LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6) & "   " & "[기간 에러]"
    End If
    LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Car_Num_Str As String
    Dim Qry As String
    Dim rs As Recordset
    Dim rs_Part As Recordset
    Dim itmX As ListItem
    
    If (KeyAscii = 13) Then
        LblCar(1).Caption = ""
        LblName(1).Caption = ""
        LblId(1).Caption = ""
        LblCarType(1).Caption = ""
        LblTel(1).Caption = ""
        LblDate(1).Caption = ""
        LblGubun(1).Caption = ""
        If ((Len(Text1) <> 4) Or Not (IsNumeric(Text1))) Then
            MsgBox "차량번호 숫자 네자리를 정확하게 입력하세요!"
            Text1 = ""
            Exit Sub
        End If
        Qry = "Select * From tb_reg Where CAR_NO LIKE CONCAT( '%', '" & Text1 & "','%')"
        Set rs = New ADODB.Recordset
        rs.Open Qry, adoConn
        ListView2.ListItems.Clear
        If (rs.EOF) Then
            LblSearch.Caption = "검색결과 : 자료가 존재 하지않습니다.."
        Else
            LblSearch.Caption = "검색결과 : " & (rs.RecordCount) & " 건"
            Do While Not (rs.EOF)
                Set itmX = ListView2.ListItems.Add(, , "" & rs!CAR_NO)
                itmX.SubItems(1) = "" & rs!DRIVER_NAME
                itmX.SubItems(2) = "" & rs!CAR_GUBUN
                itmX.SubItems(3) = "" & rs!DRIVER_PHONE
                itmX.SubItems(4) = "" & rs!CAR_MODEL
                itmX.SubItems(5) = "" & rs!Start_Date
                itmX.SubItems(6) = "" & rs!End_Date
                itmX.SubItems(7) = "" & Format(rs!REG_DATE, "YYYY-MM-DD HH:NN:SS")
                rs.MoveNext
            Loop
            ListView2.ListItems.Item(1).Selected = True
            If (rs.RecordCount = 1) Then
            Else
                ListView2.SetFocus
            End If
            LblCar(1).Caption = ListView2.SelectedItem.Text
            LblName(1).Caption = ListView2.SelectedItem.SubItems(1)
            LblId(1).Caption = ListView2.SelectedItem.SubItems(2)
            LblCarType(1).Caption = ListView2.SelectedItem.SubItems(3)
            LblTel(1).Caption = ListView2.SelectedItem.SubItems(4)
            If (ListView2.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView2.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
                LblDate(1).ForeColor = vbWhite
                LblDate(1).Caption = ListView2.SelectedItem.SubItems(5) & " - " & ListView2.SelectedItem.SubItems(6)
            Else
                LblDate(1).ForeColor = vbRed
                LblDate(1).Caption = ListView2.SelectedItem.SubItems(5) & " - " & ListView2.SelectedItem.SubItems(6) & "   " & "[기간만료]"
            End If
            LblGubun(1).Caption = ListView2.SelectedItem.SubItems(7)
        End If
        Set rs = Nothing
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Public Sub ListView_Init3()
Dim Column_to_size As Integer

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " 차량번호     "      '0
    ListView1.ColumnHeaders.Add , , " 구    분         "  '1
    ListView1.ColumnHeaders.Add , , " 이    름  "       '2
    ListView1.ColumnHeaders.Add , , " 전화번호     "  '3
    ListView1.ColumnHeaders.Add , , " 인식번호     "   '4
    ListView1.ColumnHeaders.Add , , " 종 료 일     "        '5
    ListView1.ColumnHeaders.Add , , " 인식상태     "          '6
    ListView1.ColumnHeaders.Add , , " 처리일시     "         '7
    ListView1.ColumnHeaders.Add , , " 입출구분     "    '8
    ListView1.ColumnHeaders.Add , , " 이미지명                                            "    '9
    
    ListView2.ColumnHeaders.Add , , " "
    'ListView2.SortKey = 11
    ListView2.SortOrder = lvwDescending
    ListView2.Sorted = True
    
    For Column_to_size = 0 To ListView2.ColumnHeaders.Count - 2
         SendMessage ListView2.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

End Sub

'정기권 간편검색 END  ===================================================================================================================================

'프로그램 종료
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg, Style, Title, Response
    Dim Ret As Boolean
    
    msg = "프로그램을 종료하시겠습니까?         "
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Parking Manager™  - JWT   "
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then
        Call Err_doc("호스트 : " & "프로그램 정상적으로 종료")
        Call DataBaseClose(adoConn)
        Call Unhook
        End
    End If
    Me.MousePointer = 0
    Cancel = True
End Sub

'서버 데이터 리시버
Private Sub Server1_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim car_num As String
    Dim Lpr_Cmd As String * 20
    Dim Lpr_CarNum As String * 20
    Dim Lpr_NumType As String * 2
    Dim Lpr_Path As String
    Dim Dns_Path As String
    Dim Tcp_Lpr_Path As String * 100
    Dim Lpr_Color As String * 10
    Dim image_name As String
    Dim Image_Path As String
    Dim url_name As String
'    Dim fso As New FileSystemObject
    Dim tmp_gatenum As Integer
    Dim Mcnt As Integer
    Dim Pos As Integer
    Dim Loopcnt As Long

    Dim Web_Car As String * 20
    Dim Web_date As String * 14
    Dim Web_gubun As String * 2
    Dim Web_gate As String * 2
    Dim Web_gategubun As String
    Dim Web_Golf_Gubun As String
On Error GoTo Err_P

    Debug.Print Data

    If (Mid(Data, 1, 6) = "WEB_DC") Then
        Select Case Mid(Data, 1, 7)
               Case "WEB_DC1"
                    car_num = Mid(Data, 8, LenH(Data) - 8)
                    Dim Qry As String
                    Dim rs As Recordset

                    Set rs = New ADODB.Recordset
                    Qry = "SELECT * FROM ilbancarin Where 차량번호 = '" & car_num & "'"
                    rs.Open Qry, adoConn
                    If (rs.EOF) Then
                        Server1.SendData "WEB_DC1_N", SckIndex
                    Else
                        Web_Car = rs!차량번호
                        Web_date = rs!처리일시
                        Web_gubun = rs!입출구분
                        Web_gate = rs!Gate
                        Server1.SendData "WEB_DC1_Y" & Web_Car & Web_date & Web_gubun & Web_gate, SckIndex
                    End If
                    Set rs = Nothing
               Case "WEB_DC2"
                    Web_gategubun = Trim(Mid(Data, 62, 30))
                    Select Case Web_gategubun
                           Case "렉서스"
                                adoConn.Execute "UPDATE ilbancarin SET Gate = Gate + 5, 구분 = '100%',  게이트구분 = '" & Web_gategubun & "' WHERE 처리일시 = '" & Trim(Mid(Data, 28, 14)) & "' And 차량번호 = '" & Trim(Mid(Data, 8, 20)) & "'"
                                adoConn.Execute "Insert Into tb_WebDC_log Values('" & Format(Now, "YYYYMMDDHHNNSS") & "', '" & Web_gategubun & "', '" & Trim(Mid(Data, 8, 20)) & "', '" & Trim(Mid(Data, 28, 14)) & "', 99)"
                                Server1.SendData "WEB_DC2_Y", SckIndex
                           Case "스크린골프"
                                adoConn.Execute "UPDATE ilbancarin SET Gate = Gate + 1, 입출구분 = 입출구분 + 6,  게이트구분 = '" & Web_gategubun & "' WHERE 처리일시 = '" & Trim(Mid(Data, 28, 14)) & "' And 차량번호 = '" & Trim(Mid(Data, 8, 20)) & "'"
                                adoConn.Execute "Insert Into tb_WebDC_log Values('" & Format(Now, "YYYYMMDDHHNNSS") & "', '" & Web_gategubun & "', '" & Trim(Mid(Data, 8, 20)) & "', '" & Trim(Mid(Data, 28, 14)) & "', 6)"
                                Server1.SendData "WEB_DC2_Y", SckIndex
                           Case "SJ등촌골프장"
                                Web_Golf_Gubun = Trim(Right(Data, 1))
                                If (Web_Golf_Gubun = "0") Then                                 '4시간 무료
                                    adoConn.Execute "UPDATE ilbancarin SET Gate = Gate + 5, 입출구분 = 입출구분 + 4, 구분 = '50%', 게이트구분 = '" & Web_gategubun & "' WHERE 처리일시 = '" & Trim(Mid(Data, 28, 14)) & "' And 차량번호 = '" & Trim(Mid(Data, 8, 20)) & "'"
                                    adoConn.Execute "Insert Into tb_WebDC_log Values('" & Format(Now, "YYYYMMDDHHNNSS") & "', '" & Web_gategubun & "', '" & Trim(Mid(Data, 8, 20)) & "', '" & Trim(Mid(Data, 28, 14)) & "', 4)"
                                Else
                                    adoConn.Execute "UPDATE ilbancarin SET Gate = Gate + 5, 입출구분 = 입출구분 + 16, 구분 = '50%', 게이트구분 = '" & Web_gategubun & "' WHERE 처리일시 = '" & Trim(Mid(Data, 28, 14)) & "' And 차량번호 = '" & Trim(Mid(Data, 8, 20)) & "'"
                                    adoConn.Execute "Insert Into tb_WebDC_log Values('" & Format(Now, "YYYYMMDDHHNNSS") & "', '" & Web_gategubun & "', '" & Trim(Mid(Data, 8, 20)) & "', '" & Trim(Mid(Data, 28, 14)) & "', 16)"
                                End If
                                Server1.SendData "WEB_DC2_Y", SckIndex
                           Case Else
                                adoConn.Execute "UPDATE ilbancarin SET Gate = Gate + 1, 입출구분 = 입출구분 + 2,  게이트구분 = '" & Web_gategubun & "' WHERE 처리일시 = '" & Trim(Mid(Data, 28, 14)) & "' And 차량번호 = '" & Trim(Mid(Data, 8, 20)) & "'"
                                adoConn.Execute "Insert Into tb_WebDC_log Values('" & Format(Now, "YYYYMMDDHHNNSS") & "', '" & Web_gategubun & "', '" & Trim(Mid(Data, 8, 20)) & "', '" & Trim(Mid(Data, 28, 14)) & "', 2)"
                                Server1.SendData "WEB_DC2_Y", SckIndex
                    End Select


               Case "WEB_DC3"
               Case "WEB_DC4"
        End Select
        Exit Sub
    End If

    If Data = "GET_TIME" Then
        Server1.SendData Format(Time, "HH:MM:SS"), SckIndex
        Exit Sub
    End If
    If Data = "GET_DATE" Then
        Server1.SendData Format(Date, "MM/DD/YYYY"), SckIndex
        Exit Sub
    End If
    If (Mid(Data, 1, 8) = "LPR_TEST") Then
        RemoteIP = Trim(Mid(Data, 21, 20))
        Data = Mid(Data, 41, LenH(Data) - 40)
        Server1.SendData "LPR Test Command.", SckIndex
    End If
    'Call sOutput(RemoteIP, Data)
'With FrmG4Mini
'    Call Err_doc("---------------------------------------------------------------------")
'    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LPR 데이터 수신(tcp)]  " & Data)
'    If (.Check1.value = 1) Then
'        '.List1.AddItem "==========================================================================", 0
'        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LPR 데이터 수신(tcp)] " & Data, 0
'    End If
'    If (Len(Data) > 100) Then
'        Server1.SendData Format(Now, "yyyymmddhhnnss"), SckIndex
'        Exit Sub
'    End If
'    Call Form_G4Mini(Data)
'    Server1.SendData Format(Now, "yyyymmddhhnnss"), SckIndex
'
'    If (Glo_Remote_YN = "Y") Then
'        Glo_Remote_Str = Data
'        Call Socket_ConnectRemote(Glo_Remote_IP, Glo_Remote_Port)
'    End If
'
'End With

Exit Sub

Err_P:
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [Server_DataArrival]  " & Err.Description)
End Sub

Private Sub Server1_Error(ByVal SckIndex As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
'Debug.Print Description

End Sub
'
Private Sub Form_G4_23(Data As String)
Dim i As Integer
Dim GateNo As Integer
Dim GateName As String
Dim CarNo As String
Dim rs As Recordset
Dim Qry As String
Dim Tmp_File As String

On Error Resume Next

With FrmG4_23
        GateNo = Left(Data, 1)
        i = LenH(Data)
        CarNo = Mid(Data, 3, (i - 2))

        Qry = "Select * From tb_inout_ENC Where PASS_GATE = '" & GateNo & "' And CAR_NO = '" & CarNo & "' And(PASS_DATE >= '" & Format(Now, "yyyy-mm-dd") & " " & "00:00:00" & "' AND PASS_DATE <= '" & Format(Now, "yyyy-mm-dd") & " " & "23:59:59" & "') Order By PASS_DATE Desc"
        Set rs = New ADODB.Recordset
        rs.Open Qry, adoConn

        If Not (rs.EOF) Then
            .lbl_carno(GateNo).Caption = "" & rs!CAR_NO
            Tmp_File = Dir(rs!PASS_IMAGE)
            If (Tmp_File <> "") Then
                .ImageIn(GateNo).Picture = LoadPicture(rs!PASS_IMAGE)
            End If
            For i = 0 To 3
                .Shp_Rec(i).Visible = False
            Next i
            .Shp_Rec(GateNo).Visible = True
            .lbl_time_now(GateNo).Caption = "" & rs!PASS_DATE
            .lbl_RecState(GateNo).Caption = "" & rs!PASS_RESULT
            If rs!PASS_YN = "Y" Then
                .lbl_RecState(GateNo).ForeColor = vbBlue
            Else
                .lbl_RecState(GateNo).ForeColor = vbRed
            End If
            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " GateNo : " & GateNo & ", 차량번호 : " & rs!CAR_NO & ", 처리결과 : " & rs!PASS_RESULT, 0
        Else
            'Beep
        End If
        Set rs = Nothing
End With

End Sub

Private Sub Timer1_Timer()
Dim Qry As String
Dim rs As ADODB.Recordset
Dim i As Integer

LblTime(0).Caption = "현재시간 : " & Format(Now, "yyyy년mm월dd일 hh시nn분ss초")

End Sub


'Private Sub APS_Check()
'    Dim QRY As String
'    Dim rs As ADODB.Recordset
'    Dim Total_Accnt As Long
'
'    QRY = "SELECT * From tb_account"
'    Set rs = New ADODB.Recordset
'    rs.Open QRY, adoConn
'
'    If (rs.EOF = False) Then
'        lbl_ApsState(0).Caption = "" & (rs!BILL_S10000 + rs!BILL_S5000 + rs!BILL_S1000)
'        lbl_ApsState(1).Caption = "" & (rs!COIN_C500 + rs!COIN_C100)
'        lbl_ApsState(2).Caption = "" & rs!BILL_H5000
'        lbl_ApsState(3).Caption = "" & rs!BILL_H1000
'        lbl_ApsState(4).Caption = "" & rs!COIN_H500
'        lbl_ApsState(5).Caption = "" & rs!COIN_H100
'
'        lbl_Update(0).Caption = "Update Date : " & rs!Update_date
'
'        Total_Accnt = 0
'        Total_Accnt = (10000 * rs!BILL_S10000) + (5000 * rs!BILL_S5000) + (1000 * rs!BILL_S1000)
'        Total_Accnt = Total_Accnt + (5000 * rs!BILL_H5000) + (1000 * rs!BILL_H1000)
'        Total_Accnt = Total_Accnt + (500 * rs!COIN_H500) + (100 * rs!COIN_H100)
'        Total_Accnt = Total_Accnt + (500 * rs!COIN_C500) + (100 * rs!COIN_C100)
'        txt_Total.Text = Total_Accnt & "원 "
'
'        lbl_Alarm(0).BackColor = &HFFFFFF
'        lbl_Alarm(1).BackColor = &HFFFFFF
'        lbl_Alarm(2).BackColor = &HFFFFFF
'        lbl_Alarm(3).BackColor = &HFFFFFF
'        lbl_Alarm(4).BackColor = &HFFFFFF
'        lbl_Alarm(5).BackColor = &HFFFFFF
'
'        If (lbl_ApsState(0).Caption > 900) Then
'            lbl_Alarm(0).BackColor = &HFFFF&
'        End If
'        If (lbl_ApsState(1).Caption > 2000) Then
'            lbl_Alarm(1).BackColor = &HFFFF&
'        End If
'        If (lbl_ApsState(2).Caption < 20) Then
'            lbl_Alarm(2).BackColor = &HFFFF&
'        End If
'        If (lbl_ApsState(3).Caption < 50) Then
'            lbl_Alarm(3).BackColor = &HFFFF&
'        End If
'        If (lbl_ApsState(4).Caption < 20) Then
'            lbl_Alarm(4).BackColor = &HFFFF&
'        End If
'        If (lbl_ApsState(5).Caption < 100) Then
'            lbl_Alarm(5).BackColor = &HFFFF&
'        End If
'    Else
'        'Insert 0 초기화
'        QRY = "INSERT INTO tb_account VALUES (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sysdate(now()), "")"
'        Set rs = New ADODB.Recordset
'        rs.Open QRY, adoConn
'
'    End If
'    Set rs = Nothing
'End Sub


Private Sub Socket_ConnectRemote(ByVal IP As String, ByVal Port As Long)
'    'Gate_Winsock.Close
'
'    If (Remote_Winsock.State <> sckClosed) Then
'        Remote_Winsock.Close
'        'DoEvents
'    End If
'    Remote_Winsock.Connect IP, Port
'
'    'Call sOutput("[Gate 접속]", CMD & "  " & IP)
'    'Call Err_doc("    [Gate 접속]  시도 IP = " & IP & "    PORT = " & Port)
End Sub

Private Sub Remote_Winsock_Connect()
'
'    Dim bData() As Byte
'
'    ReDim bData(Len(Glo_Remote_Str) - 1) As Byte
'    bData = StrConv(Glo_Remote_Str, vbFromUnicode)
'    Remote_Winsock.SendData bData
'    'Call sOutput("[Gate 송신]", CMD)
''    If (Check5.value = 1) Then
''        Call Err_doc("    [Gate 송신] " & CMD)
''    End If
'    'Fee_sock.Close
End Sub

Private Sub Remote_Winsock_DataArrival(ByVal bytesTotal As Long)
'    Dim strData As String
'
'    Remote_Winsock.GetData strData, , bytesTotal
'
'    'Debug.Print strData
'    'Call sOutput("[Gate 수신]", strData)
'    Remote_Winsock.Close
'
End Sub

Private Sub Remote_Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Call sOutput(Source, "[Gate 소켓] " & "에러 : " & Description)
'    If (Check5.value = 1) Then
'        Call Err_doc("   [Gate 소켓] " & "에러 : " & Description)
'    End If
End Sub


