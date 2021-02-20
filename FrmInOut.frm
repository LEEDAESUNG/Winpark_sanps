VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmInOut 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   12075
   ClientLeft      =   4710
   ClientTop       =   2115
   ClientWidth     =   17190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   805
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   1146
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7BEAD&
      BorderStyle     =   0  '없음
      Height          =   7200
      Left            =   180
      TabIndex        =   50
      Top             =   1080
      Width           =   9600
      Begin VB.OptionButton opt_ChartGraph 
         BackColor       =   &H00404040&
         Caption         =   "게이트구분그래프"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   4
         Left            =   5760
         TabIndex        =   56
         Top             =   30
         Width           =   1755
      End
      Begin VB.OptionButton opt_ChartGraph 
         BackColor       =   &H00404040&
         Caption         =   "차량구분그래프"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   3
         Left            =   4080
         TabIndex        =   55
         Top             =   30
         Width           =   1695
      End
      Begin VB.OptionButton opt_ChartGraph 
         BackColor       =   &H00404040&
         Caption         =   "시간그래프"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   30
         TabIndex        =   54
         Top             =   30
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton opt_ChartGraph 
         BackColor       =   &H00404040&
         Caption         =   "월간그래프"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   2730
         TabIndex        =   53
         Top             =   30
         Width           =   1365
      End
      Begin VB.OptionButton opt_ChartGraph 
         BackColor       =   &H00404040&
         Caption         =   "요일그래프"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   1
         Left            =   1380
         TabIndex        =   52
         Top             =   30
         Width           =   1365
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   6720
         Left            =   30
         OleObjectBlob   =   "FrmInOut.frx":0000
         TabIndex        =   51
         Top             =   450
         Width           =   9540
      End
   End
   Begin VB.ComboBox Combo5 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FrmInOut.frx":2648
      Left            =   14400
      List            =   "FrmInOut.frx":264A
      Style           =   2  '드롭다운 목록
      TabIndex        =   49
      Top             =   3855
      Width           =   1950
   End
   Begin VB.OptionButton opt_inout 
      BackColor       =   &H00404040&
      Caption         =   "주차내역"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   14400
      TabIndex        =   47
      Top             =   2550
      Width           =   1845
   End
   Begin VB.OptionButton opt_inout 
      BackColor       =   &H00404040&
      Caption         =   "입출차내역"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   11565
      TabIndex        =   46
      Top             =   2550
      Value           =   -1  'True
      Width           =   1845
   End
   Begin VB.TextBox txt_Count 
      Alignment       =   1  '오른쪽 맞춤
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   15465
      TabIndex        =   36
      Top             =   1245
      Width           =   630
   End
   Begin VB.ComboBox cmbHo 
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
      Left            =   13485
      TabIndex        =   33
      Top             =   5070
      Width           =   1290
   End
   Begin VB.ComboBox cmbDong 
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
      Left            =   11565
      TabIndex        =   32
      Top             =   5070
      Width           =   1290
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   18090
      Top             =   810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmd_Update 
      Caption         =   "변 경"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   17790
      TabIndex        =   31
      Top             =   6405
      Width           =   1380
   End
   Begin VB.TextBox txt_Update 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   10  '한글 
      Left            =   11565
      MaxLength       =   10
      TabIndex        =   30
      Top             =   6450
      Width           =   3495
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FrmInOut.frx":264C
      Left            =   11565
      List            =   "FrmInOut.frx":264E
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   3855
      Width           =   1950
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FrmInOut.frx":2650
      Left            =   11565
      List            =   "FrmInOut.frx":265A
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   4590
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FrmInOut.frx":2672
      Left            =   11565
      List            =   "FrmInOut.frx":2674
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   4230
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      IMEMode         =   10  '한글 
      Index           =   2
      Left            =   11565
      MaxLength       =   10
      TabIndex        =   1
      ToolTipText     =   "두 글자 이상 입력하세요"
      Top             =   5550
      Width           =   3225
   End
   Begin Threed.SSCommand Command1 
      Height          =   615
      Left            =   17790
      TabIndex        =   0
      Top             =   5280
      Width           =   1620
      _Version        =   65536
      _ExtentX        =   2857
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "검 색"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":2676
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   540
      Left            =   14445
      TabIndex        =   5
      Top             =   45
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   952
      _StockProps     =   78
      Caption         =   "저장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":29C7
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   540
      Left            =   15705
      TabIndex        =   6
      Top             =   45
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   952
      _StockProps     =   78
      Caption         =   "닫기"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":2D18
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   11565
      TabIndex        =   7
      Top             =   3030
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
      Format          =   288882688
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   14400
      TabIndex        =   8
      Top             =   3030
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
      Format          =   288882688
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   11565
      TabIndex        =   9
      Top             =   3405
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
      Format          =   288882691
      UpDown          =   -1  'True
      CurrentDate     =   36927
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   14475
      TabIndex        =   10
      Top             =   8550
      Width           =   2520
      _Version        =   65536
      _ExtentX        =   4445
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   "  검색 건수"
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
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   1
      Begin VB.Label LblRecordCount 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000000&
         Caption         =   "000"
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
         Height          =   270
         Left            =   1170
         TabIndex        =   11
         Top             =   60
         Width           =   1275
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3000
      Left            =   165
      TabIndex        =   12
      Top             =   9000
      Width           =   16860
      _ExtentX        =   29739
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
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   345
      Left            =   10335
      TabIndex        =   37
      Top             =   1245
      Width           =   1950
      _ExtentX        =   3440
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
      Format          =   288882688
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   345
      Left            =   12870
      TabIndex        =   38
      Top             =   1245
      Width           =   1950
      _ExtentX        =   3440
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
      Format          =   288882688
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   540
      Left            =   12585
      TabIndex        =   44
      Top             =   45
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   952
      _StockProps     =   78
      Caption         =   "방문증발급내역"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":3069
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   14400
      TabIndex        =   45
      Top             =   3405
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
      Format          =   288948227
      UpDown          =   -1  'True
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   645
      Index           =   3
      Left            =   11565
      TabIndex        =   57
      ToolTipText     =   "아래에서 선택한 주차내역을 삭제합니다."
      Top             =   7785
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "선택삭제"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":33BA
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   645
      Index           =   0
      Left            =   13380
      TabIndex        =   58
      ToolTipText     =   "조회기간 내의 모든 주차내역을 삭제합니다."
      Top             =   7785
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "기간삭제"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":370B
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   645
      Index           =   1
      Left            =   15180
      TabIndex        =   60
      ToolTipText     =   "차량 주차요금을 출구무인기로 전송합니다."
      Top             =   7785
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "무인기전송"
      ForeColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":3A5C
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   615
      Left            =   15270
      TabIndex        =   61
      Top             =   4830
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "검 색"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   15
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":3DAD
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   615
      Left            =   15300
      TabIndex        =   62
      Top             =   5550
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "그래프"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":40FE
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   585
      Left            =   15300
      TabIndex        =   63
      Top             =   6450
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "변 경"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   15
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":444F
   End
   Begin Threed.SSCommand cmd_ExSch 
      Height          =   525
      Left            =   15450
      TabIndex        =   64
      Top             =   1710
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   926
      _StockProps     =   78
      Caption         =   "중복조회"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      RoundedCorners  =   0   'False
      Picture         =   "FrmInOut.frx":446B
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "삭      제"
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
      Left            =   10335
      TabIndex        =   59
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Label lbl_back 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "후방"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   16425
      TabIndex        =   48
      Top             =   3870
      Width           =   525
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   686
      X2              =   1124
      Y1              =   162
      Y2              =   162
   End
   Begin VB.Label lbl_CaNo 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   24
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   570
      Index           =   2
      Left            =   10350
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "부터"
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
      Height          =   285
      Index           =   14
      Left            =   12345
      TabIndex        =   42
      Top             =   1275
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "까지"
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
      Height          =   285
      Index           =   15
      Left            =   14910
      TabIndex        =   41
      Top             =   1275
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "회 이상"
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
      Height          =   285
      Index           =   16
      Left            =   16125
      TabIndex        =   40
      Top             =   1275
      Width           =   735
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "중복입차한 일반차량 검색"
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
      Height          =   285
      Index           =   17
      Left            =   10320
      TabIndex        =   39
      Top             =   1710
      Width           =   2595
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "호"
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
      Height          =   345
      Left            =   14820
      TabIndex        =   35
      Top             =   5115
      Width           =   585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "동"
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
      Height          =   345
      Left            =   12915
      TabIndex        =   34
      Top             =   5115
      Width           =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   686
      X2              =   1124
      Y1              =   418
      Y2              =   418
   End
   Begin VB.Label lbl_time_now 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
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
      Left            =   5985
      TabIndex        =   29
      Top             =   1305
      Width           =   3510
   End
   Begin VB.Label lbl_CaNo 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   24
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   570
      Index           =   1
      Left            =   11565
      TabIndex        =   28
      Top             =   7110
      Width           =   3495
   End
   Begin VB.Label lbl_CaNo 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   24
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   570
      Index           =   0
      Left            =   10320
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "인식번호"
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
      Index           =   11
      Left            =   10290
      TabIndex        =   26
      Top             =   7200
      Width           =   1200
   End
   Begin VB.Label lbl_Update 
      BackStyle       =   0  '투명
      Caption         =   "Update : 2012-05-20 12:12:59"
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
      Index           =   0
      Left            =   285
      TabIndex        =   25
      Top             =   690
      Width           =   2925
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '투명
      Caption         =   " 차량 입출내역 조회"
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
      Left            =   210
      TabIndex        =   24
      Top             =   180
      Width           =   4185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   12
      X2              =   1136
      Y1              =   42
      Y2              =   42
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "차량번호"
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
      Index           =   12
      Left            =   10290
      TabIndex        =   23
      Top             =   6525
      Width           =   1200
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "차량번호"
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
      Index           =   10
      Left            =   10245
      TabIndex        =   22
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "까지"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   9
      Left            =   16425
      TabIndex        =   21
      Top             =   3060
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "까지"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   8
      Left            =   16425
      TabIndex        =   20
      Top             =   3450
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "부터"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   7
      Left            =   13590
      TabIndex        =   19
      Top             =   3060
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "부터"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   6
      Left            =   13590
      TabIndex        =   18
      Top             =   3450
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회기간"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   5
      Left            =   10335
      TabIndex        =   17
      Top             =   3060
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "정렬순서"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   10335
      TabIndex        =   16
      Top             =   4620
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회시간"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   3
      Left            =   10335
      TabIndex        =   15
      Top             =   3450
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "게이트구분"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   10335
      TabIndex        =   14
      Top             =   3870
      Width           =   1125
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "입출상태"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   10335
      TabIndex        =   13
      Top             =   4230
      Width           =   900
   End
   Begin VB.Image Image3 
      Height          =   7200
      Left            =   180
      Picture         =   "FrmInOut.frx":4487
      Stretch         =   -1  'True
      Top             =   1005
      Width           =   9600
   End
End
Attribute VB_Name = "FrmInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim excel_sql_str As String
Dim Back_Start_Index As Integer
Dim ChartArray()

Private Sub cmd_Button_Click(Index As Integer)
    If (Index = 3) Then '선택삭제
        Call DeleteInOut("DELETE_SELECT")
        
    ElseIf (Index = 0) Then '기간삭제
        Call DeleteInOut("DELETE_RANGE")
        
    ElseIf (Index = 1) Then '무인기전송
        If (Len(Trim(txt_Update)) > 0) Then
        
            MBox.Label3.FontSize = 16
            MBox.Label3.Caption = Trim(txt_Update)
            MBox.Label2.Caption = "출구무인기 전송"
            MBox.Label1.Caption = "위 차량 주차요금을 " & vbCrLf & "출구무인기로 전송하겠습니까?"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
                Glo_APS_Str = Trim(txt_Update)
                Call DataLogger("[HOST SYSTEM] APS전송 ==> " & Glo_APS_Str)
                Call APS_Connect
            End If
                
        Else
            Msg_Box.Label2.Caption = "출구무인기 전송"
            Msg_Box.Label1.Caption = "출구무인기로 주차요금을 전송할" & vbCrLf & "차량을 선택하세요."
            Msg_Box.Show 1
            Exit Sub
        End If
            
            
                
    End If
End Sub

'Custum Search
Private Sub cmd_ExSch_Click()
Dim i As Integer
Dim Cnt As Integer
Dim Current_Date As String
Dim TmpPath As String
Dim Tmp_File As String
Dim InsSQL As String
Dim Now_Flag As Boolean
Dim sql_str As String
Dim Sort_Order As String

Normal_Search_F = False

Me.MousePointer = 11

Glo_SQL_SEARCH = ""

If IsNumeric(Trim(txt_Count.text)) Then
    If (txt_Count.text = 0) Then
        MsgBox " 올바른 숫자를 입력하세요...!! "
        Me.MousePointer = 0
        Exit Sub
    Else
    
    End If
Else
    MsgBox " 숫자를 입력하세요...!! "
    Me.MousePointer = 0
    Exit Sub
End If

'쿼리 구성
'sql_str = "SELECT CAR_NO, IN_COUNT From (SELECT tb_inout.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT From tb_inout Where tb_inout.PASS_YN = 'N' AND tb_inout.PASS_INOUT = 'IN' AND tb_inout.PASS_DATE >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' AND tb_inout.PASS_DATE <= '" & Format(DTPicker6, "yyyy-mm-dd") & " 23:59:59' Group By tb_inout.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & Val(txt_Count.Text) & ""
sql_str = "SELECT CAR_NO, IN_COUNT From (SELECT tb_inout.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT From tb_inout Where tb_inout.PASS_INOUT = 'IN' AND (tb_inout.PASS_RESULT = '미등록입차' OR tb_inout.PASS_RESULT= '미인식입차') AND tb_inout.PASS_DATE >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' AND tb_inout.PASS_DATE <= '" & Format(DTPicker6, "yyyy-mm-dd") & " 23:59:59' Group By tb_inout.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & Val(txt_Count.text) & ""
'Debug.Print sql_str

Glo_SQL_SEARCH = sql_str

Call ListView_Draw_ParkOne

'SSPanel3(1).Caption = ""
Me.MousePointer = 0
On Error Resume Next

End Sub

Public Sub ListView_Draw_ParkOne()
Dim Column_to_size As Integer
Dim rs As Recordset
Dim qry As String
Dim itmX As ListItem
Dim INDEX_NO As Long
    
    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " CAR_NO         "
    ListView1.ColumnHeaders.Add , , " IN_Count       "
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
 
    Set rs = New ADODB.Recordset
    'Debug.Print Glo_SQL_SEARCH
    rs.Open Glo_SQL_SEARCH, adoConn
    LblRecordCount = rs.RecordCount

    INDEX_NO = 1

    Do While Not (rs.EOF)
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        itmX.SubItems(1) = "" & rs!CAR_NO
        itmX.SubItems(2) = "" & rs!IN_COUNT
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    INDEX_NO = 0
    Set rs = Nothing

End Sub



Private Sub cmd_Update_Click()
    Dim rs As Recordset
    Dim sQry As String
    Dim bQryResult As Boolean

    If (txt_Update.text <> "") And (lbl_time_now.Caption <> "") Then
        'adoConn.Execute "Update tb_inout Set CAR_NO = '" & Trim(txt_Update.Text) & "' Where PASS_DATE = '" & lbl_time_now & "'"
        
        Set rs = New ADODB.Recordset
        sQry = "Update tb_inout Set CAR_NO = '" & Trim(txt_Update.text) & "' Where PASS_DATE = '" & lbl_time_now & "'"
        bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
        If (bQryResult = False) Then
            Call DataLogger("[FrmInOut]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
'        sQry = "Update tb_guestReg_inout Set CAR_NO = '" & Trim(txt_Update.text) & "' Where PASS_DATE = '" & lbl_time_now & "'"
'        bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
'        If (bQryResult = False) Then
'            Call DataLogger("[FrmInOut]    " & "네트워크 및 DB 점검바랍니다")
'            Exit Sub
'        End If
        
    Else
    
    End If

End Sub

'검색 버튼
Private Sub Command1_Click()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim Sort_Order As String
    Dim sql_str As String
    
    Me.MousePointer = 11
    SSCommand1.Enabled = False
    
    If (Combo3.List(Combo3.ListIndex) = "오름차순") Then
        Sort_Order = "PASS_DATE"

    ElseIf (Combo3.List(Combo3.ListIndex) = "내림차순") Then
        Sort_Order = "PASS_DATE DESC"
    End If
    
    
    If (opt_inout(0).value = True) Then '입출차조회
        If (Combo1.ListIndex = 0) Then
            sql_str = "SELECT * FROM tb_inout WHERE (PASS_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & " 000') AND (PASS_DATE <='" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & " 999')"
        Else
            sql_str = "SELECT * FROM tb_inout WHERE (PASS_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & " 000') AND (PASS_DATE <='" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & " 999') AND " & "(PASS_RESULT = '" & Combo1.List(Combo1.ListIndex) & "')"
        End If
    Else
        If (Combo1.ListIndex = 0) Then
            sql_str = "SELECT * FROM tb_now WHERE (PASS_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & " 000') AND (PASS_DATE <='" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & " 999')"
        Else
            sql_str = "SELECT * FROM tb_now WHERE (PASS_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & " 000') AND (PASS_DATE <='" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & " 999') AND " & "(PASS_RESULT = '" & Combo1.List(Combo1.ListIndex) & "')"
        End If
    End If
    
    Select Case Combo4.ListIndex
        Case 0
            '전체
        Case 1
            sql_str = sql_str & " AND (PASS_GATE = '0')"
        Case 2
            sql_str = sql_str & " AND (PASS_GATE = '1')"
        Case 3
            sql_str = sql_str & " AND (PASS_GATE = '2')"
        Case 4
            sql_str = sql_str & " AND (PASS_GATE = '3')"
        Case 5
            sql_str = sql_str & " AND (PASS_GATE = '4')"
        Case 6
            sql_str = sql_str & " AND (PASS_GATE = '5')"
    End Select
    
    If (InStr(Combo5.text, LANE1_Name) > 0) Then
        sql_str = sql_str & " AND (PASS_GATE = '6')"
    ElseIf (InStr(Combo5.text, LANE2_Name) > 0) Then
        sql_str = sql_str & " AND (PASS_GATE = '7')"
    ElseIf (InStr(Combo5.text, LANE3_Name) > 0) Then
        sql_str = sql_str & " AND (PASS_GATE = '8')"
    ElseIf (InStr(Combo5.text, LANE4_Name) > 0) Then
        sql_str = sql_str & " AND (PASS_GATE = '9')"
    ElseIf (InStr(Combo5.text, LANE5_Name) > 0) Then
        sql_str = sql_str & " AND (PASS_GATE = '10')"
    ElseIf (InStr(Combo5.text, LANE6_Name) > 0) Then
        sql_str = sql_str & " AND (PASS_GATE = '11')"
    End If
    
    'If (cmbDong.List(cmbDong.ListIndex) <> "") Then
    If (cmbDong.text <> "") Then
        'sql_str = sql_str & " AND (DRIVER_DEPT = '" & Trim(cmbDong.List(cmbDong.ListIndex)) & "')"
        sql_str = sql_str & " AND (DRIVER_DEPT = '" & cmbDong.text & "')"
    End If
    'If (cmbHo.List(cmbHo.ListIndex) <> "") Then
    If (cmbHo.text <> "") Then
        'sql_str = sql_str & " AND (DRIVER_CLASS = '" & Trim(cmbHo.List(cmbHo.ListIndex)) & "')"
        sql_str = sql_str & " AND (DRIVER_CLASS = '" & cmbHo.text & "')"
    End If
    
'''    If (Text1(2).Text = "") Then
'''    Else
'''        If ((Len(Text1(2)) = 4) And (IsNumeric(Text1(2)))) Then
'''            sql_str = sql_str & " AND (CAR_NO LIKE '%" & Text1(2).Text & "')"
'''        Else
'''            sql_str = sql_str & " AND (CAR_NO = '" & Text1(2).Text & "')"
'''        End If
'''    End If
'''
'''    If Len(Text1(2)) <> 0 Then
'''        Glo_JIOSch = sql_str
'''    Else
'''        Glo_JIOSch = sql_str & " ORDER BY " & Sort_Order
'''    End If
'''
'''    Call ListView_Draw
'''    SSCommand1.Enabled = True
'''    Me.MousePointer = 0
    
    If (LenH(Text1(2)) = 0) Then
    ElseIf (LenH(Text1(2)) < 2) Then
        Text1(2).text = ""
        Text1(2).SetFocus
        Me.MousePointer = 0
        Exit Sub
    Else
        sql_str = sql_str & " AND (CAR_NO LIKE '%" & Text1(2).text & "%')"
    End If


    'Glo_JIOSch = sql_str & " ORDER BY " & Sort_Order
    Glo_JIOSch = sql_str '레코드 건수 많을 경우 order by

    Call ListView_Draw
    SSCommand1.Enabled = True
    Me.MousePointer = 0
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim Car_Num_Str As String
    Dim qry As String
    Dim rs As Recordset
    Dim rs_Part As Recordset
    Dim itmX As ListItem
    
    
        
    If (KeyAscii = 13) Then
        If (Len(Text1(2)) <> 0) Then
            Call Command1_Click
            Exit Sub
        ElseIf (Me.ActiveControl.name = "Text1") Then
            Call Command1_Click
            Exit Sub
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If
    
End Sub


Private Sub Init_Chart()
    
    '차트
    Frame1.Visible = False

    '차량이미지
    Image3.Visible = True
End Sub
Private Sub Form_Load()
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    'Left = 0
    'Top = 0
    
    Call Init_Chart
    
    If (Glo_GUEST_LANE1_YN = "Y" Or Glo_GUEST_LANE2_YN = "Y" Or Glo_GUEST_LANE3_YN = "Y" Or Glo_GUEST_LANE4_YN = "Y" Or Glo_GUEST_LANE5_YN = "Y" Or Glo_GUEST_LANE6_YN = "Y") Then
    Else
        SSCommand3.Visible = False
        SSCommand3.Enabled = False
    End If
    
    lbl_Update(0).Caption = "Update : " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    DTPicker3.Format = dtpCustom
    DTPicker3.CustomFormat = "HH:mm:ss"
    DTPicker3.Refresh
    
    DTPicker4.Format = dtpCustom
    DTPicker4.CustomFormat = "HH:mm:ss"
    DTPicker4.Refresh

    DTPicker1.value = Now
    DTPicker2.value = Now
    DTPicker3.value = Format("00:00:00")
    DTPicker4.value = Format("23:59:59")
    
    DTPicker5.value = Now
    DTPicker6.value = Now
    
    
    If (Glo_User_Type = "구분1/구분2") Then
            Label3.Caption = "소속"
            Label6.Caption = "직급"
    Else
            Label3.Caption = "동"
            Label6.Caption = "호"
    End If
'''    Call Set_cmbDong
'''    Call Set_cmbHo
        
        
    With Combo1
            .AddItem "전체"
            .AddItem "정상입차"
            .AddItem "정상출차"
            .AddItem "기간위반입차"
            .AddItem "기간위반출차"
            .AddItem "미등록입차"
            .AddItem "미등록출차"
            .AddItem "미인식입차"
            .AddItem "미인식출차"
            .AddItem "출입제한입차"
            .AddItem "출입제한출차"
            .AddItem "영업용입차"
            .AddItem "영업용출차"
            .AddItem "방문예약입차"
            .AddItem "방문예약출차"
            
            If (Glo_WEEK_YN = "Y") Then '요일제
                .AddItem "요일위반입차"
                .AddItem "요일위반출차"
            End If
            If (Glo_ROTATION <> "미적용") Then '부제
                .AddItem "부제위반입차"
                .AddItem "부제위반출차"
            End If

'            .Text = Combo1.List(0)
    End With
    Combo1.ListIndex = 0
    
    'Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    
    
    With Combo4
            .AddItem "전체"
            If (Glo_Screen_No >= 1) Then
                .AddItem LANE1_Name
                Back_Start_Index = 4
            End If
            If (Glo_Screen_No >= 2) Then
                .AddItem LANE2_Name
                Back_Start_Index = 4
            End If
            If (Glo_Screen_No >= 4) Then
                .AddItem LANE3_Name
                .AddItem LANE4_Name
                Back_Start_Index = 4
            End If
            If (Glo_Screen_No >= 6) Then
                .AddItem LANE5_Name
                .AddItem LANE6_Name
            End If
'            .Text = Combo4.List(0)
    End With
    Combo4.ListIndex = 0
    
    
    With Combo5
    If (Glo_Lane1_Back_YN = "Y" Or Glo_Lane2_Back_YN = "Y" Or Glo_Lane3_Back_YN = "Y" Or Glo_Lane4_Back_YN = "Y" Or Glo_Lane5_Back_YN = "Y" Or Glo_Lane6_Back_YN = "Y") Then
            .Visible = True
            lbl_back.Visible = True
            .AddItem "전체"
            If (Glo_Lane1_Back_YN = "Y") Then
                .AddItem LANE1_Name & "(후방)"
            End If
            If (Glo_Lane2_Back_YN = "Y") Then
                .AddItem LANE2_Name & "(후방)"
            End If
            If (Glo_Lane3_Back_YN = "Y") Then
                .AddItem LANE3_Name & "(후방)"
            End If
            If (Glo_Lane4_Back_YN = "Y") Then
                .AddItem LANE4_Name & "(후방)"
            End If
            If (Glo_Lane5_Back_YN = "Y") Then
                .AddItem LANE5_Name & "(후방)"
            End If
            If (Glo_Lane6_Back_YN = "Y") Then
                .AddItem LANE6_Name & "(후방)"
            End If
            Combo5.ListIndex = 0
    Else
        .Visible = False
        lbl_back.Visible = False
    End If
    End With
    
    Call opt_inout_Click(0)
    If ((LANE1_YN = "Y" And LANE1_Inout = "출구") Or (LANE2_YN = "Y" And LANE2_Inout = "출구") Or (LANE3_YN = "Y" And LANE3_Inout = "출구") Or (LANE4_YN = "Y" And LANE4_Inout = "출구") Or (LANE5_YN = "Y" And LANE5_Inout = "출구") Or (LANE6_YN = "Y" And LANE6_Inout = "출구")) Then
        opt_inout(0).Visible = True
        opt_inout(1).Visible = True
    Else
        opt_inout(0).Visible = False
        opt_inout(1).Visible = False
    End If
    
    txt_Count = "2"
    
    Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")
    
    '오늘날짜 데이터만
    'Glo_JIOSch = "SELECT * FROM tb_inout WHERE (PASS_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (PASS_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY PASS_DATE "
    Glo_JIOSch = "SELECT * FROM tb_inout WHERE (PASS_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') ORDER BY PASS_DATE "
    
    Call ListView_Draw
Exit Sub

Err_p:
    MsgBox "데이터 베이스 연결실패" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "네트웍 관리자에게 문의 바랍니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "데이터 베이스 연결전에는 자료검색 기능을 수행할수 없습니다.", vbCritical
End Sub

Private Sub Set_cmbDong()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
On Error GoTo Err_p

    qry = "SELECT tb_inout.DRIVER_DEPT From tb_inout Group By tb_inout.DRIVER_DEPT"

    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
        Call DataLogger("[FrmInOut]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    cmbDong.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmbDong.AddItem "" & rs!DRIVER_DEPT
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing

Exit Sub
Err_p:
    Call DataLogger("[FrmInOut Set_cmbDong]    " & Err.Description)
End Sub


Private Sub Set_cmbHo()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
On Error GoTo Err_p

    qry = "SELECT tb_inout.DRIVER_CLASS From tb_inout Group By tb_inout.DRIVER_CLASS"
    
    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
        Call DataLogger("[FrmInOut]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    cmbHo.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmbHo.AddItem "" & rs!DRIVER_CLASS
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
Exit Sub

Err_p:
    Call DataLogger("[FrmInOut Set_cmbHo]    " & Err.Description)
End Sub


'Private Sub Image1_Click()
'    Frame1.Visible = False
'    Image3.Visible = True
'
'    Call Command1_Click
'End Sub
'

'Private Sub Image2_Click()
'    Call cmd_Update_Click
'    Call Command1_Click
'End Sub

'Private Sub Image4_Click()
'    Frame1.Visible = True
'    Image3.Visible = False
'
'    opt_ChartGraph(0).value = True
'    Call opt_ChartGraph_Click(0)
'End Sub

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


Private Sub opt_ChartGraph_Click(Index As Integer)

    Select Case Index
        Case 0
            Call DrawChart_Time '시간그래프(24시간)
        Case 1
            Call DrawChart_Day '요일그래프(일요일~토요일)
        Case 2
            Call DrawChart_Month '월간그래프
        Case 3
            Call DrawChart_CarGubun '차량구분그래프
        Case 4
            Call DrawChart_Gate '게이트그래프
    End Select
End Sub

Private Sub opt_inout_Click(Index As Integer)

    'opt_inout(0).value = True
    'opt_inout(1).Value =True
    If (Index = 0) Then '입출차내역
        cmd_Button(0).Visible = False
        cmd_Button(3).Visible = False
        lbl_option(1).Visible = False
        cmd_Button(0).Enabled = False
        cmd_Button(3).Enabled = False
        lbl_option(1).Enabled = False
        ListView1.MultiSelect = False
        
        cmd_Button(1).Visible = False
        cmd_Button(1).Enabled = False
    Else    '주차내역
'        cmd_Button(0).Visible = True
'        cmd_Button(3).Visible = True
'        lbl_option(1).Visible = True
'        cmd_Button(0).Enabled = True
'        cmd_Button(3).Enabled = True
'        lbl_option(1).Enabled = True
        ListView1.MultiSelect = True '멀티선택
        
        '무인기 사용
        'If (Glo_ApsYN = "Y") Then
        If (Glo_ApsYN = "Y" Or Glo_PreApsYN = "Y") Then
            cmd_Button(0).Visible = True
            cmd_Button(0).Enabled = True
            cmd_Button(1).Visible = True
            cmd_Button(1).Enabled = True
            cmd_Button(3).Visible = True
            cmd_Button(3).Enabled = True
            lbl_option(1).Visible = True
        Else
            cmd_Button(0).Visible = False
            cmd_Button(0).Enabled = False
            cmd_Button(1).Visible = False
            cmd_Button(1).Enabled = False
            cmd_Button(3).Visible = False
            cmd_Button(3).Enabled = False
            lbl_option(1).Visible = False
        End If
    End If
End Sub

'인쇄
Private Sub SSCommand1_Click()
    Dim i, j As Integer
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String
    
On Error GoTo Err_p
    
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_차량입출차현황" ' & ".xls"
    
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "엑셀파일(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave

    If (CommonDialog1.CancelError = True) Then
    
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)
        
        'Call makeexcel(ListView1, tmpFileName, "차량입출차현황")
        'Call makeexcel(ListView1, tmpFileName, "검색내역")
        Call MakeCSV(ListView1, tmpFileName)
    End If
Exit Sub

Err_p:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select

End Sub

Private Sub SSCommand2_Click()
    Unload Me
    'Me.Hide
End Sub

Public Sub ListView_Draw()
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim bQryResult As Boolean

On Error GoTo Err_p

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " 차량번호      "
    ListView1.ColumnHeaders.Add , , " 인식번호      "
    ListView1.ColumnHeaders.Add , , " 차량모델  "
    ListView1.ColumnHeaders.Add , , " 차량구분  "
    ListView1.ColumnHeaders.Add , , " 이름  "
    ListView1.ColumnHeaders.Add , , " 연락처    "
    ListView1.ColumnHeaders.Add , , " 소속 / 동      "
    ListView1.ColumnHeaders.Add , , " 지급 / 호      "
    ListView1.ColumnHeaders.Add , , " 시작일    "
    ListView1.ColumnHeaders.Add , , " 종료일    "
    ListView1.ColumnHeaders.Add , , " GATE  "
    ListView1.ColumnHeaders.Add , , " IN/OUT    "
    ListView1.ColumnHeaders.Add , , " 처리일시      "
    'ListView1.ColumnHeaders.Add , , " PASS_YN   "
    'ListView1.ColumnHeaders.Add , , " 정기권차량 "
    ListView1.ColumnHeaders.Add , , " 차단기오픈 "
    ListView1.ColumnHeaders.Add , , " 처리결과  "
    ListView1.ColumnHeaders.Add , , " 이미지경로    "
    ListView1.ColumnHeaders.Add , , " IP    "
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
 
 
    Set rs = New ADODB.Recordset
    'rs.Open Glo_JIOSch, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, Glo_JIOSch, False)
    If (bQryResult = False) Then
        Call DataLogger("[FrmInOut]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    
    
    LblRecordCount = rs.RecordCount
    INDEX_NO = 1
    Do While Not (rs.EOF)
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        itmX.SubItems(1) = "" & rs!CAR_NO
        itmX.SubItems(2) = "" & rs!REC_NO
        itmX.SubItems(3) = "" & rs!CAR_MODEL
        itmX.SubItems(4) = "" & rs!CAR_GUBUN
        If rs!DRIVER_NAME <> " " Then
            itmX.SubItems(5) = "" & rs!DRIVER_NAME
        Else
            itmX.SubItems(5) = " "
        End If
        If rs!DRIVER_PHONE <> " " Then
            itmX.SubItems(6) = "" & rs!DRIVER_PHONE
        Else
            itmX.SubItems(6) = " "
        End If
        itmX.SubItems(7) = "" & rs!DRIVER_DEPT
        itmX.SubItems(8) = "" & rs!DRIVER_CLASS
        itmX.SubItems(9) = "" & rs!START_DATE
        itmX.SubItems(10) = "" & rs!END_DATE
'        itmX.SubItems(11) = "" & rs!PASS_GATE
        If (rs!PASS_GATE = 0) Then
            itmX.SubItems(11) = LANE1_Name
        ElseIf (rs!PASS_GATE = 1) Then
            itmX.SubItems(11) = LANE2_Name
        ElseIf (rs!PASS_GATE = 2) Then
            itmX.SubItems(11) = LANE3_Name
        ElseIf (rs!PASS_GATE = 3) Then
            itmX.SubItems(11) = LANE4_Name
        ElseIf (rs!PASS_GATE = 4) Then
            itmX.SubItems(11) = LANE5_Name
        ElseIf (rs!PASS_GATE = 5) Then
            itmX.SubItems(11) = LANE6_Name
        End If
        
        If (rs!PASS_GATE = 6) Then
            itmX.SubItems(11) = LANE1_Name & "(후방)"
        ElseIf (rs!PASS_GATE = 7) Then
            itmX.SubItems(11) = LANE2_Name & "(후방)"
        ElseIf (rs!PASS_GATE = 8) Then
            itmX.SubItems(11) = LANE3_Name & "(후방)"
        ElseIf (rs!PASS_GATE = 9) Then
            itmX.SubItems(11) = LANE4_Name & "(후방)"
        ElseIf (rs!PASS_GATE = 10) Then
            itmX.SubItems(11) = LANE5_Name & "(후방)"
        ElseIf (rs!PASS_GATE = 11) Then
            itmX.SubItems(11) = LANE6_Name & "(후방)"
        End If
        
        itmX.SubItems(12) = "" & rs!PASS_INOUT
        itmX.SubItems(13) = "" & Format(rs!PASS_DATE, "yyyy-mm-dd hh:nn:ss")
        itmX.SubItems(14) = "" & rs!Pass_YN
        itmX.SubItems(15) = "" & rs!PASS_RESULT
        itmX.SubItems(16) = "" & rs!pass_image
        itmX.SubItems(17) = "" & rs!PASS_IP
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    INDEX_NO = 0
    Set rs = Nothing
    
Exit Sub

Err_p:
    Call DataLogger(" [FrmInOut]  " & Err.Description)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    Dim Tmp_File As String
    Dim image_name As String
    Dim ECHO As ICMP_ECHO_REPLY
    Dim RemoteIP As String

On Error Resume Next
    
    lbl_time_now.Caption = "" & Left(ListView1.SelectedItem.SubItems(13), 19)
    lbl_CaNo(0).ForeColor = &HFF00&
    lbl_CaNo(1).ForeColor = &HFF00&
    lbl_CaNo(0).Caption = "" & ListView1.SelectedItem.SubItems(1)
    txt_Update.text = "" & ListView1.SelectedItem.SubItems(1)
    lbl_CaNo(1).Caption = "" & ListView1.SelectedItem.SubItems(2)
    
    If Trim(ListView1.SelectedItem.SubItems(1)) <> Trim(ListView1.SelectedItem.SubItems(2)) Then
        lbl_CaNo(1).ForeColor = vbRed
    End If
    
    RemoteIP = "" & ListView1.SelectedItem.SubItems(17)
    'RemoteIP = Mid(Trim(ListView1.SelectedItem.SubItems(9)), 3, InStr(3, Trim(ListView1.SelectedItem.SubItems(9)), "\", 1) - 3)
    
    'Ping Test...!!
'    Call Ping(RemoteIP, ECHO)
'    If Left$(ECHO.Data, 1) <> Chr$(0) Then
        Tmp_File = Dir(Trim(ListView1.SelectedItem.SubItems(16)))
        If (Tmp_File <> "") Then
            Image3.Picture = LoadPicture(Trim(ListView1.SelectedItem.SubItems(16)))
        Else
            Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")
        End If
'    Else
'        Image3.Picture = Nothing
'        Call DataLogger("[FrmInOut]    Ping Test Failure...!!")
        'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "Ping Test Failure...!!")
'    End If
End Sub


Private Sub SSCommand3_Click()
    'Unload Me
    'FrmGuestLog.Show 1
    
    FrmGuestLog.Show 0
    Call DataLogger("[INOUT Button]    " & "방문객내역 화면 접근")
End Sub


Private Sub DrawChart_Gate()
    Dim nX As Long
    Dim sSDT As String
    Dim sEDT As String
    
    Dim i As Long
    Dim sGateName() As String
    Dim iSum() As Long
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_p

    sSDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    qry = "SELECT PASS_GATE, COUNT(PASS_GATE) AS Pass_Count From tb_inout WHERE (PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "') GROUP BY PASS_GATE "

    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        
        i = 1
        ReDim sGateName(rs.RecordCount) As String
        ReDim iSum(rs.RecordCount) As Long
        
        Do While Not (rs.EOF)
            
            
            With Combo4
            .AddItem "전체"
            If (Glo_Screen_No >= 1) Then
                .AddItem LANE1_Name
                Back_Start_Index = 4
            End If
            If (Glo_Screen_No >= 2) Then
                .AddItem LANE2_Name
                Back_Start_Index = 4
            End If
            If (Glo_Screen_No >= 4) Then
                .AddItem LANE3_Name
                .AddItem LANE4_Name
                Back_Start_Index = 4
            End If
            If (Glo_Screen_No >= 6) Then
                .AddItem LANE5_Name
                .AddItem LANE6_Name
            End If
'            .Text = Combo4.List(0)
    End With
    
    
            Select Case rs!PASS_GATE
                Case 0
                    sGateName(i) = LANE1_Name
                Case 1
                    sGateName(i) = LANE2_Name
                Case 2
                    sGateName(i) = LANE3_Name
                Case 3
                    sGateName(i) = LANE4_Name
                Case 4
                    sGateName(i) = LANE5_Name
                Case 5
                    sGateName(i) = LANE6_Name
            End Select
            
            iSum(i) = rs!Pass_Count:
            i = i + 1
            
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    
    MSChart1.RowCount = 1
    MSChart1.ColumnCount = UBound(sGateName)
    MSChart1.RowLabel = "게이트구분 그래프"
    For nX = 1 To UBound(sGateName)
        MSChart1.row = 1
        MSChart1.Column = nX
        MSChart1.ColumnLabel = sGateName(nX)
        MSChart1.Data = iSum(nX)
        
        '차트 배경색
        MSChart1.Plot.Wall.Brush.Style = VtBrushStyleSolid
        MSChart1.Plot.Wall.Brush.FillColor.Set 245, 245, 245
        
        '그래프 색상
        'MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
        
        '그래프 상단 건수표시
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Marker.Visible = False
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        
        
    Next nX
    
    MSChart1.Refresh
    
    Exit Sub
    
Err_p:
    Call DebugLogger("[차량구분조회 그래프] 데이터 오류 : " & Err.Description)
End Sub

Private Sub DrawChart_CarGubun()
    Dim nX As Long
    Dim sSDT As String
    Dim sEDT As String
    
    Dim i As Long
    Dim iIndex As Long
    Dim sPassResult() As String
    Dim iPassCount() As Long
    Dim iMaxPassResult As Long
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_p

    sSDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    qry = "SELECT PASS_RESULT, COUNT(PASS_RESULT) AS PASS_COUNT From tb_inout WHERE (PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "') GROUP BY PASS_RESULT ORDER BY PASS_RESULT "

    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        
        i = 1
        ReDim sPassResult(rs.RecordCount) As String
        ReDim iPassCount(rs.RecordCount) As Long
        
        Do While Not (rs.EOF)
            
            sPassResult(i) = rs!PASS_RESULT:
            iPassCount(i) = rs!Pass_Count:
            i = i + 1
            
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    
    MSChart1.RowCount = 1
    MSChart1.ColumnCount = UBound(sPassResult)
    MSChart1.RowLabel = "차량구분 그래프"
    For nX = 1 To UBound(sPassResult)
        MSChart1.row = 1
        MSChart1.Column = nX
        MSChart1.ColumnLabel = sPassResult(nX)
        MSChart1.Data = iPassCount(nX)
        
        '차트 배경색
        MSChart1.Plot.Wall.Brush.Style = VtBrushStyleSolid
        MSChart1.Plot.Wall.Brush.FillColor.Set 245, 245, 245
        
        '그래프 색상
        'MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
        
        '그래프 상단 건수표시
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Marker.Visible = False
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        
        
    Next nX
    
    MSChart1.Refresh
    
    Exit Sub
    
Err_p:
    Call DebugLogger("[차량구분조회 그래프] 데이터 오류 : " & Err.Description)
End Sub
Private Sub DrawChart_Time()
    Dim nX As Long
    Dim sSDT As String
    Dim sEDT As String
    
    Dim iTimeBand As Long
    Dim iTimeCount(24) As Long
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_p

    sSDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    qry = "SELECT PASS_DATE From tb_inout WHERE PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "'"

    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            
            iTimeBand = Val(Mid(rs!PASS_DATE, 12, 2))
            iTimeCount(iTimeBand) = iTimeCount(iTimeBand) + 1

            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    
    MSChart1.RowCount = 1
    MSChart1.ColumnCount = UBound(iTimeCount)
    MSChart1.RowLabel = "시간 그래프"
    For nX = 0 To UBound(iTimeCount) - 1
        MSChart1.row = 1
        MSChart1.Column = nX + 1
        MSChart1.ColumnLabel = (nX) & "시"
        MSChart1.Data = iTimeCount(nX)
        
        '차트 배경색
        MSChart1.Plot.Wall.Brush.Style = VtBrushStyleSolid
        MSChart1.Plot.Wall.Brush.FillColor.Set 245, 245, 245
        
        '그래프 색상
        'MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
        
        '그래프 상단 건수표시
        MSChart1.Plot.SeriesCollection(nX + 1).DataPoints(-1).Marker.Visible = False
        MSChart1.Plot.SeriesCollection(nX + 1).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        
        
    Next nX
    
    MSChart1.Refresh
    
    Exit Sub
    
Err_p:
    Call DebugLogger("[시간그래프] 데이터 오류 : " & Err.Description)
End Sub

Private Sub DrawChart_Day()

    Dim nX As Long
    Dim sSDT As String
    Dim sEDT As String
    Dim sWeekday As String
    Dim iDayCount(7) As Long
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_p

    sSDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    qry = "SELECT PASS_DATE From tb_inout WHERE PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "'"

    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            
            sWeekday = Format(Left(rs!PASS_DATE, 10), "dddd")
            Select Case sWeekday
                Case "Sunday"
                    iDayCount(1) = iDayCount(1) + 1
                Case "Monday"
                    iDayCount(2) = iDayCount(2) + 1
                Case "Tuesday"
                    iDayCount(3) = iDayCount(3) + 1
                Case "Wednesday"
                    iDayCount(4) = iDayCount(4) + 1
                Case "Thursday"
                    iDayCount(5) = iDayCount(5) + 1
                Case "Friday"
                    iDayCount(6) = iDayCount(6) + 1
                Case "Saturday"
                    iDayCount(7) = iDayCount(7) + 1
            End Select
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    
    MSChart1.RowCount = 1
    MSChart1.ColumnCount = 7
    MSChart1.RowLabel = "요일 그래프"
    For nX = 1 To 7
        MSChart1.row = 1
        MSChart1.Column = nX
        If (nX = 1) Then MSChart1.ColumnLabel = "일" Else
        If (nX = 2) Then MSChart1.ColumnLabel = "월" Else
        If (nX = 3) Then MSChart1.ColumnLabel = "화" Else
        If (nX = 4) Then MSChart1.ColumnLabel = "수" Else
        If (nX = 5) Then MSChart1.ColumnLabel = "목" Else
        If (nX = 6) Then MSChart1.ColumnLabel = "금" Else
        If (nX = 7) Then MSChart1.ColumnLabel = "토"
        MSChart1.Data = iDayCount(nX)
        
        '차트 배경색
        MSChart1.Plot.Wall.Brush.Style = VtBrushStyleSolid
        MSChart1.Plot.Wall.Brush.FillColor.Set 245, 245, 245
        
        '그래프 색상
        'MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
        
        '그래프 상단 건수표시
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Marker.Visible = False
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        
        
    Next nX
    
    MSChart1.Refresh
    
    Exit Sub
    
Err_p:
    Call DebugLogger("[요일그래프] 데이터 오류 : " & Err.Description)
    
End Sub

Private Sub DrawChart_Week()

    Dim nX As Long
    Dim nRow As Long
    Dim i, j As Long
    Dim sSDT As String
    Dim sEDT As String
    Dim iMonthCount As Long '월 개수
    Dim iYearStart As Long '시작 년
    Dim iMonthStart As Long '시작 월
    Dim sName() As String '년월 저장
    Dim iCount() As Long '년월별 카운트
    Dim iWeek As Long
    Dim WeekNo As Long
    
    Dim iStartDT As String
    Dim iEndDT As String
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_p

    iMonthStart = 0
    iMonthCount = 0
    iWeek = 0
    
    sSDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    qry = "SELECT min(PASS_DATE) as StartDT, max(PASS_DATE) as EndDT  From tb_inout WHERE PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "' "

    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        iStartDT = Format(Mid(rs!StartDT, 1, 7)) 'yyyy-mm
        iEndDT = Format(Mid(rs!EndDT, 1, 7)) ''yyyy-mm
        iMonthCount = DateDiff("m", iStartDT, iEndDT) + 1 '전체 개월수 저장
    End If
    Set rs = Nothing
    
    If (iMonthCount > 0) Then
        ReDim sName(iMonthCount * 5) As String
        ReDim iCount(iMonthCount * 5) As Long
    End If
    
    Dim sMon As String
    sMon = iStartDT
    For i = 1 To iMonthCount
        For j = 1 To 5
            sName(((i - 1) * 5) + j) = Format(sMon, "yyyy-mm") & "-" & j
        Next j
        sMon = DateAdd("m", 1, iStartDT)
    Next i
    
    
    
    
    Dim iIndex As Long
    qry = "SELECT PASS_DATE From tb_inout WHERE PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "' ORDER BY PASS_DATE"
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        Do While Not (rs.EOF)

'            iIndex = DateDiff("m", iStartDT, Format(Mid(rs!PASS_DATE, 1, 7))) + 1 '전체 개월수 저장
'            iCount(iIndex) = iCount(iIndex) + 1
'            sName(iIndex) = Mid(rs!PASS_DATE, 1, 7) ' yyyy-mm

            iIndex = DateDiff("m", iStartDT, Format(Mid(rs!PASS_DATE, 1, 7))) + 1 '전체 개월수 저장
            WeekNo = DateDiff("ww", Format(Mid(rs!PASS_DATE, 1, 10), "yyyy-01-01"), Mid(rs!PASS_DATE, 1, 10))
            iCount(iIndex + WeekNo) = iCount(iIndex + WeekNo) + 1
            

            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    
    Dim ArrValue(1 To 5, 1 To 3)
    For i = 1 To 5
        ArrValue(i, 1) = "Label" & i
        ArrValue(i, 2) = i
        ArrValue(i, 3) = i * 2
    Next i
    MSChart1.ChartData = ArrValue
    MSChart1.Refresh
    
'    MSChart1.RowCount = iMonthCount
'    MSChart1.ColumnCount = 5
'    MSChart1.RowLabel = "주간그래프 조회"
'    For nRow = 1 To iMonthCount
'        For nX = 1 To 5
'            MSChart1.row = nRow
'            MSChart1.Column = nX
'            MSChart1.ColumnLabel = sName(((nRow - 1) * 5) + nX)
'
'            MSChart1.Data = iCount(((nRow - 1) * 5) + nX)
'
'            '차트 배경색
'            MSChart1.Plot.Wall.Brush.Style = VtBrushStyleSolid
'            MSChart1.Plot.Wall.Brush.FillColor.Set 245, 245, 245
'
'            '그래프 상단 건수표시
'            MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Marker.Visible = False
'            MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
'
'        Next nX
'    Next nRow
'
'    MSChart1.Refresh
    
    Exit Sub
    
Err_p:
    Call DebugLogger("[주간그래프] 데이터 오류 : " & Err.Description)
    
End Sub


Private Sub DrawChart_Month()

    Dim nX As Long
    Dim sSDT As String
    Dim sEDT As String
    Dim iMonth As Long
    Dim iMonthCount(12) As Long
    
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
On Error GoTo Err_p

    sSDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    qry = "SELECT PASS_DATE From tb_inout WHERE PASS_DATE >= '" & sSDT & "' AND PASS_DATE <= '" & sEDT & "'"

    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            
            iMonth = Val(Format(Mid(rs!PASS_DATE, 6, 2)))
            iMonthCount(iMonth) = iMonthCount(iMonth) + 1
            
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    MSChart1.RowCount = 1
    MSChart1.ColumnCount = 12
    MSChart1.RowLabel = "월간 그래프"
    For nX = 1 To 12
        MSChart1.row = 1
        MSChart1.Column = nX
        MSChart1.ColumnLabel = nX & "월"
        MSChart1.Data = iMonthCount(nX)
        
        '차트 배경색
        MSChart1.Plot.Wall.Brush.Style = VtBrushStyleSolid
        MSChart1.Plot.Wall.Brush.FillColor.Set 245, 245, 245
        
        'MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        'MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).DataPointLabel.VtFont.Style = VtFontStyleBold
        
        '그래프 상단 건수표시
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Marker.Visible = False
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        MSChart1.Plot.SeriesCollection(nX).DataPoints(-1).Marker.FillColor.Set 0, 255, 0
        
    Next nX
    
    MSChart1.Refresh
    
    Exit Sub

Err_p:
    Call DebugLogger("[월간그래프] 데이터 오류 : " & Err.Description)
End Sub

Private Sub DeleteInOut(DelteItem As String)
    Dim sCarNo, sPassDate, sDateGab As String
    Dim StartDate, StartTime, EndDate, EndTime As String
    
    Dim iListCount As Integer
    Dim i As Long
    
    '기간삭제
    If (DelteItem = "DELETE_RANGE") Then
            StartDate = Format(DTPicker1, "yyyy-mm-dd")
            StartTime = Format(DTPicker3, "hh:nn:ss")
            EndDate = Format(DTPicker2, "yyyy-mm-dd")
            EndTime = Format(DTPicker4, "hh:nn:ss") & ".999"
            
            sDateGab = Format(DTPicker1, "yyyy-mm-dd") & " ~ " & Format(DTPicker2, "yyyy-mm-dd")
            MBox.Label3.FontSize = 16
            MBox.Label3.Caption = sDateGab
            MBox.Label2.Caption = "주차내역 삭제"
            MBox.Label1.Caption = "위 기간의 모든 일반권 주차내역을 " & vbCrLf & "삭제하시겠습니까?" & vbCrLf & "삭제 후에는 복구안됩니다"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
                '삭제쿼리
                adoConn.Execute "Delete From tb_now Where (PASS_DATE >='" & StartDate & " 00:00:00.000' AND PASS_DATE <='" & EndDate & " 23:59:59.999') AND (CAR_GUBUN = '' OR CAR_GUBUN IS NULL)"
                Call DataLogger(" " & sDateGab & " 일반권 주차내역을 삭제!!")
            End If
            Call SSCommand4_Click
    
    '선택삭제
    Else
            If (Len(txt_Update.text) = 0) Then
                Msg_Box.Label2.Caption = "일반권 주차정보 삭제"
                Msg_Box.Label1.Caption = "삭제할 일반권 주차차량을 " & vbCrLf & "선택하세요."
                Msg_Box.Show 1
                Exit Sub
            End If
            
            For i = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(i).Selected = True Then
                        iListCount = iListCount + 1
                    End If
            Next i
            If (iListCount = 1) Then
                MBox.Label3.Caption = txt_Update.text
            ElseIf (iListCount >= 2) Then
                MBox.Label3.Caption = txt_Update.text & " 외 " & iListCount - 1 & "건"
            End If
            MBox.Label3.FontSize = 20
            MBox.Label2.Caption = "주차내역 삭제"
            MBox.Label1.Caption = "위 차량의 주차내역을 " & vbCrLf & "삭제하시겠습니까?" & vbCrLf & "삭제 후에는 복구안됩니다"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
                For i = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(i).Selected = True Then
    
                        sCarNo = ListView1.ListItems(i).SubItems(2)
                        sPassDate = ListView1.ListItems(i).SubItems(13)
    
                        adoConn.Execute "Delete From tb_now Where CAR_NO= '" & sCarNo & "' and PASS_DATE = '" & sPassDate & "'"
    
                        'List1.AddItem sPassDate & "  " & ScarNo & " 일반차량의 입차내역을 삭제", 0
                        Call DataLogger(" " & sCarNo & "  " & sPassDate & " 일반차량의 주차내역을 삭제")
                        
                    End If
                Next i
            End If
            Call SSCommand4_Click
    End If
End Sub

Private Sub SSCommand4_Click()
    Frame1.Visible = False
    Image3.Visible = True
    
    Call Command1_Click
End Sub

Private Sub SSCommand5_Click()
    Frame1.Visible = True
    Image3.Visible = False
    
    opt_ChartGraph(0).value = True
    Call opt_ChartGraph_Click(0)
End Sub

Private Sub SSCommand6_Click()
    Call cmd_Update_Click
    Call Command1_Click
End Sub
