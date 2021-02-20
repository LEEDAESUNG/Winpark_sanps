VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRegHistory 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   12315
   ClientLeft      =   4695
   ClientTop       =   2100
   ClientWidth     =   17145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   821
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   1143
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_ExSch 
      Caption         =   "조 회"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   23340
      TabIndex        =   34
      Top             =   1950
      Width           =   1125
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
      Left            =   23835
      TabIndex        =   27
      Top             =   1305
      Width           =   525
   End
   Begin VB.CommandButton cmd_Update 
      Caption         =   "변 경"
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
      Left            =   23280
      TabIndex        =   26
      Top             =   2970
      Width           =   945
   End
   Begin VB.TextBox txt_Update 
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
      Left            =   20190
      MaxLength       =   10
      TabIndex        =   25
      Top             =   2985
      Width           =   2745
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
      ItemData        =   "FrmRegHistory.frx":0000
      Left            =   11745
      List            =   "FrmRegHistory.frx":0010
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   840
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
      ItemData        =   "FrmRegHistory.frx":0038
      Left            =   20085
      List            =   "FrmRegHistory.frx":0042
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   615
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
      ItemData        =   "FrmRegHistory.frx":0061
      Left            =   20085
      List            =   "FrmRegHistory.frx":0086
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   180
      Width           =   1950
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
      Height          =   540
      IMEMode         =   10  '한글 
      Left            =   11745
      TabIndex        =   0
      Top             =   2925
      Width           =   2835
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   11745
      TabIndex        =   4
      Top             =   1275
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
      Format          =   159973376
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   14460
      TabIndex        =   5
      Top             =   1275
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
      Format          =   159973376
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   11745
      TabIndex        =   6
      Top             =   1740
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
      Format          =   159973378
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   14460
      TabIndex        =   7
      Top             =   1755
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
      Format          =   159973378
      CurrentDate     =   36927
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   14475
      TabIndex        =   8
      Top             =   3960
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
         TabIndex        =   9
         Top             =   60
         Width           =   1275
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   7110
      Left            =   165
      TabIndex        =   10
      Top             =   4440
      Width           =   16860
      _ExtentX        =   29739
      _ExtentY        =   12541
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
      Left            =   18705
      TabIndex        =   28
      Top             =   1305
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
      Format          =   159973376
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   345
      Left            =   21240
      TabIndex        =   29
      Top             =   1305
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
      Format          =   159973376
      CurrentDate     =   36927
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12675
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand cmd_Excel 
      Height          =   540
      Left            =   14190
      TabIndex        =   38
      Top             =   60
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Picture         =   "FrmRegHistory.frx":010C
   End
   Begin Threed.SSCommand cmd_Exit 
      Cancel          =   -1  'True
      Height          =   540
      Left            =   15690
      TabIndex        =   39
      Top             =   60
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
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
      Picture         =   "FrmRegHistory.frx":045D
   End
   Begin Threed.SSCommand cmd_Search 
      Height          =   600
      Left            =   14700
      TabIndex        =   40
      Top             =   2895
      Width           =   1740
      _Version        =   65536
      _ExtentX        =   3069
      _ExtentY        =   1058
      _StockProps     =   78
      Caption         =   "검색"
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
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmRegHistory.frx":07AE
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회구분"
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
      Left            =   10545
      TabIndex        =   37
      Top             =   855
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
      Left            =   10545
      TabIndex        =   36
      Top             =   1770
      Width           =   900
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
      Height          =   285
      Index           =   0
      Left            =   10545
      TabIndex        =   35
      Top             =   1290
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   1244
      X2              =   1675
      Y1              =   182
      Y2              =   182
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "입차한 일반차량 검색"
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
      Left            =   20790
      TabIndex        =   33
      Top             =   2055
      Width           =   2145
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
      Left            =   24465
      TabIndex        =   32
      Top             =   1335
      Width           =   735
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
      Left            =   23280
      TabIndex        =   31
      Top             =   1335
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   14
      Left            =   20745
      TabIndex        =   30
      Top             =   1335
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   1
      X1              =   1244
      X2              =   1675
      Y1              =   273
      Y2              =   273
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
      Left            =   5895
      TabIndex        =   24
      Top             =   1305
      Width           =   3405
   End
   Begin VB.Label lbl_CaNo 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Index           =   1
      Left            =   20220
      TabIndex        =   23
      Top             =   3540
      Width           =   2730
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
      Left            =   18765
      TabIndex        =   22
      Top             =   3525
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
      TabIndex        =   21
      Top             =   690
      Width           =   2925
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '투명
      Caption         =   "정기권 이력 조회"
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
      Left            =   255
      TabIndex        =   20
      Top             =   195
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
      Left            =   18765
      TabIndex        =   19
      Top             =   3030
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
      Left            =   10395
      TabIndex        =   18
      Top             =   2985
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
      Left            =   16545
      TabIndex        =   17
      Top             =   1320
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
      Left            =   16545
      TabIndex        =   16
      Top             =   1785
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
      Left            =   13830
      TabIndex        =   15
      Top             =   1335
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
      Left            =   13830
      TabIndex        =   14
      Top             =   1785
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
      Left            =   10545
      TabIndex        =   13
      Top             =   1320
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
      Left            =   18735
      TabIndex        =   12
      Top             =   645
      Width           =   900
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
      Left            =   18735
      TabIndex        =   11
      Top             =   195
      Width           =   900
   End
   Begin VB.Image Image3 
      Height          =   3060
      Left            =   165
      Picture         =   "FrmRegHistory.frx":0AFF
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   4110
   End
End
Attribute VB_Name = "FrmRegHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim excel_sql_str As String

Private Sub cmd_Exit_Click()
    Unload Me
    'Me.Hide
End Sub

Private Sub cmd_Update_Click()

'    If (txt_Update.Text <> "") And (lbl_time_now.Caption <> "") Then
'        adoConn.Execute "Update tb_inout Set CAR_NO = '" & Trim(txt_Update.Text) & "' Where REG_DATE = '" & lbl_time_now & "'"
'    Else
'
'    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        
    If (KeyAscii = 13) Then
        If (Len(txt_CarNo) <> 0) Then
            Call cmd_Search_Click
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

Private Sub Form_Load()
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    'Left = 0
    'Top = 0
    lbl_Update(0).Caption = "Update : " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    DTPicker1.value = Now
    DTPicker2.value = Now
    DTPicker3.value = Format("00:00:00")
    DTPicker4.value = Format("23:59:59")
    
    DTPicker5.value = Now
    DTPicker6.value = Now
    txt_Count.text = 2
    
    Combo1.ListIndex = 0
    'Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
    Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")
    
    '오늘날짜 데이터만
    Glo_JIOSch = "SELECT * FROM tb_reg_log WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY REG_DATE"
    
    Normal_Search_F = True
    Call ListView_Draw
Exit Sub

Err_P:
    MsgBox "데이터 베이스 연결실패" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "네트웍 관리자에게 문의 바랍니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "데이터 베이스 연결전에는 자료검색 기능을 수행할수 없습니다.", vbCritical
End Sub
'검색 버튼
Private Sub cmd_Search_Click()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim Sort_Order As String
    Dim sql_str As String
    
    Normal_Search_F = True
    Me.MousePointer = 11
    cmd_ExSch.Enabled = False
    Sort_Order = Combo3.List(Combo3.ListIndex)
    

    sql_str = "SELECT * FROM tb_reg_log WHERE (REG_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & "') AND (REG_DATE <='" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & "') "

    Select Case Combo4.ListIndex
        Case 0
        Case 1
            sql_str = sql_str & " AND (ACTION_LOG like '%등록%') "
        Case 2
            sql_str = sql_str & " AND (ACTION_LOG like '%수정%') "
        Case 3
            sql_str = sql_str & " AND (ACTION_LOG like '%삭제%') "
    End Select
    If (txt_CarNo.text = "") Then
    Else
            sql_str = sql_str & " AND (CAR_NO LIKE '%" & Trim(txt_CarNo.text) & "' )"

    End If
    
    
    
    Glo_JIOSch = sql_str & " ORDER BY REG_DATE"

    
    Call ListView_Draw
    cmd_ExSch.Enabled = True
    Me.MousePointer = 0
End Sub



Private Sub cmd_Excel_Click()
    Dim i, j As Integer
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String
    
On Error GoTo Err_P
    
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_정기차량이력조회"

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
    Case Else
    End Select
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub

Public Sub ListView_Draw()
Dim Column_to_size As Integer
Dim rs As Recordset
Dim qry As String
Dim itmX As ListItem
Dim INDEX_NO As Long
Dim bQryResult As Boolean


With Me
    Call ListViewExtended(.ListView1)
    .ListView1.View = lvwReport
    .ListView1.ListItems.Clear
    .ListView1.ColumnHeaders.Clear
    .ListView1.ColumnHeaders.Add , , " No  "
    .ListView1.ColumnHeaders.Add , , " 아이디        "
    .ListView1.ColumnHeaders.Add , , " 차량번호        "
    .ListView1.ColumnHeaders.Add , , " 이력내용                "
    .ListView1.ColumnHeaders.Add , , " 차량번호(변경후) "
    .ListView1.ColumnHeaders.Add , , " 차량모델     "
    .ListView1.ColumnHeaders.Add , , " 차량구분   "
    .ListView1.ColumnHeaders.Add , , " 월정요금   "
    .ListView1.ColumnHeaders.Add , , " 이    름     "
    .ListView1.ColumnHeaders.Add , , " 연 락 처              "
    If (Glo_User_Type = "구분1/구분2") Then
        ListView1.ColumnHeaders.Add , , " 소    속    "
        ListView1.ColumnHeaders.Add , , " 직    급    "
    Else
        ListView1.ColumnHeaders.Add , , " 거주  동    "
        ListView1.ColumnHeaders.Add , , " 거주  호    "
    End If
    .ListView1.ColumnHeaders.Add , , " 시 작 일      "
    .ListView1.ColumnHeaders.Add , , " 종 료 일      "
    .ListView1.ColumnHeaders.Add , , " 수 정 일                         "
    .ListView1.ColumnHeaders.Add , , " 결 제 일   "
    .ListView1.ColumnHeaders.Add , , " 세대통보 "
    .ListView1.ColumnHeaders.Add , , " 등록 "
    .ListView1.ColumnHeaders.Add , , " 기타 "
    .ListView1.ColumnHeaders.Add , , " 처 리 일                         "
    .ListView1.ColumnHeaders.Add , , ""
    
    
    For Column_to_size = 0 To .ListView1.ColumnHeaders.Count - 2
         SendMessage .ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
    
    
    If (Glo_JIOSch <> "") Then
    
        INDEX_NO = 1
        Set rs = New ADODB.Recordset
        'rs.Open Glo_JIOSch, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, Glo_JIOSch, False)
        If (bQryResult = False) Then
            Call DataLogger("[FrmRegHistory]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        
        LblRecordCount = rs.RecordCount
        Do While Not (rs.EOF)
            Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
            itmX.SubItems(1) = "" & rs!ACTION_ID
            itmX.SubItems(2) = "" & rs!CAR_NO
            itmX.SubItems(3) = "" & rs!ACTION_LOG
            itmX.SubItems(4) = "" & rs!AFTER_CAR_NO
            itmX.SubItems(5) = "" & rs!CAR_MODEL
            itmX.SubItems(6) = "" & rs!CAR_GUBUN
            itmX.SubItems(7) = "" & rs!CAR_FEE
            itmX.SubItems(8) = "" & rs!DRIVER_NAME
            itmX.SubItems(9) = "" & rs!DRIVER_PHONE
            itmX.SubItems(10) = "" & rs!DRIVER_DEPT
            itmX.SubItems(11) = "" & rs!DRIVER_CLASS
            itmX.SubItems(12) = "" & Format(rs!START_DATE, "yyyy-mm-dd")
            itmX.SubItems(13) = "" & Format(rs!END_DATE, "yyyy-mm-dd")
            itmX.SubItems(14) = "" & rs!Update_date
            itmX.SubItems(15) = "" & rs!FEE_DATE
            itmX.SubItems(16) = "" & rs!DAY_ROTATION_YN
            itmX.SubItems(17) = "" & rs!REG_PART
            itmX.SubItems(18) = "" & rs!ETC
            itmX.SubItems(19) = "" & Format(rs!LOG_DATE, "yyyy-mm-dd hh:nn:ss")
            
            
            rs.MoveNext
            INDEX_NO = INDEX_NO + 1
        Loop
        Set rs = Nothing
    
    End If
    
End With
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
'    Dim Tmp_File As String
'    Dim image_name As String
'    Dim ECHO As ICMP_ECHO_REPLY
'    Dim RemoteIP As String
'
'On Error Resume Next
'    If (Normal_Search_F = True) Then
'        lbl_time_now.Caption = "" & ListView1.SelectedItem.SubItems(13)
'        lbl_CaNo(0).ForeColor = &HFF00&
'        lbl_CaNo(1).ForeColor = &HFF00&
'        lbl_CaNo(0).Caption = "" & ListView1.SelectedItem.SubItems(1)
'        txt_Update.Text = "" & ListView1.SelectedItem.SubItems(1)
'        lbl_CaNo(1).Caption = "" & ListView1.SelectedItem.SubItems(2)
        
'        If Trim(ListView1.SelectedItem.SubItems(1)) <> Trim(ListView1.SelectedItem.SubItems(2)) Then
'            lbl_CaNo(1).ForeColor = vbRed
'        End If
'
'        RemoteIP = "" & ListView1.SelectedItem.SubItems(17)
'        'RemoteIP = Mid(Trim(ListView1.SelectedItem.SubItems(9)), 3, InStr(3, Trim(ListView1.SelectedItem.SubItems(9)), "\", 1) - 3)
'
'        'Ping Test...!!
'        Call Ping(RemoteIP, ECHO)
'        If Left$(ECHO.Data, 1) <> Chr$(0) Then
'            Tmp_File = Dir(Trim(ListView1.SelectedItem.SubItems(16)))
'            If (Tmp_File <> "") Then
'                Image3.Picture = LoadPicture(Trim(ListView1.SelectedItem.SubItems(16)))
'            Else
'                Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")
'            End If
'        Else
'            Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")
'            Call DataLogger("[FrmRegHistory]    Ping Test Failure...!!")
'            'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "Ping Test Failure...!!")
'        End If
'    Else
'        txt_CarNo = Trim(ListView1.SelectedItem.SubItems(1))
'    End If

End Sub


'Custum Search
Private Sub cmd_ExSch_Click()
'Dim i As Integer
'Dim Cnt As Integer
'Dim Current_Date As String
'Dim TmpPath As String
'Dim Tmp_File As String
'Dim InsSQL As String
'Dim Now_Flag As Boolean
'Dim sql_str As String
'Dim Sort_Order As String
'
'Normal_Search_F = False
'
'Me.MousePointer = 11
'
'Glo_SQL_SEARCH = ""
'
'If IsNumeric(txt_Count.Text) Then
'    If (txt_Count.Text = 0) Then
'        MsgBox " 올바른 숫자를 입력하세요...!! "
'        Me.MousePointer = 0
'        Exit Sub
'    Else
'
'    End If
'Else
'    MsgBox " 숫자를 입력하세요...!! "
'    Me.MousePointer = 0
'    Exit Sub
'End If
'
''쿼리 구성
'sql_str = "SELECT CAR_NO, IN_COUNT From (SELECT tb_inout.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT From tb_inout Where tb_inout.PASS_YN = 'N' AND tb_inout.PASS_INOUT = 'IN' AND tb_inout.REG_DATE >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' AND tb_inout.REG_DATE <= '" & Format(DTPicker6, "yyyy-mm-dd") & " 23:59:59' Group By tb_inout.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & Val(txt_Count.Text) & ""
''Debug.Print sql_str
'
'Glo_SQL_SEARCH = sql_str
'
'Call ListView_Draw_ParkOne
'
''SSPanel3(1).Caption = ""
'Me.MousePointer = 0
'On Error Resume Next

End Sub


Public Sub ListView_Draw_ParkOne()
'Dim Column_to_size As Integer
'Dim rs As Recordset
'Dim Qry As String
'Dim itmX As ListItem
'Dim INDEX_NO As Long
'
'    Call ListViewExtended(ListView1)
'    ListView1.View = lvwReport
'    ListView1.ListItems.Clear
'    ListView1.ColumnHeaders.Clear
'    ListView1.ColumnHeaders.Add , , " No "
'    ListView1.ColumnHeaders.Add , , " CAR_NO         "
'    ListView1.ColumnHeaders.Add , , " IN_Count       "
'
'    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
'         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
'    Next
'
'    Set rs = New ADODB.Recordset
'    'Debug.Print Glo_SQL_SEARCH
'    rs.Open Glo_SQL_SEARCH, adoConn
'    LblRecordCount = rs.RecordCount
'
'    INDEX_NO = 1
'
'    Do While Not (rs.EOF)
'        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
'        itmX.SubItems(1) = "" & rs!CAR_NO
'        itmX.SubItems(2) = "" & rs!IN_COUNT
'        rs.MoveNext
'        INDEX_NO = INDEX_NO + 1
'    Loop
'    INDEX_NO = 0
'    Set rs = Nothing

End Sub
