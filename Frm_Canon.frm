VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_Canon 
   Caption         =   " tb_canon"
   ClientHeight    =   13080
   ClientLeft      =   7095
   ClientTop       =   2010
   ClientWidth     =   15780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   13080
   ScaleWidth      =   15780
   Begin VB.CommandButton cmd_PhotoUpdate 
      Caption         =   "사진 등록"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   13200
      TabIndex        =   18
      Top             =   8055
      Width           =   1320
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "찾아보기"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11715
      TabIndex        =   17
      Top             =   8055
      Width           =   1320
   End
   Begin VB.TextBox txt_PhotoPath 
      Appearance      =   0  '평면
      BackColor       =   &H8000000A&
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9015
      TabIndex        =   60
      Top             =   7605
      Width           =   6405
   End
   Begin VB.TextBox txt_tmpCarNo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   8040
      TabIndex        =   26
      Top             =   210
      Width           =   2115
   End
   Begin VB.ComboBox cmb_Sch 
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
      ItemData        =   "Frm_Canon.frx":0000
      Left            =   6375
      List            =   "Frm_Canon.frx":000A
      TabIndex        =   25
      Text            =   "차량번호"
      Top             =   210
      Width           =   1575
   End
   Begin VB.CommandButton cmd_Month 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1개월 연장"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19080
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   39
      Top             =   1275
      Width           =   1305
   End
   Begin VB.ComboBox cmb_Gubun 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "Frm_Canon.frx":001E
      Left            =   1830
      List            =   "Frm_Canon.frx":0020
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   10365
      Width           =   2355
   End
   Begin VB.TextBox txt_Dong 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   17685
      TabIndex        =   38
      Top             =   2415
      Width           =   2325
   End
   Begin VB.TextBox txt_CarNo 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      TabIndex        =   0
      Top             =   6630
      Width           =   2325
   End
   Begin VB.TextBox txt_Object 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   9015
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   6195
      Width           =   6405
   End
   Begin VB.TextBox txt_Ho 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   5
      Top             =   8940
      Width           =   2325
   End
   Begin VB.TextBox txt_Phone 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      TabIndex        =   2
      Top             =   7560
      Width           =   2325
   End
   Begin VB.TextBox txt_Name 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      TabIndex        =   1
      Top             =   7095
      Width           =   2325
   End
   Begin VB.ComboBox cmb_Search 
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
      Left            =   10905
      TabIndex        =   23
      Text            =   "검색구분"
      Top             =   1050
      Width           =   2715
   End
   Begin VB.TextBox txt_CarModel 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      TabIndex        =   3
      Top             =   8040
      Width           =   2325
   End
   Begin VB.TextBox txt_Num 
      Appearance      =   0  '평면
      BackColor       =   &H8000000A&
      BorderStyle     =   0  '없음
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1830
      TabIndex        =   37
      Top             =   6150
      Width           =   2865
   End
   Begin VB.Frame frm_Rotation 
      Appearance      =   0  '평면
      BackColor       =   &H00404040&
      Caption         =   " 부제 설정 "
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
      Height          =   885
      Left            =   17700
      TabIndex        =   32
      Top             =   7200
      Width           =   7185
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "사용안함"
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
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   36
         Top             =   360
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "2 부제"
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
         Height          =   345
         Index           =   1
         Left            =   2250
         TabIndex        =   35
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "5 부제"
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
         Height          =   345
         Index           =   2
         Left            =   3900
         TabIndex        =   34
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "10 부제"
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
         Height          =   345
         Index           =   3
         Left            =   5550
         TabIndex        =   33
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame frm_Week 
      Caption         =   " 요일 설정 "
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
      Height          =   885
      Left            =   525
      TabIndex        =   31
      Top             =   10860
      Width           =   6405
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "월"
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
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "화"
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
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   10
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "수"
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
         Height          =   315
         Index           =   2
         Left            =   2085
         TabIndex        =   11
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "목"
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
         Height          =   315
         Index           =   3
         Left            =   2925
         TabIndex        =   12
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "금"
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
         Height          =   315
         Index           =   4
         Left            =   3765
         TabIndex        =   13
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "토"
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
         Height          =   315
         Index           =   5
         Left            =   4590
         TabIndex        =   14
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H8000000A&
         Caption         =   "일"
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
         Height          =   315
         Index           =   6
         Left            =   5430
         TabIndex        =   15
         Top             =   390
         Value           =   1  '확인
         Width           =   615
      End
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
      ForeColor       =   &H00C0C0C0&
      Height          =   1185
      Left            =   30
      TabIndex        =   30
      Top             =   11910
      Width           =   15735
   End
   Begin VB.ComboBox cmb_Rotation 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "Frm_Canon.frx":0022
      Left            =   19620
      List            =   "Frm_Canon.frx":002C
      Style           =   2  '드롭다운 목록
      TabIndex        =   29
      Top             =   4275
      Width           =   2325
   End
   Begin VB.ComboBox cmb_Dong 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Frm_Canon.frx":003E
      Left            =   1845
      List            =   "Frm_Canon.frx":0040
      TabIndex        =   4
      Text            =   "cmb_Dong"
      Top             =   8490
      Width           =   2340
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   4095
      Left            =   30
      TabIndex        =   40
      Top             =   1710
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   7223
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
   Begin MSMask.MaskEdBox MaskEdBox_Start 
      Height          =   375
      Left            =   1845
      TabIndex        =   6
      Top             =   9420
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   0
      Left            =   13890
      TabIndex        =   28
      Top             =   150
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "종 료"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":0042
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   2
      Left            =   14160
      TabIndex        =   21
      Top             =   11190
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "삭 제"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":0393
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   4
      Left            =   12750
      TabIndex        =   20
      Top             =   11190
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "수 정"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":06E4
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   1
      Left            =   11325
      TabIndex        =   19
      Top             =   11190
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "등 록"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":0A35
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   3
      Left            =   9900
      TabIndex        =   22
      Top             =   11190
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "초기화"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":0D86
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   5
      Left            =   12600
      TabIndex        =   27
      Top             =   150
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "Excel"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":10D7
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   495
      Index           =   6
      Left            =   13875
      TabIndex        =   24
      Top             =   990
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "검 색"
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
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Canon.frx":1428
   End
   Begin MSMask.MaskEdBox MaskEdBox_End 
      Height          =   375
      Left            =   1845
      TabIndex        =   7
      Top             =   9900
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox_Fee 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """\""#,##0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   18420
      TabIndex        =   41
      Top             =   675
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0"
      PromptChar      =   "_"
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   7
      Left            =   19260
      TabIndex        =   42
      Top             =   5565
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "결 제"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin VB.Image ImgDriver 
      Height          =   3000
      Left            =   9030
      Picture         =   "Frm_Canon.frx":1779
      Stretch         =   -1  'True
      Top             =   7995
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "사     진"
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
      Height          =   375
      Left            =   7965
      TabIndex        =   59
      Top             =   8070
      Width           =   960
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "구       분"
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
      Left            =   720
      TabIndex        =   58
      Top             =   10410
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "요     금"
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
      Height          =   465
      Left            =   17460
      TabIndex        =   57
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lbl_dept 
      BackStyle       =   0  '투명
      Caption         =   "소       속"
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
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   56
      Top             =   8535
      Width           =   960
   End
   Begin VB.Label lbl_clas 
      BackStyle       =   0  '투명
      Caption         =   "차량모델"
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
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   55
      Top             =   8070
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "등록건수 :"
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
      Height          =   255
      Index           =   0
      Left            =   495
      TabIndex        =   54
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label lbl_COUNT 
      BackStyle       =   0  '투명
      Caption         =   "0000"
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
      Height          =   375
      Left            =   1785
      TabIndex        =   53
      Top             =   1185
      Width           =   1425
   End
   Begin VB.Label lbl_Phone 
      BackStyle       =   0  '투명
      Caption         =   "전화번호"
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
      Height          =   375
      Left            =   720
      TabIndex        =   52
      Top             =   7590
      Width           =   1020
   End
   Begin VB.Label lbl_StartDate 
      BackStyle       =   0  '투명
      Caption         =   "시  작  일"
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
      Height          =   375
      Left            =   720
      TabIndex        =   51
      Top             =   9435
      Width           =   960
   End
   Begin VB.Label lbl_Object 
      BackStyle       =   0  '투명
      Caption         =   "메     모"
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
      Height          =   450
      Left            =   7935
      TabIndex        =   50
      Top             =   6255
      Width           =   1185
   End
   Begin VB.Label lbl_EndDate 
      BackStyle       =   0  '투명
      Caption         =   "종  료  일"
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
      Height          =   375
      Left            =   720
      TabIndex        =   49
      Top             =   9930
      Width           =   960
   End
   Begin VB.Label lbl_dept 
      BackStyle       =   0  '투명
      Caption         =   "직       급"
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
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   48
      Top             =   8970
      Width           =   960
   End
   Begin VB.Label lbl_Num 
      BackStyle       =   0  '투명
      Caption         =   "등록일시"
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
      Height          =   375
      Left            =   720
      TabIndex        =   47
      Top             =   6180
      Width           =   1020
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "이       름"
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
      Height          =   375
      Left            =   720
      TabIndex        =   46
      Top             =   7110
      Width           =   1020
   End
   Begin VB.Label lbl_CarNo 
      BackStyle       =   0  '투명
      Caption         =   "차량번호"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   720
      TabIndex        =   45
      Top             =   6660
      Width           =   1020
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '투명
      Caption         =   "차량 등록 관리"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   690
      TabIndex        =   44
      Top             =   150
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "부제적용"
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
      Height          =   465
      Left            =   18510
      TabIndex        =   43
      Top             =   4305
      Width           =   1185
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   720
      Left            =   30
      Top             =   45
      Width           =   15735
   End
End
Attribute VB_Name = "Frm_Canon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpCarNo As String
Dim tmpIndexNo As String
Dim PART_NAME_TMP As String
Dim RegQry As String
Public Glo_EndDate As Integer

Private Sub cmd_PhotoUpdate_Click()
    '사진 등록하기
    Dim mystream As ADODB.Stream
    Dim rs As ADODB.Recordset

    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    
    mystream.Open
    mystream.LoadFromFile txt_PhotoPath.Text

    Set rs = New ADODB.Recordset
    rs.Open "Select * From tb_reg_week Where CAR_NO = '" & txt_CarNo & "'", adoConn, adOpenStatic, adLockOptimistic
    rs!DRIVER_PIC = mystream.Read
    rs.Update
    
    mystream.Close
    rs.Close

    Call Photo_Show(txt_CarNo)

End Sub

Private Sub Photo_Show(Data As String)
Dim rs As Recordset
Dim Qry As String
Dim mystream As ADODB.Stream

    Qry = "Select * From tb_reg_week Where CAR_NO = '" & Data & "'"
    
    Set rs = New ADODB.Recordset
    Set mystream = New ADODB.Stream
    rs.Open Qry, adoConn
    
    If Not (rs.EOF) Then
        If Not IsNull(rs!DRIVER_PIC) Then
            mystream.Type = adTypeBinary
            mystream.Open
            mystream.Write rs!DRIVER_PIC
            mystream.SaveToFile App.Path & "\tmpDriver.jpg", adSaveCreateOverWrite
            ImgDriver.Picture = LoadPicture(App.Path & "\tmpDriver.jpg")
            mystream.Close
        Else
            ImgDriver.Picture = LoadPicture(App.Path & "\NoUser.jpg")
        End If
    Else
        ImgDriver.Picture = LoadPicture(App.Path & "\NoUser.jpg")
        Beep
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdFind_Click()
    FrmPhoto.Show 1
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim rs As Recordset
Dim Qry As String

    Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.

    'cmb_Dong
    Qry = "SELECT DRIVER_DEPT From tb_reg_week Group By DRIVER_DEPT"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    Do While Not (rs.EOF)
        cmb_Dong.AddItem rs!DRIVER_DEPT
        rs.MoveNext
    Loop
    Set rs = Nothing

    '운영 모드
    'User_Type = Get_Ini("System Config", "User_Type", "0")
    
    '정기권 등록시 종료일 기본 설정
    Glo_EndDate = Val(Get_Ini("등록 설정", "종료일 기본값", "99"))
    
    RegQry = "SELECT * From tb_reg_week ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"

    With cmb_Gubun
        .AddItem "정기권"
        .AddItem "업무용"
        .AddItem "협력업체"
        .AddItem "예외처리"
        .AddItem "출입제한"
        .Text = cmb_Gubun.List(0)
    End With
    With cmb_Search
        .AddItem "전체"
        .AddItem "정기권"
        .AddItem "업무용"
        .AddItem "협력업체"
        .AddItem "예외처리"
        .AddItem "출입제한"
        .AddItem "기간초과"
        .Text = cmb_Search.List(0)
    End With
    
    Call Clear_Field
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 시작...!!", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 시작...!!")

End Sub

'1개월 연장
Private Sub cmd_Month_Click()
    If (MaskEdBox_End.Text <> "9999-12-31") Then
        MaskEdBox_End.Text = DateAdd("m", 1, MaskEdBox_End.Text)
    End If
End Sub

Public Sub ListView_REG_SQL()
Dim rs As Recordset
Dim Qry As String
Dim itmX As ListItem
Dim INDEX_NO As Long

INDEX_NO = 1
Set rs = New ADODB.Recordset
rs.Open RegQry, adoConn
lbl_COUNT = rs.RecordCount
Do While Not (rs.EOF)
    Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
    itmX.SubItems(1) = "" & rs!CAR_NO
    itmX.SubItems(2) = "" & rs!CAR_MODEL
    itmX.SubItems(3) = "" & rs!CAR_GUBUN
    itmX.SubItems(4) = "" & rs!DRIVER_NAME
    itmX.SubItems(5) = "" & rs!DRIVER_PHONE
    itmX.SubItems(6) = "" & rs!DRIVER_DEPT
    itmX.SubItems(7) = "" & rs!DRIVER_CLASS
    itmX.SubItems(8) = "" & rs!Start_Date
    itmX.SubItems(9) = "" & rs!End_Date
    itmX.SubItems(10) = "" & rs!REG_DATE
    itmX.SubItems(11) = "" & rs!Update_date
    itmX.SubItems(12) = "" & rs!DAY_MON
    itmX.SubItems(13) = "" & rs!DAY_TUE
    itmX.SubItems(14) = "" & rs!DAY_WEN
    itmX.SubItems(15) = "" & rs!DAY_THU
    itmX.SubItems(16) = "" & rs!DAY_FRI
    itmX.SubItems(17) = "" & rs!DAY_SAT
    itmX.SubItems(18) = "" & rs!DAY_SUN
    
    rs.MoveNext
    INDEX_NO = INDEX_NO + 1
Loop
Set rs = Nothing
End Sub

Public Sub ListView_REG_Draw()
Dim Column_to_size As Integer

With Me
    Call ListViewExtended(.ListView_REG)
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No  "
    .ListView_REG.ColumnHeaders.Add , , " 차량번호        "
    .ListView_REG.ColumnHeaders.Add , , " 차량모델     "
    .ListView_REG.ColumnHeaders.Add , , " 차량구분   "
    .ListView_REG.ColumnHeaders.Add , , " 이    름      "
    .ListView_REG.ColumnHeaders.Add , , " 연 락 처              "
    .ListView_REG.ColumnHeaders.Add , , " 소    속         "
    .ListView_REG.ColumnHeaders.Add , , " 직    급         "
    .ListView_REG.ColumnHeaders.Add , , " 시 작 일      "
    .ListView_REG.ColumnHeaders.Add , , " 종 료 일      "
    .ListView_REG.ColumnHeaders.Add , , " 등 록 일                         "
    .ListView_REG.ColumnHeaders.Add , , " 수 정 일                         "
    .ListView_REG.ColumnHeaders.Add , , " 월  "
    .ListView_REG.ColumnHeaders.Add , , " 화  "
    .ListView_REG.ColumnHeaders.Add , , " 수  "
    .ListView_REG.ColumnHeaders.Add , , " 목  "
    .ListView_REG.ColumnHeaders.Add , , " 금  "
    .ListView_REG.ColumnHeaders.Add , , " 토  "
    .ListView_REG.ColumnHeaders.Add , , " 일  "
    .ListView_REG.ColumnHeaders.Add , , " "
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
ListView_REG.SetFocus
txt_CarNo = ListView_REG.SelectedItem.SubItems(1)
Call Photo_Show(txt_CarNo)
End Sub

Public Sub Clear_Field()
    tmpCarNo = ""
    tmpIndexNo = ""
    txt_Num.Text = ""
    txt_CarNo.Text = ""
    txt_Name.Text = ""
    txt_Phone.Text = ""
    txt_CarModel.Text = ""
    cmb_Gubun.ListIndex = 0
    cmb_Rotation.ListIndex = 0
    'txt_Dong.Text = ""
    cmb_Dong.Text = ""
    txt_Ho.Text = ""
    MaskEdBox_Start.Text = Format(Now, "yyyy-mm-dd")
    '종료일 설정
    Select Case Glo_EndDate
        Case 99
            MaskEdBox_End.Text = "9999-12-31"
        Case Else
            MaskEdBox_End.Text = Format(DateAdd("m", Glo_EndDate, Date), "yyyy-mm-dd")
    End Select
    MaskEdBox_Fee.Text = "0"
    txt_Object.Text = ""
    chk_Week(0).value = 1
    chk_Week(1).value = 1
    chk_Week(2).value = 1
    chk_Week(3).value = 1
    chk_Week(4).value = 1
    chk_Week(5).value = 1
    chk_Week(6).value = 1
    
    txt_PhotoPath.Text = ""
    ImgDriver.Picture = LoadPicture(App.Path & "\NoUser.jpg")
    
    On Error Resume Next
    txt_CarNo.SetFocus
End Sub

'데이터 삭제
Sub Delete_Record()
    adoConn.Execute "DELETE FROM tb_reg_week WHERE CAR_NO = '" & txt_CarNo & "'"
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    차량정보 삭제 완료", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    차량정보 삭제 완료")
    Call ListView_REG_Draw
    Call ListView_REG_SQL
End Sub

Sub Insert_Record()
    Dim rs_COUNT As Recordset
    Dim rs As Recordset
    Dim SQL_COUNT As String
    Dim SQL_QUARY As String
    Dim i As Integer
    Dim Cnt As Integer
    Dim tmp As String
    Dim tmpName, tmpPhone As String
    Dim P As String

'    tmpName = Encrypt(txt_Name)
'    tmpPhone = Encrypt(txt_Phone)

    If (tmpIndexNo = "") Then '신규등록
        'INSERT
        adoConn.Execute "INSERT INTO tb_reg_week (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, DAY_MON, DAY_TUE, DAY_WEN, DAY_THU, DAY_FRI, DAY_SAT, DAY_SUN) VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun.Text & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & txt_Object & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '" & chk_Week(0).value & "', '" & chk_Week(1).value & "', '" & chk_Week(2).value & "', '" & chk_Week(3).value & "', '" & chk_Week(4).value & "', '" & chk_Week(5).value & "', '" & chk_Week(6).value & "')"
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    차량등록 완료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    차량등록 완료")
'        If (MaskEdBox_Fee <> "0") Then
'            '대화상자 처리해야됨...!!!
'            MBox.Label3.Caption = txt_CarNo.Text & vbCrLf & MaskEdBox_Fee.Text & "원"
'            MBox.Label3.FontSize = 20
'            MBox.Label1.Caption = "위 차량의 차량결제를 등록합니다. 등록하시겠습니까?"
'            MBox.Label2.Caption = "차량결제 정보 등록"
'            MBox.Show 1
'            If (Glo_MsgRet = True) Then
'                adoConn.Execute "UPDATE tb_reg SET FEE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
'                adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
'                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료", 0
'                Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료")
'            End If
'        End If
        If (Len(txt_PhotoPath.Text) <> 0) Then
            Call cmd_PhotoUpdate_Click
        End If
    Else
        If (tmpCarNo <> txt_CarNo.Text) Then '기존 차량번호를 변경하면
            adoConn.Execute "UPDATE tb_reg_week SET CAR_NO = '" & txt_CarNo & "', CAR_MODEL = '" & txt_CarModel & "', CAR_GUBUN = '" & cmb_Gubun & "', CAR_FEE = '" & MaskEdBox_Fee.Text & "', DRIVER_NAME = '" & txt_Name & "', DRIVER_PHONE = '" & txt_Phone & "', DRIVER_DEPT = '" & cmb_Dong & "', DRIVER_CLASS = '" & txt_Ho & "', START_DATE = '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', END_DATE = '" & Format(MaskEdBox_End, "YYYYMMDD") & "', ETC = '" & txt_Object & "', UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "',DAY_MON = '" & chk_Week(0).value & "', DAY_TUE = '" & chk_Week(1).value & "', DAY_WEN = '" & chk_Week(2).value & "', DAY_THU = '" & chk_Week(3).value & "', DAY_FRI = '" & chk_Week(4).value & "', DAY_SAT = '" & chk_Week(5).value & "', DAY_SUN = '" & chk_Week(6).value & "' WHERE CAR_NO = '" & tmpCarNo & "'"
        Else
            adoConn.Execute "UPDATE tb_reg_week SET CAR_MODEL = '" & txt_CarModel & "', CAR_GUBUN = '" & cmb_Gubun & "', CAR_FEE = '" & MaskEdBox_Fee.Text & "', DRIVER_NAME = '" & txt_Name & "', DRIVER_PHONE = '" & txt_Phone & "', DRIVER_DEPT = '" & cmb_Dong & "', DRIVER_CLASS = '" & txt_Ho & "', START_DATE = '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', END_DATE = '" & Format(MaskEdBox_End, "YYYYMMDD") & "', ETC = '" & txt_Object & "', UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "',DAY_MON = '" & chk_Week(0).value & "', DAY_TUE = '" & chk_Week(1).value & "', DAY_WEN = '" & chk_Week(2).value & "', DAY_THU = '" & chk_Week(3).value & "', DAY_FRI = '" & chk_Week(4).value & "', DAY_SAT = '" & chk_Week(5).value & "', DAY_SUN = '" & chk_Week(6).value & "' WHERE CAR_NO = '" & tmpCarNo & "'"
        End If
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    차량정보 수정 완료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    차량정보 수정 완료")
    End If
    
    Call ListView_REG_Draw
    Call ListView_REG_SQL

On Error Resume Next
    
'    If (Err = 3022) Then
'        Msg_Box.Label2.Caption = "데이터 베이스 오류"
'        Msg_Box.Label1.Caption = "중복된 차량번호를 허용하지않습니다."
'        Msg_Box.Show 1
'    End If

End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String

On Error Resume Next

Select Case Index
    Case 0  '종료
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 종료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 종료")
        Unload Me
        Exit Sub
       
    Case 1  '신규입력
        If (tmpIndexNo = "") Then
            RegQry = "SELECT * From tb_reg_week WHERE CAR_NO = '" & txt_CarNo & "'"
            If (Data_Error_Check = False) Then
                Msg_Box.Label2.Caption = "필드 입력 오류"
                Msg_Box.Label1.Caption = "중요한 항목을 입력하지 않았습니다."
                Msg_Box.Show 1
            Else
                Call Insert_Record
                Call Clear_Field
            End If
        Else
            Msg_Box.Label2.Caption = "신규 데이터 입력 오류"
            Msg_Box.Label1.Caption = "신규 데이터가 아닙니다." & vbCrLf & vbCrLf & " 다시 한번 확인하세요."
            Msg_Box.Show 1
            Call Clear_Field
        End If
        Exit Sub
    
    Case 2  '삭제
        If (tmpIndexNo = "") Then
           Call Clear_Field
           Exit Sub
        End If
        If (tmpCarNo <> Me.txt_CarNo) Then
            Msg_Box.Label2.Caption = "데이터 선택 오류"
            Msg_Box.Label1.Caption = "삭제할 데이터를 다시 선택해 주십시요."
            Msg_Box.Show 1
            Exit Sub
        End If
        MBox.Label3.Caption = txt_CarNo.Text
        MBox.Label1.Caption = "위 차량의 차량등록 정보를 삭제합니다." & vbCrLf & vbCrLf & " 삭제하시겠습니까?"
        MBox.Label2.Caption = "차량등록 정보 삭제"
        MBox.Show 1
        If (Glo_MsgRet = True) Then
           Call Delete_Record
        End If
        Call Clear_Field
        Exit Sub
        
    Case 3   '초기화
        Call Clear_Field
        Exit Sub
            
    Case 4  '수정
        If (tmpIndexNo = "") Then
            Msg_Box.Label2.Caption = "필드 오류"
            Msg_Box.Label1.Caption = "신규 등록자료 입니다." & vbCrLf & vbCrLf & " 다시 확인 하세요."
            Msg_Box.Show 1
            Exit Sub
        Else
            If (txt_CarNo.Text = tmpCarNo) Then
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "필드 입력 오류"
                    Msg_Box.Label1.Caption = "중요한 항목을 누락 또는 잘못 입력하였습니다."
                    Msg_Box.Show 1
                Else
                    MBox.Label3.Caption = txt_CarNo.Text
                    MBox.Label1.Caption = "선택하신 차량등록 정보가 변경됩니다." & vbCrLf & vbCrLf & " 수정 하시겠습니까?"
                    MBox.Label2.Caption = "차량등록 자료 수정"
                    MBox.Show 1
                    If (Glo_MsgRet = True) Then
                       RegQry = "SELECT * From tb_reg_week WHERE CAR_NO = '" & txt_CarNo & "'"
                       Call Insert_Record
                       Call Clear_Field
                       'txt_CarNo.SetFocus
                    End If
                End If
            Else
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "필드 입력 오류"
                    Msg_Box.Label1.Caption = "중요한 항목을 누락 또는 잘못 입력하였습니다."
                    Msg_Box.Show 1
                Else
                    MBox.Label3.Caption = txt_CarNo.Text
                    MBox.Label1.Caption = "선택하신 자료의 차량번호가 변경됩니다." & vbCrLf & vbCrLf & " 수정 하시겠습니까?"
                    MBox.Label2.Caption = "차량등록 정보 수정"
                    MBox.Show 1
                    If (Glo_MsgRet = True) Then
                       RegQry = "SELECT * From tb_reg_week WHERE CAR_NO = '" & txt_CarNo & "'"
                       Call Insert_Record
                       Call Clear_Field
                       'txt_CarNo.SetFocus
                    End If
                End If
            End If
        End If
        Exit Sub

    Case 5
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_등록차량_" & cmb_Search.Text & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "검색내역")
        Exit Sub
        
    Case 6
        '차량등록정보 검색
        Select Case cmb_Search.Text
            Case "전체"
                RegQry = "SELECT * From tb_reg_week ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"
            Case "기간초과"
                '기간초과차량검색
                RegQry = "SELECT * From tb_reg_week WHERE END_DATE < " & Format(Now, "YYYYMMDD") & " ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"
            Case Else
                RegQry = "SELECT * From tb_reg_week WHERE CAR_GUBUN = '" & cmb_Search.Text & "' ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"
        End Select
        'Lbl_search.Caption = cmb_Search.Text
        Call Clear_Field
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Exit Sub
        
'    Case 7  '결제
'        If (tmpIndexNo <> "") Then
'            If (MaskEdBox_Fee <> "0") Then
'                '대화상자 처리해야됨...!!!
'                MBox.Label3.Caption = txt_CarNo.Text & vbCrLf & MaskEdBox_Fee.Text & "원"
'                MBox.Label3.FontSize = 20
'                MBox.Label1.Caption = "위 차량의 차량결제를 등록합니다." & vbCrLf & vbCrLf & " 등록하시겠습니까?"
'                MBox.Label2.Caption = "차량결제 정보 등록"
'                MBox.Show 1
'                If (Glo_MsgRet = True) Then
'                    adoConn.Execute "UPDATE tb_reg SET FEE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
'                    adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
'                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료", 0
'                    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료")
'                End If
'            Else
'                MsgBox "잘못된 금액입니다. 확인하세요."
'            End If
'        Else
'            MsgBox "잘못된 명령입니다. 확인하세요."
'        End If
'        Call Clear_Field
'        Call ListView_REG_Draw
'        Call ListView_REG_SQL
'        Exit Sub

End Select

End Sub


'필수 입력 데이터 확인
Private Function Data_Error_Check()
    Dim Error_Flag As Boolean
        
    Error_Flag = True
    
    If Not ((LenH(txt_CarNo.Text) = 11) Or (LenH(txt_CarNo.Text) = 12) Or (LenH(txt_CarNo.Text) = 8)) Then
        Error_Flag = False
    End If
    If (LenH(txt_CarNo.Text) = 0) Then
        Error_Flag = False
    End If
    If (IsNumeric(MaskEdBox_Fee.Text) = False) Then
        MaskEdBox_Fee.Text = "0"
        'Error_Flag = False
    End If
    If (LenH(txt_Ho.Text) = 0) Then
        'txt_Phone.Text = " "
        'Error_Flag = False
    Else
        txt_Ho.Text = Mid(txt_Ho.Text, 1, 16)
    End If
    If (LenH(cmb_Dong.Text) = 0) Then
        'txt_CarModel.Text = " "
        'Error_Flag = False
    Else
        cmb_Dong.Text = MidH(cmb_Dong.Text, 1, 16)
    End If
    
    If (IsDate(MaskEdBox_Start.Text) = False) Then
        Error_Flag = False
    End If
    If (IsDate(MaskEdBox_End.Text) = False) Then
        Error_Flag = False
    End If
    If (Len(txt_Object.Text) = 0) Then
        txt_Object.Text = " "
        'Error_Flag = False
    Else
        txt_Object.Text = MidH(txt_Object.Text, 1, 64)
    End If
    
    Data_Error_Check = Error_Flag

End Function

Private Sub txt_CarNo_Change()
    If (LenH(txt_CarNo) > 7) Then
        Call Search_Record
    End If
End Sub

Sub Search_Record()
    Dim rs As Recordset
    Dim SQL_SEARCH As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    
    SQL_SEARCH = "SELECT * From tb_reg_week WHERE CAR_NO = '" & txt_CarNo & "' ORDER BY CAR_GUBUN"
    'Debug.Print SQL_SEARCH
    Set rs = New ADODB.Recordset
    rs.Open SQL_SEARCH, adoConn
    
    If (rs.RecordCount <> 0) Then
        tmpCarNo = rs!CAR_NO
        tmpIndexNo = ListView_REG.SelectedItem.Text
        txt_Num = "" & rs!REG_DATE
        txt_Name = "" & rs!DRIVER_NAME
        txt_Phone = "" & rs!DRIVER_PHONE
        txt_CarModel = "" & rs!CAR_MODEL
        MaskEdBox_Fee.Text = "" & rs!CAR_FEE
        'txt_Dong = "" & rs!DRIVER_DEPT
        cmb_Dong = "" & rs!DRIVER_DEPT
        txt_Ho = "" & rs!DRIVER_CLASS
        MaskEdBox_Start.Text = Format(rs!Start_Date, "####-##-##")
        MaskEdBox_End.Text = Format(rs!End_Date, "####-##-##")
        cmb_Gubun.Text = rs!CAR_GUBUN
        txt_Object = "" & rs!ETC
    
        'Week
        chk_Week(0).value = Val(rs!DAY_MON)
        chk_Week(1).value = Val(rs!DAY_TUE)
        chk_Week(2).value = Val(rs!DAY_WEN)
        chk_Week(3).value = Val(rs!DAY_THU)
        chk_Week(4).value = Val(rs!DAY_FRI)
        chk_Week(5).value = Val(rs!DAY_SAT)
        chk_Week(6).value = Val(rs!DAY_SUN)
        
        Call Photo_Show(txt_CarNo)
    
    End If
    Set rs = Nothing

End Sub

'엔터키 입력시 탭 실행
'폼속성 keypreview = true 설정
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim Car_Num_Str As String
Dim Qry As String
Dim rs As Recordset
Dim rs_Part As Recordset
Dim itmX As ListItem
    
If (KeyAscii = 13) Then
    If (Len(txt_tmpCarNo) <> 0) Then
        If (Len(txt_tmpCarNo) = 0) Then
            MsgBox "검색 대상을 정확하게 입력하세요!"
            txt_tmpCarNo = ""
            Exit Sub
        Else
            Select Case cmb_Sch.Text
                Case "차량번호"
                    RegQry = "SELECT * From tb_reg_week WHERE CAR_NO LIKE '%" & txt_tmpCarNo & "'"
                Case "이름"
                    RegQry = "SELECT * From tb_reg_week WHERE DRIVER_NAME LIKE '%" & txt_tmpCarNo & "%'"
                Case Else
                    RegQry = "SELECT * From tb_reg_week"
            End Select
            txt_tmpCarNo = ""
            Call Clear_Field
            Call ListView_REG_Draw
            Call ListView_REG_SQL
            KeyAscii = 0
            Exit Sub
        End If
    End If
End If

If (KeyAscii = vbKeyReturn) Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If

End Sub



