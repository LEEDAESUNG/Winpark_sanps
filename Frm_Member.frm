VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_Member 
   Caption         =   " 등록권 관리"
   ClientHeight    =   10530
   ClientLeft      =   1950
   ClientTop       =   2430
   ClientWidth     =   15195
   BeginProperty Font 
      Name            =   "나눔고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   15195
   Begin VB.ComboBox cmb_LotCode 
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
      Left            =   5910
      TabIndex        =   16
      Text            =   "cmb_LotCode"
      Top             =   8355
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Caption         =   " 결제처리"
      Height          =   1170
      Left            =   8895
      TabIndex        =   55
      Top             =   6930
      Width           =   6180
      Begin VB.CommandButton cmd_Month 
         BackColor       =   &H0000FFFF&
         Caption         =   "1개월 연장"
         Height          =   495
         Left            =   3585
         MaskColor       =   &H80000004&
         TabIndex        =   21
         Top             =   420
         Width           =   1170
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   510
         Index           =   7
         Left            =   4860
         TabIndex        =   22
         Top             =   405
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "결 제"
         ForeColor       =   65535
         RoundedCorners  =   0   'False
         Picture         =   "Frm_Member.frx":0000
      End
      Begin VB.Label lbl_tmpCode 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   375
         TabIndex        =   58
         Top             =   720
         Width           =   2985
      End
      Begin VB.Label lbl_tmpCarNo 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   375
         TabIndex        =   57
         Top             =   360
         Width           =   2985
      End
   End
   Begin VB.ComboBox cmb_State 
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
      Left            =   5910
      Style           =   2  '드롭다운 목록
      TabIndex        =   15
      Top             =   7875
      Width           =   2325
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
      Left            =   5910
      TabIndex        =   14
      Text            =   "cmb_Rotation"
      Top             =   7365
      Width           =   2325
   End
   Begin VB.ComboBox cmb_Sch 
      Height          =   345
      ItemData        =   "Frm_Member.frx":0351
      Left            =   7590
      List            =   "Frm_Member.frx":035E
      TabIndex        =   25
      Text            =   "차량번호"
      Top             =   210
      Width           =   1575
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
      Left            =   1740
      TabIndex        =   4
      Text            =   "cmb_Gubun"
      Top             =   6855
      Width           =   2340
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
      Left            =   1740
      TabIndex        =   2
      Top             =   5850
      Width           =   2325
   End
   Begin VB.TextBox txt_Object 
      Height          =   1380
      Left            =   8880
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   17
      Top             =   5325
      Width           =   6165
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
      Left            =   1740
      TabIndex        =   8
      Top             =   8835
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
      Left            =   1740
      TabIndex        =   6
      Top             =   7845
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
      Left            =   1740
      TabIndex        =   5
      Top             =   7350
      Width           =   2325
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
      Left            =   9255
      TabIndex        =   26
      Top             =   195
      Width           =   2115
   End
   Begin VB.ComboBox cmb_Search 
      Height          =   345
      ItemData        =   "Frm_Member.frx":037E
      Left            =   11325
      List            =   "Frm_Member.frx":0394
      TabIndex        =   27
      Text            =   "등록일"
      Top             =   1050
      Width           =   2445
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
      Left            =   1740
      TabIndex        =   3
      Top             =   6345
      Width           =   2325
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1185
      Left            =   15
      TabIndex        =   31
      Top             =   9330
      Width           =   15180
   End
   Begin VB.ComboBox cmb_DayNight 
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
      Left            =   5910
      Style           =   2  '드롭다운 목록
      TabIndex        =   13
      Top             =   6870
      Width           =   2325
   End
   Begin VB.TextBox txt_Num 
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
      Height          =   375
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   1
      Top             =   5370
      Width           =   2325
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
      Left            =   1740
      TabIndex        =   7
      Top             =   8340
      Width           =   2325
   End
   Begin VB.TextBox txt_CardNo 
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
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   0
      Top             =   4905
      Width           =   2325
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   3135
      Left            =   15
      TabIndex        =   30
      Top             =   1545
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   5530
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
      Left            =   5910
      TabIndex        =   9
      Top             =   4905
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
      Height          =   525
      Index           =   0
      Left            =   13860
      TabIndex        =   29
      Top             =   135
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   926
      _StockProps     =   78
      Caption         =   "닫 기"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":03CA
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   2
      Left            =   13770
      TabIndex        =   20
      Top             =   8565
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "삭 제"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":071B
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   4
      Left            =   12660
      TabIndex        =   19
      Top             =   8565
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "수 정"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":0A6C
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   1
      Left            =   11550
      TabIndex        =   18
      Top             =   8565
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "등 록"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":0DBD
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   3
      Left            =   10440
      TabIndex        =   23
      Top             =   8565
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "초기화"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":110E
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   525
      Index           =   5
      Left            =   12555
      TabIndex        =   24
      Top             =   135
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   926
      _StockProps     =   78
      Caption         =   "Excel"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":145F
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   480
      Index           =   6
      Left            =   13845
      TabIndex        =   28
      Top             =   975
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "정 렬"
      ForeColor       =   14737632
      RoundedCorners  =   0   'False
      Picture         =   "Frm_Member.frx":17B0
   End
   Begin MSMask.MaskEdBox MaskEdBox_End 
      Height          =   375
      Left            =   5910
      TabIndex        =   10
      Top             =   5385
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
   Begin MSMask.MaskEdBox MaskEdBox_Reg 
      Height          =   375
      Left            =   5910
      TabIndex        =   11
      Top             =   5865
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
      Left            =   5910
      TabIndex        =   12
      Top             =   6360
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
   Begin VB.Label Label8 
      BackStyle       =   0  '투명
      Caption         =   "주차코드"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   56
      Top             =   8400
      Width           =   1305
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '투명
      Caption         =   "입출상태"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   54
      Top             =   7920
      Width           =   1305
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "등 록 일"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   53
      Top             =   5910
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "구     분"
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
      Height          =   465
      Left            =   480
      TabIndex        =   52
      Top             =   6900
      Width           =   1185
   End
   Begin VB.Label lbl_dept 
      BackStyle       =   0  '투명
      Caption         =   "구분1 / 동"
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
      Height          =   465
      Index           =   0
      Left            =   480
      TabIndex        =   51
      Top             =   8370
      Width           =   1305
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
      ForeColor       =   &H00000000&
      Height          =   465
      Index           =   0
      Left            =   480
      TabIndex        =   50
      Top             =   6390
      Width           =   1305
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   5505
      TabIndex        =   49
      Top             =   1065
      Width           =   1305
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6945
      TabIndex        =   48
      Top             =   1065
      Width           =   1425
   End
   Begin VB.Label lbl_Phone 
      BackStyle       =   0  '투명
      Caption         =   "연 락 처"
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
      Height          =   465
      Left            =   480
      TabIndex        =   47
      Top             =   7875
      Width           =   1305
   End
   Begin VB.Label lbl_StartDate 
      BackStyle       =   0  '투명
      Caption         =   "시 작 일"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   46
      Top             =   4935
      Width           =   1305
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   8940
      TabIndex        =   45
      Top             =   4920
      Width           =   1185
   End
   Begin VB.Label lbl_EndDate 
      BackStyle       =   0  '투명
      Caption         =   "종 료 일"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   44
      Top             =   5430
      Width           =   1305
   End
   Begin VB.Label lbl_dept 
      BackStyle       =   0  '투명
      Caption         =   "구분2 / 호"
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
      Height          =   465
      Index           =   1
      Left            =   480
      TabIndex        =   43
      Top             =   8835
      Width           =   1305
   End
   Begin VB.Label lbl_Num 
      BackStyle       =   0  '투명
      Caption         =   "등록코드"
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
      Height          =   465
      Left            =   480
      TabIndex        =   42
      Top             =   5400
      Width           =   1305
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "고 객 명"
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
      Height          =   465
      Left            =   480
      TabIndex        =   41
      Top             =   7395
      Width           =   1305
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
      Height          =   465
      Left            =   480
      TabIndex        =   40
      Top             =   5895
      Width           =   1305
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '투명
      Caption         =   "# 차량 등록 현황"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   195
      TabIndex        =   39
      Top             =   1005
      Width           =   2295
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '투명
      Caption         =   "간편 검색"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   2
      Left            =   5685
      TabIndex        =   38
      Top             =   225
      Width           =   1770
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "차 량 등 록 / 관 리"
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
      Height          =   435
      Left            =   360
      TabIndex        =   37
      Top             =   210
      Width           =   3525
   End
   Begin VB.Label lbl_Search 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "전체"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3285
      TabIndex        =   36
      Top             =   1065
      Width           =   1875
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
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4680
      TabIndex        =   35
      Top             =   7410
      Width           =   1185
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Caption         =   "요금구분"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   34
      Top             =   6915
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "월정요금"
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
      Height          =   465
      Left            =   4680
      TabIndex        =   33
      Top             =   6405
      Width           =   1305
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Caption         =   "카드번호"
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
      Height          =   465
      Left            =   480
      TabIndex        =   32
      Top             =   4950
      Width           =   1305
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   795
      Left            =   0
      Top             =   0
      Width           =   15210
   End
End
Attribute VB_Name = "Frm_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Qry_Reg As String
Dim CardNo_tmp As String
Dim CarNo_tmp As String
Dim IndexNo_tmp As String

Private Sub Form_Load()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim Qry As String
    
    Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    'cmb_Gubun
    Qry = "SELECT CAR_GUBUN From tb_member Group By CAR_GUBUN"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    
    Do While Not (rs.EOF)
        If (rs!CAR_GUBUN <> "정기권") Or (rs!CAR_GUBUN <> "분실권") Or (rs!CAR_GUBUN <> "예외처리") Then
            cmb_Gubun.AddItem rs!CAR_GUBUN
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    cmb_Gubun.Text = cmb_Gubun.List(0)
    
    'cmb_DayNight
    Qry = "SELECT DAY_NIGHT From tb_member Group By DAY_NIGHT"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    Do While Not (rs.EOF)
        If (rs!DAY_NIGHT <> "") Then
                cmb_DayNight.AddItem rs!DAY_NIGHT
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    cmb_DayNight.Text = cmb_DayNight.List(0)
    
    'cmb_Rotation
    Qry = "SELECT ROTATION From tb_member Group By ROTATION"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    Do While Not (rs.EOF)
        If (rs!ROTATION <> "") Then
                cmb_Rotation.AddItem rs!ROTATION
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    cmb_Rotation.Text = cmb_Rotation.List(0)
    
    'cmb_State
    Qry = "SELECT INOUT_ST From tb_member Group By INOUT_ST"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    Do While Not (rs.EOF)
        If (rs!INOUT_ST <> "") Then
                cmb_State.AddItem rs!INOUT_ST
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    cmb_State.Text = cmb_State.List(0)
    
    'cmb_LotCode
    Qry = "SELECT LOT_CODE From tb_member Group By LOT_CODE"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    Do While Not (rs.EOF)
        If (rs!LOT_CODE <> "") Then
                cmb_LotCode.AddItem rs!LOT_CODE
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    cmb_LotCode.Text = cmb_LotCode.List(0)
    
    'Me.cmb_Gubun = Me.cmb_Gubun.List(0)
    Qry_Reg = "SELECT * From tb_member ORDER BY CAR_NO"
    cmb_Gubun.Text = cmb_Gubun.List(0)
    'lbl_dept(0).Caption = "소  속 :"
    'lbl_dept(1).Caption = "거주 호 :"
    cmb_Search.Text = cmb_Search.List(0)
    
    Call Clear_Field
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 시작...!!", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 시작...!!")

End Sub

'1개월 연장
Private Sub cmd_Month_Click()
    MaskEdBox_End.Text = DateAdd("m", 1, MaskEdBox_End.Text)
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    1 개월 연장 버튼 클릭 - " & lbl_tmpCarNo & "  " & lbl_tmpCode)
End Sub

Public Sub ListView_REG_SQL()
    Dim rs As ADODB.Recordset
    'Dim Qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    
    INDEX_NO = 1
    Set rs = New ADODB.Recordset
    rs.Open Qry_Reg, adoConn
    lbl_COUNT = rs.RecordCount
    Do While Not (rs.EOF)
        Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
        itmX.SubItems(1) = "" & rs!RF_NO
        itmX.SubItems(2) = "" & rs!RF_CODE
        itmX.SubItems(3) = "" & rs!CAR_NO
        itmX.SubItems(4) = "" & rs!CAR_MODEL
        itmX.SubItems(5) = "" & rs!CAR_GUBUN
        itmX.SubItems(6) = "" & rs!DRIVER_NAME
        itmX.SubItems(7) = "" & rs!DRIVER_PHONE
        itmX.SubItems(8) = "" & rs!DRIVER_DEPT
        itmX.SubItems(9) = "" & rs!DRIVER_CLASS
        itmX.SubItems(10) = "" & rs!DT_START
        itmX.SubItems(11) = "" & rs!DT_END
        itmX.SubItems(12) = "" & rs!DT_REG
        itmX.SubItems(13) = "" & rs!DT_UPDATE
        itmX.SubItems(14) = "" & rs!CAR_FEE
        itmX.SubItems(15) = "" & rs!DAY_NIGHT
        itmX.SubItems(16) = "" & rs!ROTATION
        itmX.SubItems(17) = "" & rs!INOUT_ST
        itmX.SubItems(18) = "" & rs!LOT_CODE
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
    .ListView_REG.ColumnHeaders.Add , , " 카드번호        "
    .ListView_REG.ColumnHeaders.Add , , " 등록코드        "
    .ListView_REG.ColumnHeaders.Add , , " 차량번호          "
    .ListView_REG.ColumnHeaders.Add , , " 차량모델       "
    .ListView_REG.ColumnHeaders.Add , , " 차량구분       "
    .ListView_REG.ColumnHeaders.Add , , " 고객명      "
    .ListView_REG.ColumnHeaders.Add , , " 연락처          "
    .ListView_REG.ColumnHeaders.Add , , " 구분1        "
    .ListView_REG.ColumnHeaders.Add , , " 구분2        "
    .ListView_REG.ColumnHeaders.Add , , " 시 작 일        "
    .ListView_REG.ColumnHeaders.Add , , " 종 료 일        "
    .ListView_REG.ColumnHeaders.Add , , " 등 록 일        "
    .ListView_REG.ColumnHeaders.Add , , " Update        "
    .ListView_REG.ColumnHeaders.Add , , " 월요금(원)     "
    .ListView_REG.ColumnHeaders.Add , , " 주차구분     "
    .ListView_REG.ColumnHeaders.Add , , " 부제적용     "
    .ListView_REG.ColumnHeaders.Add , , " 입출상태     "
    .ListView_REG.ColumnHeaders.Add , , " 주차코드     "
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With

End Sub

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
    ListView_REG.SetFocus
    If Len(ListView_REG.SelectedItem.SubItems(3)) <> 0 Then
        txt_CarNo = ListView_REG.SelectedItem.SubItems(3)
    Else
        txt_Num = ListView_REG.SelectedItem.SubItems(2)
    End If
End Sub

Public Sub Clear_Field()
    Dim i As Integer
    
    CardNo_tmp = ""
    IndexNo_tmp = ""
    txt_CardNo.Text = ""            '카드번호
    txt_Num.Text = ""                '카드코드
    txt_CarNo.Text = ""
    cmb_Gubun.ListIndex = 0
    cmb_Gubun.Text = ""
    txt_Ho.Text = ""
    txt_Name.Text = ""
    txt_Phone.Text = ""
    txt_CarModel.Text = ""
    cmb_Gubun.ListIndex = 0

    txt_Dong.Text = ""
    txt_Ho.Text = ""
    MaskEdBox_Reg.Text = Format(Now, "yyyy-mm-dd")
    MaskEdBox_Start.Text = Format(Now, "yyyy-mm-dd")

    '종료일 설정
    Select Case Glo_EndDate
        Case 99
            MaskEdBox_End.Text = "9999-12-31"
        Case Else
            MaskEdBox_End.Text = Format(DateAdd("m", Glo_EndDate, Date), "yyyy-mm-dd")
    End Select
    MaskEdBox_Fee.Text = "0"

    'cmb_DayNight.ListIndex = 0
    'cmb_Rotation.ListIndex = 0
    'cmb_State.ListIndex = 0
    'cmb_LotCode.ListIndex = 0
    
    'cmb_DayNight.Text = ""
    cmb_Rotation.Text = ""
    'cmb_State.Text = ""
    cmb_LotCode.Text = ""
    
    txt_Object.Text = ""

    lbl_tmpCarNo.Caption = ""
    lbl_tmpCode.Caption = ""

On Error Resume Next
    txt_CardNo.SetFocus

End Sub

'데이터 삭제
Sub Delete_Record()
    Dim rs As Recordset
    Dim Qry As String
    
    Qry = "SELECT * From tb_member WHERE (CAR_NO = '" & txt_CarNo & "' AND RF_NO ='" & txt_CardNo & "')"
    'Debug.Print Qry
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    Select Case rs.RecordCount
        Case 0
            MsgBox "해당되는 자료가 없습니다. 확인하세요."
        Case 1
            adoConn.Execute "DELETE FROM tb_member WHERE (CAR_NO = '" & txt_CarNo & "' AND RF_NO ='" & txt_CardNo & "')"
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CardNo & "카드     " & txt_CarNo & "    차량정보 삭제 완료", 0
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CardNo & "카드     " & txt_CarNo & "    차량정보 삭제 완료")
            Call ListView_REG_Draw
            Call ListView_REG_SQL
        Case Else
            MsgBox "해당되는 자료가 많습니다. 확인하세요."
    End Select
    Set rs = Nothing
    
End Sub

Sub Insert_Record()
    Dim rs_COUNT As Recordset
    Dim rs As Recordset
    Dim SQL_COUNT As String
    Dim SQL_QUARY As String
    Dim i As Integer
    Dim Cnt As Integer
    Dim tmp As String
    Dim Qry As String
    Dim inout As Integer
    Dim Ora_QRY As String
    Dim OraConn_F As Boolean

On Error GoTo Error

    'RF_Code 보정하기
    Cnt = 10 - txt_Num.MaxLength
    For i = 1 To Cnt
        tmp = tmp & "0"
    Next i

    If (IndexNo_tmp = "") Then '신규등록
        'INSERT
        Qry = "INSERT INTO tb_member VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone.Text & "', '" & txt_Dong.Text & "', '" & txt_Ho.Text & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & txt_Object & "', '" & Format(MaskEdBox_Reg, "YYYYMMDD") & "', '" & Format(Now, "YYYYMMDDHHNNSS") & "', ' ', '" & cmb_Rotation & "', '" & cmb_DayNight & "', '" & cmb_State.ListIndex & "', '" & cmb_LotCode & "')"
        adoConn.Execute Qry
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CardNo & " : " & txt_Num & "카드     " & txt_CarNo & "    차량등록 완료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CardNo & " : " & txt_Num & "카드     " & txt_CarNo & "    차량등록 완료")
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Qry)
        If (MaskEdBox_Fee > 0) Then
            '대화상자 처리해야됨...!!!
            MBox.Label3.Caption = txt_CarNo.Text & vbCrLf & MaskEdBox_Fee.Text & "원"
            MBox.Label3.FontSize = 20
            MBox.Label1.Caption = "위 차량의 월주차요금 결제를 처리합니다. 등록하시겠습니까?"
            MBox.Label2.Caption = "차량결제 내역 등록"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
                adoConn.Execute "UPDATE tb_member SET DT_FEE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "' AND RF_NO '" & txt_CardNo & "'"
                adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & txt_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                'HOON SCH 결재내역 입력처리
                
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료", 0
                Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료")
            Else
            
            End If
        End If
    Else
        Qry = "UPDATE tb_member SET CAR_NO = '" & txt_CarNo & "', CAR_MODEL = '" & txt_CarModel & "', CAR_GUBUN = '" & cmb_Gubun & "', CAR_FEE = '" & MaskEdBox_Fee.Text & "', DRIVER_NAME = '" & txt_Name & "', DRIVER_PHONE = '" & txt_Phone & "', DRIVER_DEPT = '" & txt_Dong & "', DRIVER_CLASS = '" & txt_Ho & "', DT_START = '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', DT_END = '" & Format(MaskEdBox_End, "YYYYMMDD") & "', ETC = '" & txt_Object & "', DT_UPDATE = '" & Format(Now, "YYYYMMDDHHNNSS") & "', ROTATION = '" & cmb_Rotation & "', DAY_NIGHT = '" & cmb_DayNight & "', INOUT_ST = '" & cmb_State.ListIndex & "', LOT_CODE = '" & cmb_LotCode & "' WHERE (CAR_NO = '" & lbl_tmpCarNo & "' AND RF_CODE = '" & lbl_tmpCode & "')"
        'Debug.Print QRY
        adoConn.Execute Qry
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "    " & txt_CardNo & "  " & txt_Num & " 카드     " & txt_CarNo & "    차량정보 수정 완료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "    " & txt_CardNo & "  " & txt_Num & " 카드     " & txt_CarNo & "    차량정보 수정 완료")
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Qry)
    End If
    
    Qry_Reg = "SELECT * FROM tb_member WHERE (CAR_NO = '" & txt_CarNo & "' AND RF_CODE = '" & txt_Num & "')"
    
    Call ListView_REG_Draw
    Call ListView_REG_SQL

Error:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & Err.Number & "  " & Err.Description, 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Number & "  " & Err.Description)
End Sub

'Private Sub Form_Activate()
'    Select Case JUNGIN_TYPE(0)
'           Case ID_TECH
'                txt_Num.MaxLength = 6
'           Case TAG_MASTER_A, TAG_MASTER_B
'                txt_Num.MaxLength = 8
'           Case REMOCON_T
'                txt_Num.MaxLength = 4
'           Case REMOCON_N
'                txt_Num.MaxLength = 5
'           Case CREDIPASS
'                txt_Num.MaxLength = 10
'    End Select
'End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String

Select Case Index
    Case 0  '종료
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 종료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 종료")
        Unload Me
        Exit Sub
       
    Case 1  '신규입력
        If (lbl_tmpCarNo.Caption = "" And lbl_tmpCode.Caption = "") Then
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
        If (lbl_tmpCarNo.Caption = "" And lbl_tmpCode.Caption = "") Then
           Call Clear_Field
           Exit Sub
        End If
        MBox.Label3.Caption = lbl_tmpCarNo & "    " & lbl_tmpCode
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
        If (lbl_tmpCarNo.Caption = "" And lbl_tmpCode.Caption = "") Then
            Msg_Box.Label2.Caption = "필드 오류"
            Msg_Box.Label1.Caption = "신규 등록자료 입니다." & vbCrLf & vbCrLf & " 다시 확인 하세요."
            Msg_Box.Show 1
            Exit Sub
        Else
            If (txt_CarNo.Text = lbl_tmpCarNo.Caption And txt_Num = lbl_tmpCode.Caption) Then
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "필드 입력 오류"
                    Msg_Box.Label1.Caption = "중요한 항목을 누락 또는 잘못 입력하였습니다."
                    Msg_Box.Show 1
                Else
                    MBox.Label3.Caption = lbl_tmpCarNo & "    " & lbl_tmpCode
                    MBox.Label1.Caption = "선택하신 차량등록 정보가 변경됩니다." & vbCrLf & vbCrLf & " 수정 하시겠습니까?"
                    MBox.Label2.Caption = "차량등록 자료 수정"
                    MBox.Show 1
                    If (Glo_MsgRet = True) Then
                       Call Insert_Record
                       Call Clear_Field
                    End If
                End If
            Else
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "필드 입력 오류"
                    Msg_Box.Label1.Caption = "중요한 항목을 누락 또는 잘못 입력하였습니다."
                    Msg_Box.Show 1
                Else
                    MBox.Label3.Caption = lbl_tmpCarNo & "    " & lbl_tmpCode
                    MBox.Label1.Caption = "선택하신 자료의 차량번호가 변경됩니다." & vbCrLf & vbCrLf & " 수정 하시겠습니까?"
                    MBox.Label2.Caption = "차량등록 정보 수정"
                    MBox.Show 1
                    If (Glo_MsgRet = True) Then
                       Call Insert_Record
                       Call Clear_Field
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
            Case "카드번호"
                Qry_Reg = "SELECT * From tb_member ORDER BY RF_NO"
            Case "차량번호"
                Qry_Reg = "SELECT * From tb_member ORDER BY CAR_NO"
            Case "고객명"
                Qry_Reg = "SELECT * From tb_member ORDER BY DRIVER_NAME"
            Case "종료일"
                Qry_Reg = "SELECT * From tb_member ORDER BY DT_END"
            Case "등록일"
                Qry_Reg = "SELECT * From tb_member ORDER BY DT_REG"
            Case "구분"
                Qry_Reg = "SELECT * From tb_member ORDER BY CAR_GUBUN"
        End Select
        lbl_Search.Caption = cmb_Search.Text
        Call Clear_Field
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Exit Sub
        
    Case 7  '결제
        If (IndexNo_tmp <> "") Then
            If (MaskEdBox_Fee <> "0") Then
                '대화상자 처리해야됨...!!!
                MBox.Label3.Caption = txt_CarNo.Text & vbCrLf & MaskEdBox_Fee.Text & "원"
                MBox.Label3.FontSize = 20
                MBox.Label1.Caption = "위 차량의 차량결제를 등록합니다." & vbCrLf & vbCrLf & " 등록하시겠습니까?"
                MBox.Label2.Caption = "차량결제 정보 등록"
                MBox.Show 1
                If (Glo_MsgRet = True) Then
                    adoConn.Execute "UPDATE tb_member SET DT_FEE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
                    adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & txt_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                    'HOON SCH 입력해야 함...!!
                    
                    
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료", 0
                    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "원    차량결제 완료")
                End If
            Else
                MsgBox "잘못된 금액입니다. 확인하세요."
            End If
        Else
            MsgBox "잘못된 명령입니다. 확인하세요."
        End If
        Call Clear_Field
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Exit Sub

End Select

On Error Resume Next

End Sub


'필수 입력 데이터 확인
Private Function Data_Error_Check()
Dim Error_Flag As Boolean
    
Error_Flag = True

If (Len(txt_CardNo.Text) = 0) Then
    txt_CardNo.Text = ""
End If
If (Len(txt_Num.Text) = 0) Then
    txt_Num.Text = ""
End If
If (Len(txt_CarNo.Text) = 0) Then
    txt_CarNo.Text = ""
End If
txt_CardNo = Trim(txt_CardNo)
txt_Num = Trim(txt_Num)
txt_CarNo = Trim(txt_CarNo)
If (LenH(txt_CarModel.Text) = 0) Then
    txt_CarModel.Text = ""
End If
If (Len(txt_Name.Text) = 0) Then
    txt_Name.Text = ""
Else
    txt_Name.Text = MidH(txt_Name.Text, 1, 16)
End If
If (LenH(txt_Phone.Text) = 0) Then
    txt_Phone.Text = ""
End If
If (LenH(txt_Dong.Text) = 0) Then
    txt_CarModel.Text = ""
End If
If (LenH(txt_Ho.Text) = 0) Then
    txt_Phone.Text = " "
End If
If (IsDate(MaskEdBox_Start.Text) = False) Then
    Error_Flag = False
End If
If (IsDate(MaskEdBox_End.Text) = False) Then
    Error_Flag = False
End If
If (IsDate(MaskEdBox_Reg.Text) = False) Then
    Error_Flag = False
End If
If (IsNumeric(MaskEdBox_Fee.Text) = False) Then
    Error_Flag = False
End If
If (LenH(txt_Object.Text) = 0) Then
    txt_Object.Text = ""
Else
    txt_Object.Text = MidH(txt_Object.Text, 1, 128)
End If

Data_Error_Check = Error_Flag

End Function

Private Sub txt_Num_Change()
    If (LenH(txt_Num) > 9) Then
        If (Len(txt_CarNo) <> 0) Then
            Exit Sub
        End If
        Call Search_Car
    End If
End Sub

Private Sub txt_CarNo_Change()
    If (LenH(txt_CarNo) >= 7) Then
        If (Len(txt_CardNo) <> 0) Then
            Exit Sub
        End If
        Call Search_Car
    End If
End Sub

Sub Search_Car()
    Dim rs As ADODB.Recordset
    Dim Qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long

    Qry = "SELECT * From tb_member WHERE CAR_NO = '" & txt_CarNo & "'"
    'Debug.Print Qry
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    If (rs.RecordCount <> 0) Then
        CardNo_tmp = "" & rs!RF_NO
        IndexNo_tmp = ListView_REG.SelectedItem.Text
        txt_Num = "" & rs!RF_CODE
        txt_CardNo = "" & rs!RF_NO
        txt_CarNo = "" & rs!CAR_NO
        txt_CarModel = "" & rs!CAR_MODEL
        txt_Name = "" & rs!DRIVER_NAME
        txt_Phone = "" & rs!DRIVER_PHONE
        txt_Dong = "" & rs!DRIVER_DEPT
        txt_Ho = "" & rs!DRIVER_CLASS
        MaskEdBox_Start.Text = Format(rs!DT_START, "####-##-##")
        MaskEdBox_End.Text = Format(rs!DT_END, "####-##-##")
        MaskEdBox_End.Text = Format(rs!DT_REG, "####-##-##")
        MaskEdBox_Fee.Text = rs!CAR_FEE
        cmb_DayNight.Text = "" & rs!DAY_NIGHT
        cmb_Rotation.Text = rs!ROTATION
        cmb_Gubun.Text = rs!CAR_GUBUN
        Select Case rs!INOUT_ST
            Case 0
                 cmb_State.ListIndex = 0
            Case 1
                 cmb_State.ListIndex = 1
            Case 2
                 cmb_State.ListIndex = 2
        End Select
        cmb_LotCode.Text = rs!LOT_CODE
        txt_Object = "" & rs!ETC
        
        lbl_tmpCarNo.Caption = rs!CAR_NO
        lbl_tmpCode.Caption = rs!RF_CODE
    Else
    
    End If
    Set rs = Nothing
End Sub

Sub Search_Card()
    Dim rs As ADODB.Recordset
    Dim Qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long

    Qry = "SELECT * From tb_member WHERE RF_NO = '" & txt_CardNo & "'"
    'Debug.Print Qry
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    If (rs.RecordCount <> 0) Then
        CardNo_tmp = rs!RF_NO
        IndexNo_tmp = ListView_REG.SelectedItem.Text
        txt_Num = "" & rs!RF_CODE
        txt_CardNo = "" & rs!RF_NO
        txt_CarNo = "" & rs!CAR_NO
        txt_CarModel = "" & rs!CAR_MODEL
        txt_Name = "" & rs!DRIVER_NAME
        txt_Phone = "" & rs!DRIVER_PHONE
        txt_Dong = "" & rs!DRIVER_DEPT
        txt_Ho = "" & rs!DRIVER_CLASS
        MaskEdBox_Start.Text = Format(rs!DT_START, "####-##-##")
        MaskEdBox_End.Text = Format(rs!DT_END, "####-##-##")
        MaskEdBox_End.Text = Format(rs!DT_REG, "####-##-##")
        MaskEdBox_Fee.Text = rs!CAR_FEE
        cmb_DayNight.Text = rs!DAY_NIGHT
        cmb_Rotation.Text = rs!ROTATION
        cmb_Gubun.Text = rs!CAR_GUBUN
        Select Case rs!INOUT_ST
            Case 0
                 cmb_State.ListIndex = 0
            Case 1
                 cmb_State.ListIndex = 1
            Case 2
                 cmb_State.ListIndex = 2
        End Select
        cmb_LotCode.Text = rs!LOT_CODE
        txt_Object = "" & rs!ETC
        
        lbl_tmpCarNo.Caption = rs!CAR_NO
        lbl_tmpCode.Caption = rs!RF_CODE
    Else
    
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
                Case "카드번호"
                    Qry_Reg = "SELECT * From tb_member WHERE RF_NO LIKE '" & txt_tmpCarNo & "'"
                Case "차량번호"
                    Qry_Reg = "SELECT * From tb_member WHERE CAR_NO LIKE '%" & txt_tmpCarNo & "'"
                Case "고객명"
                    Qry_Reg = "SELECT * From tb_member WHERE DRIVER_NAME LIKE '%" & txt_tmpCarNo & "%'"
                Case Else
                    Qry_Reg = "SELECT * From tb_member"
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




