VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Jung_New 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "정기권 등록 & 관리"
   ClientHeight    =   14955
   ClientLeft      =   3930
   ClientTop       =   1425
   ClientWidth     =   19200
   BeginProperty Font 
      Name            =   "나눔고딕"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jung_New.frx":0000
   ScaleHeight     =   14955
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmb_SGubun 
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
      ItemData        =   "Jung_New.frx":F564
      Left            =   16350
      List            =   "Jung_New.frx":F577
      TabIndex        =   69
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox cmb_Gubun 
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
      ItemData        =   "Jung_New.frx":F5A7
      Left            =   9270
      List            =   "Jung_New.frx":F5BA
      TabIndex        =   2
      Top             =   9876
      Width           =   2325
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   690
      Left            =   120
      TabIndex        =   66
      Top             =   14190
      Width           =   18975
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Jung_New.frx":F5EA
      Left            =   16350
      List            =   "Jung_New.frx":F5FD
      TabIndex        =   52
      Top             =   3030
      Width           =   1725
   End
   Begin VB.CommandButton cmd_Month 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1개월 연장"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   16890
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   51
      Top             =   10320
      Width           =   1365
   End
   Begin VB.TextBox Txt2 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   10  '한글 
      Left            =   9270
      MaxLength       =   20
      TabIndex        =   0
      Top             =   8952
      Width           =   3120
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9270
      MaxLength       =   10
      TabIndex        =   24
      Top             =   8520
      Width           =   1680
   End
   Begin VB.TextBox Txt9 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14520
      MaxLength       =   50
      TabIndex        =   9
      Top             =   10785
      Width           =   3915
   End
   Begin VB.TextBox Txt3 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   10  '한글 
      Left            =   9270
      MaxLength       =   20
      TabIndex        =   1
      Top             =   9414
      Width           =   3120
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  '평면
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
      IMEMode         =   10  '한글 
      Left            =   5595
      MaxLength       =   15
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1980
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  '평면
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
      IMEMode         =   10  '한글 
      Left            =   9180
      MaxLength       =   10
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1875
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  '평면
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
      IMEMode         =   10  '한글 
      Left            =   12540
      MaxLength       =   10
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1965
   End
   Begin VB.TextBox Txt4 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   10  '한글 
      Left            =   8520
      MaxLength       =   10
      TabIndex        =   21
      Text            =   "거주동"
      Top             =   16320
      Width           =   2250
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      ItemData        =   "Jung_New.frx":F625
      Left            =   8535
      List            =   "Jung_New.frx":F627
      Style           =   2  '드롭다운 목록
      TabIndex        =   20
      Top             =   15615
      Width           =   2220
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      ItemData        =   "Jung_New.frx":F629
      Left            =   2535
      List            =   "Jung_New.frx":F62B
      Style           =   2  '드롭다운 목록
      TabIndex        =   19
      Top             =   15540
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Txt8 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9270
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "전화번호"
      Top             =   10770
      Width           =   2925
   End
   Begin VB.TextBox Txt7 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   10  '한글 
      Left            =   22125
      MaxLength       =   10
      TabIndex        =   11
      Text            =   "색상"
      Top             =   9870
      Width           =   2265
   End
   Begin VB.TextBox Txt6 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   10  '한글 
      Left            =   9270
      MaxLength       =   15
      TabIndex        =   3
      Text            =   "차종"
      Top             =   10308
      Width           =   2925
   End
   Begin VB.TextBox Txt5 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   10  '한글 
      Left            =   8505
      MaxLength       =   15
      TabIndex        =   18
      Text            =   "거주호수"
      Top             =   16755
      Width           =   2265
   End
   Begin Threed.SSCommand cmd_1 
      Height          =   735
      Index           =   0
      Left            =   14010
      TabIndex        =   10
      Top             =   12735
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "신규 저장"
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
      Picture         =   "Jung_New.frx":F62D
   End
   Begin Threed.SSCommand cmd_bt2 
      Height          =   750
      Left            =   19320
      TabIndex        =   22
      Top             =   15570
      Visible         =   0   'False
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "삭제"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "Jung_New.frx":F97E
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Index           =   0
      Left            =   14520
      TabIndex        =   5
      Top             =   8955
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#######"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Index           =   1
      Left            =   14520
      TabIndex        =   6
      Top             =   9420
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Index           =   2
      Left            =   14520
      TabIndex        =   7
      Top             =   9870
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Index           =   3
      Left            =   14520
      TabIndex        =   8
      Top             =   10320
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   825
      TabIndex        =   25
      Top             =   11145
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   53936128
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand cmd_1 
      Height          =   735
      Index           =   1
      Left            =   17190
      TabIndex        =   13
      Top             =   12735
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1296
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
      Picture         =   "Jung_New.frx":FCCF
   End
   Begin Threed.SSCommand cmd_1 
      Height          =   735
      Index           =   3
      Left            =   15600
      TabIndex        =   12
      Top             =   12735
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "수 정"
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
      Picture         =   "Jung_New.frx":10020
   End
   Begin Threed.SSCommand cmd_1 
      Height          =   735
      Index           =   2
      Left            =   12420
      TabIndex        =   14
      Top             =   12735
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "입력창 초기화"
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
      Picture         =   "Jung_New.frx":10371
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   19950
      Top             =   210
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ParkHost"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DataJung 
      Height          =   375
      Left            =   19950
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ParkHost"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   690
      Left            =   17100
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   690
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2725
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "닫 기"
      ForeColor       =   12648447
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
      Picture         =   "Jung_New.frx":106C2
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   3525
      Left            =   360
      TabIndex        =   47
      Top             =   3540
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   6218
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   720
      Index           =   0
      Left            =   8310
      TabIndex        =   50
      Top             =   12720
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1270
      _StockProps     =   78
      Caption         =   "엑셀저장"
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
      Picture         =   "Jung_New.frx":10A13
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   720
      Index           =   1
      Left            =   9900
      TabIndex        =   54
      Top             =   12720
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1270
      _StockProps     =   78
      Caption         =   "결  제"
      ForeColor       =   65535
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
      Picture         =   "Jung_New.frx":10D64
   End
   Begin Threed.SSCommand cmd_Option 
      Height          =   570
      Index           =   0
      Left            =   5040
      TabIndex        =   58
      Top             =   9210
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "검 색"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "Jung_New.frx":110B5
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   810
      TabIndex        =   59
      Top             =   8490
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   53936128
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   3855
      TabIndex        =   60
      Top             =   8490
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   53936128
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand cmd_Option 
      Height          =   570
      Index           =   1
      Left            =   3390
      TabIndex        =   61
      Top             =   9210
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "엑셀저장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "Jung_New.frx":11406
   End
   Begin Threed.SSCommand cmd_Option 
      Height          =   570
      Index           =   2
      Left            =   5040
      TabIndex        =   64
      Top             =   11850
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "검 색"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "Jung_New.frx":11757
   End
   Begin Threed.SSCommand cmd_Option 
      Height          =   570
      Index           =   3
      Left            =   1740
      TabIndex        =   65
      Top             =   11850
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "삭 제"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "Jung_New.frx":11AA8
   End
   Begin Threed.SSCommand cmd_Option 
      Height          =   570
      Index           =   4
      Left            =   3390
      TabIndex        =   68
      Top             =   11850
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "엑셀저장"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "Jung_New.frx":11DF9
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "구 분 별 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   15300
      TabIndex        =   70
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '투명
      Caption         =   "간 편 검 색"
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
      Left            =   600
      TabIndex        =   67
      Top             =   2070
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "부터"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   0
      Left            =   3060
      TabIndex        =   63
      Top             =   8550
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "까지"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   1
      Left            =   6090
      TabIndex        =   62
      Top             =   8550
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   6435
      Index           =   1
      Left            =   7260
      Top             =   7530
      Width           =   11655
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "기간초과 자료 조회"
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
      Height          =   495
      Index           =   5
      Left            =   570
      TabIndex        =   57
      Top             =   10500
      Width           =   5055
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "등록일자별 조회"
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
      Height          =   495
      Index           =   4
      Left            =   570
      TabIndex        =   56
      Top             =   7800
      Width           =   5055
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "정기권 등록 / 관리"
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
      Height          =   495
      Index           =   3
      Left            =   7860
      TabIndex        =   55
      Top             =   7800
      Width           =   5085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "정렬 방법 : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   15120
      TabIndex        =   53
      Top             =   3060
      Width           =   1425
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "구     분 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   8100
      TabIndex        =   49
      Top             =   9915
      Width           =   885
   End
   Begin VB.Label Lbl_left 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "차     번 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   8100
      TabIndex        =   48
      Top             =   8565
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "정기권 관리 / 등록"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   26.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   1410
      TabIndex        =   46
      Top             =   720
      Width           =   4110
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "  등록 건수 :"
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
      Height          =   360
      Index           =   15
      Left            =   8805
      TabIndex        =   45
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Label LblRecordCount 
      BackColor       =   &H00000000&
      BackStyle       =   0  '투명
      Caption         =   "123"
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
      Left            =   10620
      TabIndex        =   44
      Top             =   3060
      Width           =   4200
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "차량번호 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   4500
      TabIndex        =   43
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "이   름 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   8340
      TabIndex        =   42
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Lbl_search 
      BackStyle       =   0  '투명
      Caption         =   "소 속 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   11790
      TabIndex        =   41
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Lbl_left 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "차량번호 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   8100
      TabIndex        =   40
      Top             =   8985
      Width           =   1035
   End
   Begin VB.Label Lbl_left 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "이     름 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   8100
      TabIndex        =   39
      Top             =   9435
      Width           =   885
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "분     류 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   6630
      TabIndex        =   38
      Top             =   15615
      Width           =   1695
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "소속구분 :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   630
      TabIndex        =   37
      Top             =   15525
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "거주  호 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   6600
      TabIndex        =   36
      Top             =   16755
      Width           =   1665
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "소     속 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   8100
      TabIndex        =   35
      Top             =   10365
      Width           =   885
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "색     상 :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   5
      Left            =   630
      TabIndex        =   34
      Top             =   16005
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Lbl_center 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "전화번호 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   8100
      TabIndex        =   33
      Top             =   10815
      Width           =   1035
   End
   Begin VB.Label Lbl_right 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "월정요금 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   13305
      TabIndex        =   32
      Top             =   8985
      Width           =   1035
   End
   Begin VB.Label Lbl_right 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "발 급 일 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   13305
      TabIndex        =   31
      Top             =   9465
      Width           =   930
   End
   Begin VB.Label Lbl_right 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "시 작 일 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   13305
      TabIndex        =   30
      Top             =   9915
      Width           =   930
   End
   Begin VB.Label Lbl_right 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "종 료 일 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   13305
      TabIndex        =   29
      Top             =   10365
      Width           =   930
   End
   Begin VB.Label Lbl_right 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "비     고 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   13305
      TabIndex        =   28
      Top             =   10845
      Width           =   885
   End
   Begin VB.Label Lbl_under 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "자료관리 :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   12120
      TabIndex        =   27
      Top             =   15765
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Lbl_under 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "기준  등록기간 만료 자료"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   26
      Top             =   11205
      Width           =   2490
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   6435
      Index           =   0
      Left            =   270
      Top             =   7530
      Width           =   6855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   270
      Top             =   1770
      Width           =   18675
   End
End
Attribute VB_Name = "Jung_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyText(1 To 8) As New clsText
Dim DataField_Enabled As Boolean
Dim Save_TagNum As String
Dim CAR_NO_TMP As String

Public Sub Clear_Field()
Dim i As Integer
Dim tmp As String

Txt1.Text = ""
Txt2.Text = ""
Txt3.Text = ""
Txt4.Text = ""
Txt5.Text = ""
Txt6.Text = ""
Txt7.Text = ""
Txt8.Text = ""
Txt9.Text = ""

Text18.Text = ""
Text19.Text = ""

MaskEdBox1(0).Text = "0"
MaskEdBox1(1).Text = Format(Now, "yyyy-mm-dd")
MaskEdBox1(2).Text = Format(Now, "yyyy-mm-dd")
'tmp = Format(DateAdd("m", 1, Now), "yyyy-mm-01")
MaskEdBox1(3).Text = Format(DateAdd("m", 1, Now), "yyyy-mm-dd")

cmd_Option(1).Enabled = False
cmd_Option(3).Enabled = False
cmd_Option(4).Enabled = False
SSCommand1(0).Enabled = True

'cmb_Gubun.ListIndex = 0

'On Error Resume Next

'Txt2.SetFocus
'Adodc1.Refresh
'Adodc1.Recordset.MoveLast
'LblRecordCount.Caption = Adodc1.Recordset.RecordCount

End Sub


'Sub DataBaseToField()
'Dim i As Integer
''On Error Resume Next
'Dim Cnt1 As Integer
'Dim Cnt2 As Integer
'Dim Cnt3 As Integer
'
'Txt1.Text = Right(DataJung.Recordset!차량번호, 4) & ""
'Txt2.Text = DataJung.Recordset!차량번호 & ""
'Txt3.Text = DataJung.Recordset!이름 & ""
''성훈
''Txt4.Text = DataJung.Recordset!교직원번호 & ""
'
''Txt7.Text = DataJung.Recordset!구분 & ""
'cmb_Gubun.Text = DataJung.Recordset!구분 & ""
'
''Txt4.Text = DataJung.Recordset!거주동 & ""
''Txt5.Text = DataJung.Recordset!거주호수 & ""
'Txt6.Text = DataJung.Recordset!차종 & ""
''Txt7.Text = DataJung.Recordset!색상 & ""
'Txt8.Text = DataJung.Recordset!전화번호 & ""
''Txt9.Text = DataJung.Recordset!비고 & ""
'MaskEdBox1(0).Text = DataJung.Recordset!월정요금
'MaskEdBox1(1).Text = DataJung.Recordset!발급일 & ""
'MaskEdBox1(2).Text = DataJung.Recordset!시작일 & ""
'MaskEdBox1(3).Text = DataJung.Recordset!종료일 & ""
'
'End Sub

Sub Search_Record()
Dim rs As Recordset
Dim SQL_SEARCH As String
Dim itmX As ListItem
Dim INDEX_NO As Integer

SQL_SEARCH = "SELECT * From regcar WHERE 차량번호 = '" & Txt2.Text & "'"
'Debug.Print SQL_SEARCH

Set rs = New ADODB.Recordset
rs.Open SQL_SEARCH, adoConn

If (rs.RecordCount <> 0) Then
    CAR_NO_TMP = rs!차량번호
    
    Txt1.Text = Right(rs!차량번호, 4) & ""
    Txt2.Text = rs!차량번호 & ""
    Txt3.Text = rs!이름 & ""
    'Txt4.Text = rs!거주동 & ""
    'Txt5.Text = rs!거주호수 & ""
    Txt6.Text = rs!차종 & ""
    'Txt7.Text = rs!구분 & ""
    cmb_Gubun.Text = rs!구분 & ""
    Txt8.Text = rs!전화번호 & ""
    MaskEdBox1(0).Text = rs!월정요금 & ""
    MaskEdBox1(1).Text = rs!발급일 & ""
    MaskEdBox1(2).Text = rs!시작일 & ""
    MaskEdBox1(3).Text = rs!종료일 & ""
    Txt9.Text = rs!비고 & ""
    DataField_Enabled = True
Else

End If

Set rs = Nothing

End Sub


Sub Insert_Record()
Dim i As Integer
Dim Cnt As Integer
Dim tmp As String

Dim rs_COUNT As Recordset
Dim rs As Recordset
Dim SQL_COUNT As String
Dim Glo_Reg_Qry As String


If (Txt2.Text = "") Then
    Txt2.Text = " "
End If
Txt2.Text = MidH(Txt2.Text, 1, 20)

If (Txt3.Text = "") Then
    Txt3.Text = " "
End If
Txt3.Text = MidH(Txt3.Text, 1, 20)

If (Txt6.Text = "") Then
    Txt6.Text = " "
End If
Txt6.Text = MidH(Txt6.Text, 1, 30)

If (Txt8.Text = "") Then
    Txt8.Text = " "
End If
Txt8.Text = MidH(Txt8.Text, 1, 20)

If (Txt9.Text = "") Then
    Txt9.Text = " "
End If
Txt9.Text = MidH(Txt9.Text, 1, 100)

If (MaskEdBox1(0).Text = "") Then
    MaskEdBox1(0).Text = "0"
End If
MaskEdBox1(0).Text = Val(MaskEdBox1(0).Text)

If (Txt1.Text = "") Then '신규등록
    'INSERT
    adoConn.Execute "INSERT INTO regcar (차량번호, 차번, 차종, 구분, 이름, 전화번호, 월정요금, 비고, 발급일, 발급시간, 시작일, 종료일) VALUES ('" & Txt2.Text & "', '" & Right(Txt2.Text, 4) & "','" & Txt6.Text & "', '" & cmb_Gubun.Text & "', '" & Txt3.Text & " ', '" & Txt8.Text & " ', '" & MaskEdBox1(0).Text & " ', '" & Txt9.Text & " ','" & Format(Now, "YYYY-MM-DD") & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '" & MaskEdBox1(2).Text & "', '" & MaskEdBox1(3).Text & "')"
    'Debug.Print Glo_Reg_Qry
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt2 & "    차량등록 완료", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt2 & "    차량등록 완료")
    If (MaskEdBox1(0) <> "0") Then
        '대화상자 처리해야됨...!!!
        MBox.Label3.Caption = Txt2.Text & vbCrLf & MaskEdBox1(0).Text & "원"
        MBox.Label3.FontSize = 20
        MBox.Label1.Caption = "위 차량의 차량결제를 등록합니다." & vbCrLf & " 결제처리 하시겠습니까?"
        MBox.Label2.Caption = "차량결제 정보 등록"
        MBox.Show 1
        If (Glo_MsgRet = True) Then
            '변경된 날짜 수정하고
            adoConn.Execute "UPDATE regcar SET 종료일 = '" & Format(MaskEdBox1(3), "YYYY-MM-DD") & "' WHERE 차량번호 = '" & Txt2 & "'"
            '결제 내역 저장
            adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & Txt2 & "', '" & Txt6 & "', '" & cmb_Gubun.Text & "', '" & MaskEdBox1(0).Text & "', '" & Txt3 & "', '" & Txt8 & "', '', '', '" & Format(MaskEdBox1(2), "YYYYMMDD") & "', '" & Format(MaskEdBox1(3), "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
            'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt2 & "    " & MaskEdBox1(0).Text & "원    차량결제 완료", 0
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt2 & "    " & MaskEdBox1(0).Text & "원    차량결제 완료")
        End If
    End If
Else
    adoConn.Execute "UPDATE regcar SET 차량번호 = '" & Txt2.Text & "', 차번 = '" & Right(Txt2.Text, 4) & "', 차종 = '" & Txt6.Text & "', 구분 = '" & cmb_Gubun.Text & "', 이름 = '" & Txt3.Text & "', 전화번호 = '" & Txt8.Text & "', 월정요금 = '" & MaskEdBox1(0).Text & "', 비고 = '" & Txt9.Text & "', 발급시간 = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "',  시작일 = '" & MaskEdBox1(2).Text & "', 종료일 = '" & MaskEdBox1(3).Text & "' WHERE 차량번호 = '" & CAR_NO_TMP & "'"
    'Debug.Print Glo_Reg_Qry
End If

'Set rs = New ADODB.Recordset
'rs.Open Glo_Reg_Qry, adoConn
'Set rs = Nothing

Call ListView_REG_Draw
Call ListView_REG_SQL

'On Error Resume Next
'
'If (Err = 3022) Then
'    Msg_Box.Label2.Caption = "데이터 베이스 오류"
'    Msg_Box.Label1.Caption = "중복된 태그 번호를 허용하지않습니다."
'    Msg_Box.Show 1
'End If

Call Clear_Field

End Sub

Private Sub cmb_Gubun_KeyPress(KeyAscii As Integer)
If (KeyAscii = vbKeyReturn) Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If
End Sub

Private Sub cmd_1_Click(index As Integer)

Dim i As Integer
Dim CAR_NO As String

Select Case index
           Case 0           '신규 저장
                If (Txt1.Text = "") Then
                    If (Data_Error_Check = False) Then
                        Msg_Box.Label2.Caption = "필드 입력 오류"
                        Msg_Box.Label1.Caption = "중요한 항목을 입력하지 않았습니다."
                        Msg_Box.Show 1
                    Else
                        Call Insert_Record
                        'DataJung.Refresh
                        'Call Clear_Field
                        DataField_Enabled = False
                    End If
                Else
                    Msg_Box.Label2.Caption = "신규 데이터 입력 오류"
                        Msg_Box.Label1.Caption = "신규 데이터가 아닙니다. 다시 한번 확인하세요."
                        Msg_Box.Show 1
                End If
                
                '성훈
                Call Clear_Field
                
                Call ListView_REG_Draw
                Call ListView_REG_SQL
                
                
           Case 1           '삭제
                CAR_NO = Txt2.Text
                If (DataField_Enabled = False) Then
                    Exit Sub
                End If
                
                If (CAR_NO_TMP <> Txt2.Text) Then
                    Msg_Box.Label2.Caption = "데이터 선택 오류"
                    Msg_Box.Label1.Caption = "삭제할 데이터를 다시 선택해 주십시요."
                    Msg_Box.Show 1
                    
                    Exit Sub
                End If
                MBox.Label3.Caption = Txt2.Text
                MBox.Label1.Caption = "위 차량의 정기권 자료를 삭제합니다. 삭제 하시겠습니까?"
                MBox.Label2.Caption = "정기권 자료 삭제"
                MBox.Show 1
                
                If (Glo_MsgRet = True) Then
                    Call Delete_Record
                Else
                  
                End If
                '성훈
                Call Clear_Field
                Txt2.SetFocus
                
           Case 2           '신규입력 초기화
                
                Call Clear_Field
                Glo_Reg_Qry = "Select * From regcar"
                DataField_Enabled = False
                Call ListView_REG_Draw
                Call ListView_REG_SQL
           
           Case 3           '수정
                If (Txt1.Text = "") Then
                    Msg_Box.Label2.Caption = "필드 오류"
                    Msg_Box.Label1.Caption = "신규 데이터 입니다. 다시 확인 하세요."
                    Msg_Box.Show 1
                Else
                    CAR_NO = Right(Txt2.Text, 4)
                    If (Txt1.Text = CAR_NO) Then
                          If (Data_Error_Check = False) Then
                              Msg_Box.Label2.Caption = "필드 입력 오류"
                              Msg_Box.Label1.Caption = "중요한 항목을 누락 또는 잘못 입력하였습니다."
                              Msg_Box.Show 1
                          Else
                              MBox.Label1.Caption = "선택하신 정기권 자료가 변경됩니다. 자료를 수정 하시겠습니까?"
                              MBox.Label2.Caption = "정기권 자료 수정"
                              MBox.Show 1
                                                     
                              If (Glo_MsgRet = True) Then
                                  If (DataField_Enabled = True) Then
                                      Call Insert_Record
                                      'Call Clear_Field
                                      DataField_Enabled = False
                                  End If
                                  Txt2.SetFocus
                              Else
                        
                              End If
                          End If
                    Else
                          MBox.Label1.Caption = "정기권 자료의 차량번호가 변경되었습니다. 자료를 수정 하시겠습니까?"
                          MBox.Label2.Caption = "정기권 자료 수정"
                          MBox.Show 1

                          If (Glo_MsgRet = False) Then
                                Exit Sub
                          End If

                          If (Data_Error_Check = False) Then
                                  Msg_Box.Label2.Caption = "필드 입력 오류"
                                  Msg_Box.Label1.Caption = "중요한 항목을 입력하지 않았습니다."
                                  Msg_Box.Show 1
                          Else
                              MBox.Label1.Caption = "선택하신 정기권 자료가 변경됩니다. 자료를 수정 하시겠습니까?"
                              MBox.Label2.Caption = "정기권 자료 수정"
                              MBox.Show 1

                              If (Glo_MsgRet = True) Then
                                  If (DataField_Enabled = True) Then
                                      Call Insert_Record
                                      'Call Clear_Field
                                      DataField_Enabled = False
                                  End If
                                  Txt2.SetFocus
                              Else
                                  Exit Sub
                              End If
                          End If
                     End If
                    
                     '성훈
                     Call Clear_Field
                
                End If
           
                Call ListView_REG_Draw
                Call ListView_REG_SQL
           
'           Case 4           '종료
'
'                Unload Me
'                Exit Sub
           
End Select

'On Error Resume Next
'Adodc1.Recordset.MoveLast
'LblRecordCount.Caption = Adodc1.Recordset.RecordCount
End Sub

Sub Delete_Record()
Dim rs As Recordset
Dim SQL_DELETE As String
        
SQL_DELETE = "DELETE FROM regcar WHERE 차량번호 = '" & Txt2.Text & "' AND 이름 = '" & Txt3.Text & "'"
'Debug.Print SQL_DELETE

Set rs = New ADODB.Recordset
rs.Open SQL_DELETE, adoConn
Set rs = Nothing

DataField_Enabled = False

Call ListView_REG_Draw
Call ListView_REG_SQL

End Sub

Private Sub cmd_bt2_Click()
Dim msg, Style, Title, Response
Dim Ret As Boolean

msg = "종료일이 " & Format(DTPicker1.value, "yyyy-mm-dd") & " 일 이전 정기권 자료를 삭제하게됩니다." & Chr$(13) & Chr$(10) & "자료를 삭제하면 복구 할 수 없습니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "삭제 하시겠습니까?"
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = "Parking System"

Response = MsgBox(msg, Style, Title)

If Response = vbYes Then
    adoConn.Execute "DELETE FROM regcar WHERE 종료일 < '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'"
    Call Form_Activate
    Call Err_doc("호스트 : 종료일 " & Format(DTPicker1.value, "yyyy-mm-dd") & " 일 이전 정기권 자료 삭제")
End If

End Sub

'성훈
'필수 입력 데이터 확인
Private Function Data_Error_Check()
Dim Error_Flag
Error_Flag = True

If Not ((LenH(Txt2.Text) = 11) Or (LenH(Txt2.Text) = 12) Or (LenH(Txt2.Text) = 8)) Then
    Error_Flag = False
End If

If (LenH(Txt2.Text) = 0) Then
    Error_Flag = False
End If
        
If Not (IsNumeric(Right(Txt2.Text, 4))) Then
    Error_Flag = False
End If
  
If (Len(Txt3.Text) = 0) Then
    Error_Flag = False
End If

'성훈
'If Not (IsNumeric(Txt4.Text)) Then
'    Error_Flag = False
'End If
'
'If Not (IsNumeric(Txt5.Text)) Then
'    Error_Flag = False
'End If

If (Len(MaskEdBox1(0).Text) = 0) Then
    Error_Flag = False
End If

'Debug.Print MaskEdBox1(1).Text

'성훈 ???
If (IsDate(MaskEdBox1(1).Text) = False) Then
    Error_Flag = False
End If
If (IsDate(MaskEdBox1(2).Text) = False) Then
    Error_Flag = False
End If
'If (IsDate(MaskEdBox1(3).Text) = False) Then
'    Error_Flag = False
'End If

Data_Error_Check = Error_Flag

End Function

Private Sub cmd_Month_Click()

MaskEdBox1(3).Text = DateAdd("m", 1, MaskEdBox1(3).Text)

End Sub

Private Sub cmd_Option_Click(index As Integer)
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim sql_str As String
Dim Sort_Order As String
Dim msg, Style, Title, Response
Dim Ret As Boolean

Select Case index
    Case 0
        Me.MousePointer = 11
        Call Clear_Field
        '쿼리 구성
        sql_str = "SELECT * FROM regcar WHERE (발급시간 >= '" & Format(DTPicker2, "yyyy-mm-dd") & " 00:00:00') AND (발급시간 <= '" & Format(DTPicker3, "yyyy-mm-dd") & " 23:59:59')"
        'Debug.Print sql_str
        Glo_Reg_Qry = sql_str
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Me.MousePointer = 0
        cmd_Option(1).Enabled = True
        cmd_Option(3).Enabled = False
        cmd_Option(4).Enabled = False
        SSCommand1(0).Enabled = False
        Exit Sub
    
    Case 1
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_등록일자별 검색내역" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "등록일자별 검색내역")
        cmd_Option(1).Enabled = False
        cmd_Option(3).Enabled = False
        cmd_Option(4).Enabled = False
        SSCommand1(0).Enabled = True
        Exit Sub
    
    Case 2
        Me.MousePointer = 11
        Call Clear_Field
        '쿼리 구성
        sql_str = "SELECT * FROM regcar WHERE 종료일 < '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'"
        'Debug.Print sql_str
        Glo_Reg_Qry = sql_str
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Me.MousePointer = 0
        cmd_Option(1).Enabled = False
        cmd_Option(3).Enabled = True
        cmd_Option(4).Enabled = True
        SSCommand1(0).Enabled = False
        Exit Sub
    
    Case 3
        msg = "종료일이 " & Format(DTPicker1.value, "yyyy-mm-dd") & " 일 기준으로 이전 정기권 자료를 삭제하게됩니다." & Chr$(13) & Chr$(10) & "자료를 삭제하면 복구 할 수 없습니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "삭제 하시겠습니까?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Parking System"
        Response = MsgBox(msg, Style, Title)
        If Response = vbYes Then
            adoConn.Execute "DELETE FROM regcar WHERE 종료일 < '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'"
            'Call Form_Activate
            Call Err_doc("호스트 : 종료일 " & Format(DTPicker1.value, "yyyy-mm-dd") & " 일 이전 정기권 자료 삭제")
        End If
        Call cmd_1_Click(2)
        Exit Sub
    
    Case 4
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_기간초과자료 검색내역" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "기간초과자료 검색내역")
        cmd_Option(1).Enabled = False
        cmd_Option(3).Enabled = False
        cmd_Option(4).Enabled = False
        SSCommand1(0).Enabled = True
        Exit Sub
End Select

End Sub

Private Sub Combo1_Click(index As Integer)
Dim i As Integer
If (index = 1) Then
    Combo1(2).Clear
    Select Case Combo1(1).ListIndex
           Case 0
                For i = 0 To 10
                    Combo1(2).AddItem kyo_str(i)
                Next i
           Case 1
                Combo1(2).AddItem kyo_str(11)
           Case 2
                For i = 12 To 17
                    Combo1(2).AddItem kyo_str(i)
                Next i
           Case 3
                For i = 18 To 21
                    Combo1(2).AddItem kyo_str(i)
                Next i
           Case 4
                Combo1(2).AddItem kyo_str(22)
           Case 5
                Combo1(2).AddItem kyo_str(23)
           Case 6
                Combo1(2).AddItem kyo_str(24)
           Case 7
                Combo1(2).AddItem kyo_str(25)
           Case 8
                For i = 26 To 32
                    Combo1(2).AddItem kyo_str(i)
                Next i
           Case 9
                Combo1(2).AddItem kyo_str(33)
    End Select
    Combo1(2).ListIndex = 0
Else
End If

End Sub

'정렬방법 콤보박스
Private Sub Combo2_Click()
'Glo_Reg_Qry = Glo_Reg_Qry & " ORDER BY " & Combo2.Text & " ASC"
Call Clear_Field
Call ListView_REG_Draw
Call ListView_REG_SQL
Glo_cmd_menu_index = 99
DTPicker1.value = Format(DateAdd("m", -1, Now), "yyyy-mm-dd")
End Sub

Private Sub Command1_Click()
Me.MousePointer = 11
JungList1.Show 1
Me.MousePointer = 0
End Sub

Private Sub Form_Activate()
'Adodc1.ConnectionString = AdoConn_Str
'DataJung.ConnectionString = AdoConn_Str
'Adodc1.RecordSource = "select * from regcar"
'DataJung.RecordSource = "select * from regcar"
'Adodc1.Refresh
'DataJung.Refresh
'If (Adodc1.Recordset.RecordCount <> 0) Then
'    Adodc1.Recordset.MoveLast
'End If
'LblRecordCount.Caption = Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim SQL As String

Dim rs As Recordset
Dim QRY As String
Dim Column_to_size As Integer

'Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
'Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
Left = 0
Top = 0

Glo_Reg_Qry = "Select * From regcar"
DataField_Enabled = False

Set MyText(1).MyText = Me.Txt1
Set MyText(2).MyText = Me.Txt2
Set MyText(3).MyText = Me.Txt3
Set MyText(4).MyText = Me.Txt4
Set MyText(5).MyText = Me.Txt5
Set MyText(6).MyText = Me.Txt6
Set MyText(7).MyText = Me.Txt7
Set MyText(8).MyText = Me.Txt8

cmb_Gubun.ListIndex = 0

Call Clear_Field
Combo2.ListIndex = 0
'Call ListView_REG_Draw
'Call ListView_REG_SQL

Glo_cmd_menu_index = 99
DTPicker1.value = Format(DateAdd("m", -1, Now), "yyyy-mm-dd")
DTPicker2.value = Format(DateAdd("m", -1, Now), "yyyy-mm-dd")
DTPicker3.value = Format(DateAdd("m", -1, Now), "yyyy-mm-dd")
End Sub


Public Sub ListView_REG_SQL()
Dim rs_REG As Recordset
Dim QRY As String
Dim itmX As ListItem
Dim INDEX_NO As Integer
Dim str As String

INDEX_NO = 1

Select Case Combo2.ListIndex
    Case 0
        str = "차량번호"
    Case 1
        str = "구분"
    Case 2
        str = "이름"
    Case 3
        str = "종료일"
    Case 4
        str = "차종"
End Select
QRY = Glo_Reg_Qry & " ORDER BY " & str & " ASC"

'리스트 표시
'List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "     " & Glo_Reg_Qry, 0
Set rs_REG = New ADODB.Recordset
rs_REG.Open QRY, adoConn
LblRecordCount = rs_REG.RecordCount

Do While Not (rs_REG.EOF)
    Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
    'itmX.SubItems(1) = "" & rs_REG!PART_NAME
    'itmX.SubItems(1) = "" & rs_REG!차번
    itmX.SubItems(1) = "" & rs_REG!차량번호
    itmX.SubItems(2) = "" & rs_REG!이름
    itmX.SubItems(3) = "" & rs_REG!전화번호
    itmX.SubItems(4) = "" & rs_REG!구분
    itmX.SubItems(5) = "" & rs_REG!차종
    itmX.SubItems(6) = "" & rs_REG!발급일
    itmX.SubItems(7) = "" & rs_REG!발급시간
    itmX.SubItems(8) = "" & rs_REG!시작일
    itmX.SubItems(9) = "" & rs_REG!종료일
    itmX.SubItems(10) = "" & rs_REG!월정요금
    itmX.SubItems(11) = "" & rs_REG!비고
    'itmX.SubItems(12) = "" & rs_REG!CAR_OBJECT
    rs_REG.MoveNext
    INDEX_NO = INDEX_NO + 1
Loop

Set rs_REG = Nothing

If Glo_Index > 1 Then
    ListView_REG.ListItems(Glo_Index).Selected = True
End If

End Sub


Public Sub ListView_REG_Draw()
Dim Column_to_size As Integer

With Me
    
    Call ListViewExtended(.ListView_REG)
    
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    
    .ListView_REG.ColumnHeaders.Add , , " No  "
    '.ListView_REG.ColumnHeaders.Add , , " 사 업 부            "
    '.ListView_REG.ColumnHeaders.Add , , " 차  번     "
    .ListView_REG.ColumnHeaders.Add , , " 차량번호          "
    .ListView_REG.ColumnHeaders.Add , , " 이  름            "
    .ListView_REG.ColumnHeaders.Add , , " 전화번호                "
    .ListView_REG.ColumnHeaders.Add , , " 구  분         "
    .ListView_REG.ColumnHeaders.Add , , " 소  속                        "
    .ListView_REG.ColumnHeaders.Add , , " 발 급 일       "
    .ListView_REG.ColumnHeaders.Add , , " Update                        "
    .ListView_REG.ColumnHeaders.Add , , " 시 작 일           "
    .ListView_REG.ColumnHeaders.Add , , " 종 료 일           "
    .ListView_REG.ColumnHeaders.Add , , " 월정요금      "
    .ListView_REG.ColumnHeaders.Add , , " 비  고        "
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
    

End With

End Sub


Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)

ListView_REG.SetFocus

Txt2 = ListView_REG.SelectedItem.SubItems(1)
Glo_Index = ListView_REG.SelectedItem.index

End Sub


Private Sub SSCommand1_Click(index As Integer)
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
    
Select Case index
    Case 0
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_정기권등록현황" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "정기권등록현황")
        Exit Sub
    
    Case 1
        If (Txt1 <> "") Then
            If (MaskEdBox1(0) <> "0") Then
                '대화상자 처리해야됨...!!!
                MBox.Label3.Caption = Txt2.Text & vbCrLf & MaskEdBox1(0).Text & "원"
                MBox.Label3.FontSize = 20
                MBox.Label1.Caption = "위 차량의 차량결제를 등록합니다." & vbCrLf & " 결제처리 하시겠습니까?"
                MBox.Label2.Caption = "차량결제 정보 등록"
                MBox.Show 1
                If (Glo_MsgRet = True) Then
                    '변경된 날짜 수정하고
                    adoConn.Execute "UPDATE regcar SET 종료일 = '" & Format(MaskEdBox1(3), "YYYY-MM-DD") & "' WHERE 차량번호 = '" & Txt2 & "'"
                    '결제내역 저장
                    adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & Txt2 & "', '" & Txt6 & "', '" & cmb_Gubun.Text & "', '" & MaskEdBox1(0).Text & "', '" & Txt3 & "', '" & Txt8 & "', '', '', '" & Format(MaskEdBox1(2), "YYYYMMDD") & "', '" & Format(MaskEdBox1(3), "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt2 & "    " & MaskEdBox1(0).Text & "원    차량결제 완료", 0
                    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt2 & "    " & MaskEdBox1(0).Text & "원    차량결제 완료")
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

End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub

'Private Sub Txt1_Change()
'    'If (Len(Txt1.Text) = Txt1.MaxLength) Then
'    '    Call Search_Record
'    'End If
'End Sub

Private Sub MaskEdBox1_KeyPress(index As Integer, KeyAscii As Integer)
If (KeyAscii = vbKeyReturn) Then
    If index = 3 Then
        cmd_1(0).SetFocus
    Else
        SendKeys "{TAB}", True
    End If
        KeyAscii = 0
End If
End Sub

Private Sub MaskEdBox1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If index = 3 Then
            cmd_1(0).SetFocus
        Else
            SendKeys "{TAB}", True
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        SendKeys "+{TAB}", True
        KeyCode = 0
    End If
End Sub

Private Sub MaskEdBox1_GotFocus(index As Integer)
    MaskEdBox1(index).SelStart = 0
    MaskEdBox1(index).SelLength = Len(MaskEdBox1(index).Text)
End Sub

Private Sub Text18_GotFocus()
    Text18.SelStart = 0
    Text18.SelLength = Len(Text18.Text)
End Sub

Private Sub Text19_GotFocus()
    Text19.SelStart = 0
    Text19.SelLength = Len(Text19.Text)
End Sub

Private Sub Text20_GotFocus()
    Text20.SelStart = 0
    Text20.SelLength = Len(Text20.Text)
End Sub

'소속 검색
Private Sub Text20_KeyPress(KeyAscii As Integer)
Dim Car_Num_Str As String
Dim QRY As String
Dim rs_REG As Recordset
Dim itmX As ListItem
Dim Column_to_size As Integer
Dim INDEX_NO As Integer

On Error GoTo erro_p

If (KeyAscii = 13) Then
    With Me
        Call ListViewExtended(.ListView_REG)
        
        .ListView_REG.View = lvwReport
        .ListView_REG.ListItems.Clear
        .ListView_REG.ColumnHeaders.Clear
        .ListView_REG.ColumnHeaders.Add , , " No  "
        '.ListView_REG.ColumnHeaders.Add , , " 사 업 부            "
        '.ListView_REG.ColumnHeaders.Add , , " 차  번     "
        .ListView_REG.ColumnHeaders.Add , , " 차량번호          "
        .ListView_REG.ColumnHeaders.Add , , " 이  름            "
        .ListView_REG.ColumnHeaders.Add , , " 전화번호                "
        .ListView_REG.ColumnHeaders.Add , , " 구  분         "
        .ListView_REG.ColumnHeaders.Add , , " 소  속                        "
        .ListView_REG.ColumnHeaders.Add , , " 발 급 일       "
        .ListView_REG.ColumnHeaders.Add , , " Update                        "
        .ListView_REG.ColumnHeaders.Add , , " 시 작 일           "
        .ListView_REG.ColumnHeaders.Add , , " 종 료 일           "
        .ListView_REG.ColumnHeaders.Add , , " 월정요금      "
        .ListView_REG.ColumnHeaders.Add , , " 비  고        "
        For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
             SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With
    
    INDEX_NO = 1
    Glo_Reg_Qry = "SELECT * FROM regcar WHERE 차종 = '" & Text20.Text & "'"
    QRY = Glo_Reg_Qry & " order by " & Combo2.Text & "'"
    Set rs_REG = New ADODB.Recordset
    rs_REG.Open Glo_Reg_Qry, adoConn
    LblRecordCount = rs_REG.RecordCount
    ListView_REG.ListItems.Clear
    If (rs_REG.EOF) Then
        LblRecordCount.Caption = " 자료가 존재 하지않습니다.."
    Else
        LblRecordCount.Caption = " " & rs_REG.RecordCount & " 건"
        Do While Not (rs_REG.EOF)
            Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
            'itmX.SubItems(1) = "" & rs_REG!PART_NAME
            'itmX.SubItems(1) = "" & rs_REG!차번
            itmX.SubItems(1) = "" & rs_REG!차량번호
            itmX.SubItems(2) = "" & rs_REG!이름
            itmX.SubItems(3) = "" & rs_REG!전화번호
            itmX.SubItems(4) = "" & rs_REG!구분
            itmX.SubItems(5) = "" & rs_REG!차종
            itmX.SubItems(6) = "" & rs_REG!발급일
            itmX.SubItems(7) = "" & rs_REG!발급시간
            itmX.SubItems(8) = "" & rs_REG!시작일
            itmX.SubItems(9) = "" & rs_REG!종료일
            itmX.SubItems(10) = "" & rs_REG!월정요금
            itmX.SubItems(11) = "" & rs_REG!비고
            'itmX.SubItems(12) = "" & rs_REG!CAR_OBJECT
            rs_REG.MoveNext
            INDEX_NO = INDEX_NO + 1
        Loop
    End If
    Set rs_REG = Nothing
    KeyAscii = 0
    
    Text20 = ""
    Exit Sub

erro_p:
    MsgBox Err.Description
End If

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)

Dim Car_Num_Str As String
Dim QRY As String
Dim rs_REG As Recordset
Dim itmX As ListItem
Dim Column_to_size As Integer
Dim INDEX_NO As Integer
'On Error GoTo erro_p

If (KeyAscii = 13) Then
    If ((Len(Text18) <> 4) Or Not (IsNumeric(Text18))) Then
        MsgBox "차량번호 숫자 네지리를 정확하게 입력하세요!"
        Text18 = ""
        Exit Sub
    End If
    With Me
        Call ListViewExtended(.ListView_REG)
        .ListView_REG.View = lvwReport
        .ListView_REG.ListItems.Clear
        .ListView_REG.ColumnHeaders.Clear
        .ListView_REG.ColumnHeaders.Add , , " No  "
        '.ListView_REG.ColumnHeaders.Add , , " 사 업 부            "
        '.ListView_REG.ColumnHeaders.Add , , " 차  번     "
        .ListView_REG.ColumnHeaders.Add , , " 차량번호          "
        .ListView_REG.ColumnHeaders.Add , , " 이  름            "
        .ListView_REG.ColumnHeaders.Add , , " 전화번호                "
        .ListView_REG.ColumnHeaders.Add , , " 구  분         "
        .ListView_REG.ColumnHeaders.Add , , " 소  속                        "
        .ListView_REG.ColumnHeaders.Add , , " 발 급 일       "
        .ListView_REG.ColumnHeaders.Add , , " Update                        "
        .ListView_REG.ColumnHeaders.Add , , " 시 작 일           "
        .ListView_REG.ColumnHeaders.Add , , " 종 료 일           "
        .ListView_REG.ColumnHeaders.Add , , " 월정요금      "
        .ListView_REG.ColumnHeaders.Add , , " 비  고        "
        For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
             SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With
    
    INDEX_NO = 1
    Glo_Reg_Qry = "SELECT * FROM regcar WHERE 차량번호  Like '%" & Text18.Text & "%'"
    QRY = Glo_Reg_Qry '& " order by " & Combo2.Text & "'"
    Set rs_REG = New ADODB.Recordset
    rs_REG.Open QRY, adoConn
    LblRecordCount = rs_REG.RecordCount
    ListView_REG.ListItems.Clear
    If (rs_REG.EOF) Then
        LblRecordCount.Caption = " 자료가 존재 하지않습니다.."
    Else
        LblRecordCount.Caption = " " & rs_REG.RecordCount & " 건"
        Do While Not (rs_REG.EOF)
            Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
            'itmX.SubItems(1) = "" & rs_REG!PART_NAME
            'itmX.SubItems(1) = "" & rs_REG!차번
            itmX.SubItems(1) = "" & rs_REG!차량번호
            itmX.SubItems(2) = "" & rs_REG!이름
            itmX.SubItems(3) = "" & rs_REG!전화번호
            itmX.SubItems(4) = "" & rs_REG!구분
            itmX.SubItems(5) = "" & rs_REG!차종
            itmX.SubItems(6) = "" & rs_REG!발급일
            itmX.SubItems(7) = "" & rs_REG!발급시간
            itmX.SubItems(8) = "" & rs_REG!시작일
            itmX.SubItems(9) = "" & rs_REG!종료일
            itmX.SubItems(10) = "" & rs_REG!월정요금
            itmX.SubItems(11) = "" & rs_REG!비고
            'itmX.SubItems(12) = "" & rs_REG!CAR_OBJECT
            rs_REG.MoveNext
            INDEX_NO = INDEX_NO + 1
        Loop
        
    End If
    Set rs_REG = Nothing
    KeyAscii = 0
    Text18 = ""
    Exit Sub

erro_p:
    MsgBox Err.Description
End If

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)

Dim Car_Num_Str As String
Dim QRY As String
Dim rs_REG As Recordset
Dim itmX As ListItem
Dim Column_to_size As Integer
Dim INDEX_NO As Integer
'On Error GoTo erro_p

If (KeyAscii = 13) Then
    With Me
        Call ListViewExtended(.ListView_REG)
        
        .ListView_REG.View = lvwReport
        .ListView_REG.ListItems.Clear
        .ListView_REG.ColumnHeaders.Clear
        .ListView_REG.ColumnHeaders.Add , , " No  "
        '.ListView_REG.ColumnHeaders.Add , , " 사 업 부            "
        '.ListView_REG.ColumnHeaders.Add , , " 차  번     "
        .ListView_REG.ColumnHeaders.Add , , " 차량번호          "
        .ListView_REG.ColumnHeaders.Add , , " 이  름            "
        .ListView_REG.ColumnHeaders.Add , , " 전화번호                "
        .ListView_REG.ColumnHeaders.Add , , " 구  분         "
        .ListView_REG.ColumnHeaders.Add , , " 소  속                        "
        .ListView_REG.ColumnHeaders.Add , , " 발 급 일       "
        .ListView_REG.ColumnHeaders.Add , , " Update                        "
        .ListView_REG.ColumnHeaders.Add , , " 시 작 일           "
        .ListView_REG.ColumnHeaders.Add , , " 종 료 일           "
        .ListView_REG.ColumnHeaders.Add , , " 월정요금      "
        .ListView_REG.ColumnHeaders.Add , , " 비  고        "
        For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
             SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With
    
    INDEX_NO = 1
    Glo_Reg_Qry = "SELECT * FROM regcar WHERE 이름  Like '" & Text19.Text & "%'"
    QRY = Glo_Reg_Qry & " order by " & Combo2.Text & "'"
    Set rs_REG = New ADODB.Recordset
    rs_REG.Open Glo_Reg_Qry, adoConn
    LblRecordCount = rs_REG.RecordCount
    ListView_REG.ListItems.Clear
    If (rs_REG.EOF) Then
        LblRecordCount.Caption = " 자료가 존재 하지않습니다.."
    Else
        LblRecordCount.Caption = " " & rs_REG.RecordCount & " 건"
        Do While Not (rs_REG.EOF)
            Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
            'itmX.SubItems(1) = "" & rs_REG!PART_NAME
            'itmX.SubItems(1) = "" & rs_REG!차번
            itmX.SubItems(1) = "" & rs_REG!차량번호
            itmX.SubItems(2) = "" & rs_REG!이름
            itmX.SubItems(3) = "" & rs_REG!전화번호
            itmX.SubItems(4) = "" & rs_REG!구분
            itmX.SubItems(5) = "" & rs_REG!차종
            itmX.SubItems(6) = "" & rs_REG!발급일
            itmX.SubItems(7) = "" & rs_REG!발급시간
            itmX.SubItems(8) = "" & rs_REG!시작일
            itmX.SubItems(9) = "" & rs_REG!종료일
            itmX.SubItems(10) = "" & rs_REG!월정요금
            itmX.SubItems(11) = "" & rs_REG!비고
            'itmX.SubItems(12) = "" & rs_REG!CAR_OBJECT
            rs_REG.MoveNext
            INDEX_NO = INDEX_NO + 1
        Loop
        
    End If
    Set rs_REG = Nothing
    KeyAscii = 0
    
    Text19 = ""
    Exit Sub

erro_p:
    MsgBox Err.Description
End If

End Sub

'구분별 간편검색
Private Sub cmb_SGubun_Click()
Dim Car_Num_Str As String
Dim QRY As String
Dim rs_REG As Recordset
Dim itmX As ListItem
Dim Column_to_size As Integer
Dim INDEX_NO As Integer

On Error GoTo erro_p

With Me
    Call ListViewExtended(.ListView_REG)
    
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No  "
    '.ListView_REG.ColumnHeaders.Add , , " 사 업 부            "
    '.ListView_REG.ColumnHeaders.Add , , " 차  번     "
    .ListView_REG.ColumnHeaders.Add , , " 차량번호          "
    .ListView_REG.ColumnHeaders.Add , , " 이  름            "
    .ListView_REG.ColumnHeaders.Add , , " 전화번호                "
    .ListView_REG.ColumnHeaders.Add , , " 구  분         "
    .ListView_REG.ColumnHeaders.Add , , " 소  속                        "
    .ListView_REG.ColumnHeaders.Add , , " 발 급 일       "
    .ListView_REG.ColumnHeaders.Add , , " Update                        "
    .ListView_REG.ColumnHeaders.Add , , " 시 작 일           "
    .ListView_REG.ColumnHeaders.Add , , " 종 료 일           "
    .ListView_REG.ColumnHeaders.Add , , " 월정요금      "
    .ListView_REG.ColumnHeaders.Add , , " 비  고        "
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With

INDEX_NO = 1
Glo_Reg_Qry = "SELECT * FROM regcar WHERE 구분 = '" & cmb_SGubun & "'"
QRY = Glo_Reg_Qry & " order by " & Combo2.Text & "'"
Set rs_REG = New ADODB.Recordset
rs_REG.Open Glo_Reg_Qry, adoConn
LblRecordCount = rs_REG.RecordCount
ListView_REG.ListItems.Clear
If (rs_REG.EOF) Then
    LblRecordCount.Caption = " 자료가 존재 하지않습니다.."
Else
    LblRecordCount.Caption = " " & rs_REG.RecordCount & " 건"
    Do While Not (rs_REG.EOF)
        Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
        'itmX.SubItems(1) = "" & rs_REG!PART_NAME
        'itmX.SubItems(1) = "" & rs_REG!차번
        itmX.SubItems(1) = "" & rs_REG!차량번호
        itmX.SubItems(2) = "" & rs_REG!이름
        itmX.SubItems(3) = "" & rs_REG!전화번호
        itmX.SubItems(4) = "" & rs_REG!구분
        itmX.SubItems(5) = "" & rs_REG!차종
        itmX.SubItems(6) = "" & rs_REG!발급일
        itmX.SubItems(7) = "" & rs_REG!발급시간
        itmX.SubItems(8) = "" & rs_REG!시작일
        itmX.SubItems(9) = "" & rs_REG!종료일
        itmX.SubItems(10) = "" & rs_REG!월정요금
        itmX.SubItems(11) = "" & rs_REG!비고
        'itmX.SubItems(12) = "" & rs_REG!CAR_OBJECT
        rs_REG.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    
End If
Set rs_REG = Nothing

'cmb_SGubun.Text = ""

Exit Sub

erro_p:
    MsgBox Err.Description

End Sub





Private Sub Txt2_Change()
    Call Search_Record
    'If (Len(Txt1.Text) = Txt1.MaxLength) Then
    '    'Call Search_Record
    'End If

End Sub


