VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form JIOSch 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "월정입출"
   ClientHeight    =   14730
   ClientLeft      =   2685
   ClientTop       =   1560
   ClientWidth     =   19185
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   24
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "JIOSch.frx":0000
   ScaleHeight     =   14730
   ScaleWidth      =   19185
   Begin Threed.SSCommand Command1 
      Height          =   615
      Left            =   16635
      TabIndex        =   17
      Top             =   7905
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
      Picture         =   "JIOSch.frx":DC0A
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
      Left            =   13110
      MaxLength       =   10
      TabIndex        =   10
      Top             =   7950
      Width           =   3195
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "JIOSch.frx":DF5B
      Left            =   12600
      List            =   "JIOSch.frx":DF74
      Style           =   2  '드롭다운 목록
      TabIndex        =   9
      Top             =   5115
      Width           =   1950
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "JIOSch.frx":DFC6
      Left            =   12600
      List            =   "JIOSch.frx":DFE5
      Style           =   2  '드롭다운 목록
      TabIndex        =   8
      Top             =   5502
      Width           =   1950
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "JIOSch.frx":E049
      Left            =   12600
      List            =   "JIOSch.frx":E05C
      Style           =   2  '드롭다운 목록
      TabIndex        =   7
      Top             =   6663
      Width           =   1950
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      ItemData        =   "JIOSch.frx":E08E
      Left            =   12600
      List            =   "JIOSch.frx":E09B
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   5889
      Width           =   1950
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   690
      Left            =   15540
      TabIndex        =   0
      Top             =   13470
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1217
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
      RoundedCorners  =   0   'False
      Picture         =   "JIOSch.frx":E0B1
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   690
      Left            =   17145
      TabIndex        =   1
      Top             =   13470
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "종 료(&X)"
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
      Picture         =   "JIOSch.frx":E402
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   540
      Index           =   0
      Left            =   22920
      TabIndex        =   4
      Top             =   5280
      Width           =   3225
      _Version        =   65536
      _ExtentX        =   5689
      _ExtentY        =   952
      _StockProps     =   15
      Caption         =   " "
      ForeColor       =   16777215
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1155
      Index           =   2
      Left            =   22950
      TabIndex        =   5
      Top             =   5910
      Visible         =   0   'False
      Width           =   3120
      _Version        =   65536
      _ExtentX        =   5503
      _ExtentY        =   2037
      _StockProps     =   15
      ForeColor       =   65535
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   36
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Enabled         =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   12600
      TabIndex        =   11
      Top             =   7050
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16646144
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   15420
      TabIndex        =   12
      Top             =   7050
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16646144
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   12615
      TabIndex        =   13
      Top             =   6270
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16646146
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   15420
      TabIndex        =   14
      Top             =   6270
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16646146
      CurrentDate     =   36927
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   495
      Index           =   7
      Left            =   11430
      TabIndex        =   15
      Top             =   840
      Width           =   3420
      _Version        =   65536
      _ExtentX        =   6032
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "  검색 건수"
      ForeColor       =   16777215
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
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
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1590
         TabIndex        =   16
         Top             =   90
         Width           =   1275
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3000
      Left            =   420
      TabIndex        =   18
      Top             =   10260
      Width           =   18315
      _ExtentX        =   32306
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
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Threed.SSPanel SSPanel3 
      DataField       =   "CARNO"
      DataSource      =   "Adodc1"
      Height          =   825
      Index           =   1
      Left            =   12420
      TabIndex        =   31
      Top             =   2520
      Width           =   4380
      _Version        =   65536
      _ExtentX        =   7726
      _ExtentY        =   1455
      _StockProps     =   15
      ForeColor       =   16777215
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   24
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
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
      Left            =   11250
      TabIndex        =   30
      Top             =   5100
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "인식상태"
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
      Index           =   1
      Left            =   11250
      TabIndex        =   29
      Top             =   5505
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
      Left            =   11250
      TabIndex        =   28
      Top             =   5880
      Width           =   1125
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
      Left            =   11250
      TabIndex        =   27
      Top             =   6270
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
      Left            =   11250
      TabIndex        =   26
      Top             =   6660
      Width           =   900
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
      Left            =   11250
      TabIndex        =   25
      Top             =   7035
      Width           =   900
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
      Left            =   14700
      TabIndex        =   24
      Top             =   6270
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
      Left            =   14700
      TabIndex        =   23
      Top             =   7080
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
      Left            =   17520
      TabIndex        =   22
      Top             =   6270
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
      Index           =   9
      Left            =   17520
      TabIndex        =   21
      Top             =   7065
      Width           =   450
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
      Left            =   11550
      TabIndex        =   20
      Top             =   8010
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
      Index           =   12
      Left            =   10830
      TabIndex        =   19
      Top             =   2730
      Width           =   1200
   End
   Begin VB.Label Label10 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   420
      TabIndex        =   3
      Top             =   13875
      Visible         =   0   'False
      Width           =   14715
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   420
      TabIndex        =   2
      Top             =   13410
      Visible         =   0   'False
      Width           =   14715
   End
   Begin VB.Image Image3 
      Height          =   7125
      Left            =   420
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   9510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      Height          =   7185
      Left            =   405
      Top             =   1680
      Width           =   9570
   End
End
Attribute VB_Name = "JIOSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim excel_sql_str As String

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

Sort_Order = Combo3.List(Combo3.ListIndex)

If (Combo1.ListIndex = 0) Then
    sql_str = "SELECT * FROM regcarinout WHERE (처리일시>='" & Format(DTPicker1, "yyyymmdd") & Format(DTPicker3, "hhnnss") & "') AND (처리일시<='" & Format(DTPicker2, "yyyymmdd") & Format(DTPicker4, "hhnnss") & "')"
Else
    sql_str = "SELECT * FROM regcarinout WHERE (처리일시>='" & Format(DTPicker1, "yyyymmdd") & Format(DTPicker3, "hhnnss") & "') AND (처리일시<='" & Format(DTPicker2, "yyyymmdd") & Format(DTPicker4, "hhnnss") & "') AND " & "(입출상태='" & Combo1.List(Combo1.ListIndex) & "')"
End If

If (Combo2.ListIndex = 0) Then
Else
    sql_str = sql_str & " AND (인식상태='" & Combo2.List(Combo2.ListIndex) & "')"
End If

Select Case Combo4.ListIndex
    Case 0
    
    Case 1
        sql_str = sql_str & " AND (입출구분 = '0')"
    Case 2
        sql_str = sql_str & " AND (입출구분='1')"
End Select


On Error Resume Next

If (Text1(2).Text = "") Then

Else
    If ((Len(Text1(2)) = 4) And (IsNumeric(Text1(2)))) Then
        sql_str = sql_str & " AND (차량번호 Like '%" & Text1(2).Text & "')"
    Else
        sql_str = sql_str & " AND (차량번호='" & Text1(2).Text & "')"
    End If
End If

'Debug.Print sql_str

Glo_JIOSch = sql_str & " ORDER BY " & Sort_Order

Call ListView_Draw

SSCommand1.Enabled = True
Me.MousePointer = 0

On Error Resume Next

End Sub


Private Sub Form_Load()

'Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
'Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
Left = 0
Top = 0

DTPicker1.value = Now
DTPicker2.value = Now
DTPicker3.value = Format("00:00:00")
DTPicker4.value = Format("23:59:59")

Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")

'오늘날짜 데이터만
Glo_JIOSch = "SELECT * FROM regcarinout WHERE (처리일시 >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (처리일시 <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 00:00:00')"

Call ListView_Draw

Exit Sub

Err_P:
        MsgBox "데이터 베이스 연결실패" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "네트웍 관리자에게 문의 바랍니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "데이터 베이스 연결전에는 자료검색 기능을 수행할수 없습니다.", vbCritical
End Sub

'인쇄
Private Sub SSCommand1_Click()
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
    
tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
tmpFileName = App.Path & "\Excel\" & tmpFileName & "_검색내역" & ".xls"
'Call makeexcel(ListView1, tmpFileName, "차량입출차현황")
Call makeexcel(ListView1, tmpFileName, "검색내역")

Exit Sub

End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub

Public Sub ListView_Draw()
Dim Column_to_size As Integer
Dim rs As Recordset
Dim Qry As String
Dim itmX As ListItem
Dim INDEX_NO As Long
    
    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " 입/츨차 일자 "
    ListView1.ColumnHeaders.Add , , " 시간  "
    ListView1.ColumnHeaders.Add , , " 입출구분 "
    ListView1.ColumnHeaders.Add , , " 입출상태    "
    ListView1.ColumnHeaders.Add , , " 차량번호      "
    ListView1.ColumnHeaders.Add , , " 이 름         "
    ListView1.ColumnHeaders.Add , , " 연락처           "
    ListView1.ColumnHeaders.Add , , " 구 분         "
    ListView1.ColumnHeaders.Add , , " 이미지경로                "
    
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
 
    Set rs = New ADODB.Recordset
    rs.Open Glo_JIOSch, adoConn
    LblRecordCount = rs.RecordCount

    INDEX_NO = 1

    Do While Not (rs.EOF)
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        itmX.SubItems(1) = "" & rs!입차일자
        itmX.SubItems(2) = "" & rs!입차시간
        If rs!입출구분 = 0 Then
            itmX.SubItems(3) = "" & "입차"
        Else
            itmX.SubItems(3) = "" & "출차"
        End If
        itmX.SubItems(4) = "" & rs!입출상태
        itmX.SubItems(5) = "" & rs!차량번호
        itmX.SubItems(6) = "" & rs!이름
        itmX.SubItems(7) = "" & rs!전화번호
        itmX.SubItems(8) = "" & rs!구분
        itmX.SubItems(9) = "" & rs!이미지명
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    INDEX_NO = 0
    Set rs = Nothing

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim Tmp_File As String
Dim image_name As String
Dim ECHO As ICMP_ECHO_REPLY
Dim RemoteIP As String
'On Error Resume Next

SSPanel3(1).Caption = " " & ListView1.SelectedItem.SubItems(5)
RemoteIP = Mid(Trim(ListView1.SelectedItem.SubItems(9)), 3, InStr(3, Trim(ListView1.SelectedItem.SubItems(9)), "\", 1) - 3)

'Ping Test...!!
Call Ping(RemoteIP, ECHO)
If Left$(ECHO.Data, 1) <> Chr$(0) Then
    Tmp_File = Dir(Trim(ListView1.SelectedItem.SubItems(9)))
    If (Tmp_File <> "") Then
        Image3.Picture = LoadPicture(Trim(ListView1.SelectedItem.SubItems(9)))
    Else
        Image3.Picture = LoadPicture(App.Path & "\NoCar.jpg")
    End If
Else
    Image3.Picture = Nothing
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "Ping Test Failure...!!")
    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & CarNum & "  " & Tmp_Path)
End If


End Sub



