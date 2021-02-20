VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form IlIOList 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   ClientHeight    =   14955
   ClientLeft      =   28410
   ClientTop       =   405
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   Picture         =   "IlIOList.frx":0000
   ScaleHeight     =   14955
   ScaleWidth      =   19080
   Begin Threed.SSCommand Command1 
      Height          =   570
      Left            =   11880
      TabIndex        =   9
      Top             =   13710
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "확인"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "IlIOList.frx":1007C
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "IlIOList.frx":103CD
      Left            =   2565
      List            =   "IlIOList.frx":103E0
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   13815
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   10395
      MaxLength       =   4
      TabIndex        =   5
      Top             =   13710
      Width           =   1275
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   780
      Left            =   15900
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   13350
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "인쇄(&P)"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "IlIOList.frx":10414
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   19545
      Top             =   255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\ehwa\SAMSUNG\Report\mlist.rpt"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   1024
      WindowHeight    =   768
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   780
      Left            =   17310
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   13350
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "닫 기(X)"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "IlIOList.frx":10765
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   20190
      Top             =   285
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      DataSourceName  =   ""
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
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "IlIOList.frx":10AB6
      Height          =   10320
      Left            =   345
      TabIndex        =   2
      Top             =   1590
      Width           =   18510
      _ExtentX        =   32650
      _ExtentY        =   18203
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   21
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "영수증번호"
         Caption         =   " 영수증번호"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "주차권번호"
         Caption         =   " 주차권번호"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "발권기번호"
         Caption         =   " 발권기번호"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "출구번호"
         Caption         =   " 출구번호"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "구분"
         Caption         =   " 구 분"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "명칭"
         Caption         =   " 명 칭"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "입차일자"
         Caption         =   " 입차일자"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "입차시간"
         Caption         =   " 입차시간"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "출차일자"
         Caption         =   " 출차일자"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "출차시간"
         Caption         =   " 출차시간"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "주차시간"
         Caption         =   " 주차시간"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "할인시간"
         Caption         =   " 할인시간"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "요금"
         Caption         =   " 요 금"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """\""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   11460
      TabIndex        =   3
      Top             =   885
      Width           =   3210
      _Version        =   65536
      _ExtentX        =   5662
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   "  레코드 건수"
      ForeColor       =   16777215
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
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
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1725
         TabIndex        =   4
         Top             =   75
         Width           =   1275
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   2565
      TabIndex        =   7
      Top             =   13260
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
      Format          =   16646144
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   5310
      TabIndex        =   8
      Top             =   13260
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
      Format          =   16646144
      CurrentDate     =   36927
   End
End
Attribute VB_Name = "IlIOList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Order$
Dim SelectionFormula$

Private Sub Command1_Click()
Dim i As Integer
Me.MousePointer = 11
If (Text1.Text = "") Then
    Data1.RecordSource = "SELECT * FROM ilbaninout WHERE (출차일자>='" & Format(DTPicker1.value, "yyyy-mm-dd") & "') AND (출차일자<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "')"
    SelectionFormula$ = "{ilbaninout.출차일자} >= '" & Format(DTPicker1.value, "yyyy-mm-dd") & "' AND {ilbaninout.출차일자}<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "'"
Else
    Data1.RecordSource = "SELECT * FROM ilbaninout WHERE (출차일자>='" & Format(DTPicker1.value, "yyyy-mm-dd") & "') AND (출차일자<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "') AND (주차권번호 = '" & Text1.Text & "')"
    SelectionFormula$ = "{ilbaninout.출차일자} >= '" & Format(DTPicker1.value, "yyyy-mm-dd") & "' AND {ilbaninout.출차일자}<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "'" & " AND {ilbaninout.주차권번호}='" & Text1.Text & "'"
End If
Data1.Refresh
If (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    LblRecordCount = 0
Else
    Data1.Recordset.MoveLast
    LblRecordCount = Data1.Recordset.RecordCount
End If
SSCommand1.Enabled = True
Me.MousePointer = 0
End Sub



Private Sub Form_Load()
Dim InsSQL As String
Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
Report1.Connect = AdoConn_Str
Report1.Connect = AdoConn_Str
'Data1.ConnectionString = AdoConn_Str
Data1.RecordSource = "SELECT * FROM ilbaninout WHERE 출차일자='" & Format(Now, "yyyy-mm-dd") & "'"
Data1.Refresh
If (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    LblRecordCount = 0
Else
    Data1.Recordset.MoveLast
    LblRecordCount = Data1.Recordset.RecordCount
End If
DTPicker1.value = Now
DTPicker2.value = Now
Combo1.ListIndex = 0
End Sub


Private Sub SSCommand1_Click()
Dim tmp%
 
On Error GoTo PrintReportError
DoEvents
Report1.ReportFileName = Report_Path_Name$ & "tmpilio.rpt"
Report1.Formulas(0) = "StartDate=' 자료검색일 : " & Format(DTPicker1.value, "yyyy년 mm월 dd일") & " ~ " & Format(DTPicker2.value, "yyyy년 mm월 dd일") & "'"

Select Case Combo1.ListIndex
       Case 0
            Report1.SortFields(0) = "+{ilbaninout.주차권번호}"
            Report1.SortFields(1) = "+{ilbaninout.입차일자}"
            Report1.SortFields(2) = "+{ilbaninout.입차시간}"
       Case 1
            Report1.SortFields(0) = "+{ilbaninout.입차일자}"
            Report1.SortFields(1) = "+{ilbaninout.입차시간}"
            Report1.SortFields(2) = "+{ilbaninout.주차권번호}"
       Case 2
            Report1.SortFields(0) = "+{ilbaninout.출차일자}"
            Report1.SortFields(1) = "+{ilbaninout.출차시간}"
            Report1.SortFields(2) = "+{ilbaninout.주차권번호}"
       Case 3
            Report1.SortFields(0) = "+{ilbaninout.명칭}"
            Report1.SortFields(1) = "+{ilbaninout.입차일자}"
            Report1.SortFields(2) = "+{ilbaninout.입차시간}"
       Case 4
            Report1.SortFields(0) = "+{ilbaninout.요금}"
            Report1.SortFields(1) = "+{ilbaninout.입차일자}"
            Report1.SortFields(2) = "+{ilbaninout.입차시간}"
End Select
 
Report1.SelectionFormula = SelectionFormula$
Report1.CopiesToPrinter = 1
Report1.Action = 1
Exit Sub

PrintReportError:
Msg_Box.Label2.Caption = "프린트 작업 오류"
Msg_Box.Label1.Caption = "프린터가 정상작동 하지않습니다."
Msg_Box.Show 1

End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub

