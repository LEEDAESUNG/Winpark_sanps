VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Rptmenu 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "영업 보고"
   ClientHeight    =   6000
   ClientLeft      =   24825
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Rptmenu.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   10050
   Begin Threed.SSCommand Command1 
      Height          =   495
      Index           =   0
      Left            =   14445
      TabIndex        =   8
      Top             =   645
      Width           =   2130
      _Version        =   65536
      _ExtentX        =   3757
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "SSCommand1"
   End
   Begin Threed.SSCommand Command2 
      Height          =   660
      Left            =   8235
      TabIndex        =   7
      Top             =   1350
      Width           =   1380
      _Version        =   65536
      _ExtentX        =   2434
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "검  색"
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
      Picture         =   "Rptmenu.frx":43E5
   End
   Begin VB.Frame Frame2 
      Caption         =   "자료검색"
      Height          =   855
      Left            =   10470
      TabIndex        =   0
      Top             =   2340
      Width           =   9465
      Begin VB.Label Label1 
         Caption         =   "                                                부터                                                  까지"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   1
         Top             =   375
         Width           =   6735
      End
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "Rptmenu.frx":4736
      Left            =   10755
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\WinPark\DATA\Sail.rpt"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   1024
      WindowHeight    =   768
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   10620
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
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "Rptmenu.frx":474A
      Height          =   3075
      Left            =   465
      TabIndex        =   2
      Top             =   2505
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   5424
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "근무자명"
         Caption         =   "  근무자명"
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
         DataField       =   "시작일시"
         Caption         =   "  시작일시"
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
         DataField       =   "종료일시"
         Caption         =   "  종료일시"
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
         DataField       =   "정산소구분"
         Caption         =   "  정산소구분"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            DividerStyle    =   1
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1695.118
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Index           =   0
      Left            =   1335
      TabIndex        =   3
      Top             =   1485
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
      Format          =   16711680
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Index           =   1
      Left            =   4725
      TabIndex        =   4
      Top             =   1485
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
      Format          =   16711680
      CurrentDate     =   36927
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   345
      Left            =   3360
      TabIndex        =   5
      Top             =   1485
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   345
      Left            =   6735
      TabIndex        =   6
      Top             =   1485
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   " "
   End
   Begin Threed.SSCommand Command1 
      Height          =   615
      Index           =   1
      Left            =   7695
      TabIndex        =   9
      Top             =   2700
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3149
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "시간대별 현황"
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
      Picture         =   "Rptmenu.frx":475E
   End
   Begin Threed.SSCommand Command1 
      Height          =   615
      Index           =   2
      Left            =   7695
      TabIndex        =   10
      Top             =   3540
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3149
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "할인항목별 현황"
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
      Picture         =   "Rptmenu.frx":4AAF
   End
   Begin Threed.SSCommand Command1 
      Height          =   615
      Index           =   3
      Left            =   7695
      TabIndex        =   11
      Top             =   4860
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3149
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "종  료"
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
      Picture         =   "Rptmenu.frx":4E00
   End
End
Attribute VB_Name = "Rptmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
Dim QRY As String
Dim rs As ADODB.Recordset


MaskEdBox2.Text = "00:00"
MaskEdBox4.Text = "23:59"
Data1.RecordSource = "SELECT 정산소구분, 근무자명, 시작일시, 종료일시 FROM charge_amt WHERE (시작일시>='" & Format(DTPicker1(0).value, "yyyy-mm-dd 00:00") & "') AND (종료일시<='" & Format(DTPicker1(1).value, "yyyy-mm-dd 23:59") & "') ORDER BY 시작일시"
Data1.Refresh


Set rs = New ADODB.Recordset
QRY = "SELECT MIN(시작일시) AS 시작일시최소값, MAX(종료일시) AS 종료일시최대값 FROM charge_amt WHERE (시작일시>='" & Format(DTPicker1(0).value, "yyyy-mm-dd 00:00") & "') AND (종료일시<='" & Format(DTPicker1(1).value, "yyyy-mm-dd 23:59") & "')"
rs.Open QRY, adoConn

If IsNull(rs!시작일시최소값) Then
    MaskEdBox2.Text = "00:00"
    MaskEdBox4.Text = "23:59"
Else
    MaskEdBox2.Text = Mid(rs!시작일시최소값, 12, 5)
    MaskEdBox4.Text = Mid(rs!종료일시최대값, 12, 5)

End If
End Sub

Private Sub Form_Load()
'Dim rs As Recordset
Dim QRY As String
Dim rs As ADODB.Recordset

Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
MaskEdBox2.Text = "00:00"
MaskEdBox4.Text = "23:59"
DTPicker1(0).value = Now
DTPicker1(1).value = Now


Data1.ConnectionString = AdoConn_Str
Data1.RecordSource = "SELECT 정산소구분, 근무자명, 시작일시, 종료일시 FROM charge_amt WHERE (시작일시>='" & Format(DTPicker1(0).value, "yyyy-mm-dd 00:00") & "') AND (종료일시<='" & Format(DTPicker1(1).value, "yyyy-mm-dd 23:59") & "') ORDER BY 시작일시"
Data1.Refresh

Set rs = New ADODB.Recordset
QRY = "SELECT MIN(시작일시) AS 시작일시최소값, MAX(종료일시) AS 종료일시최대값 FROM charge_amt WHERE (시작일시>='" & Format(DTPicker1(0).value, "yyyy-mm-dd 00:00") & "') AND (종료일시<='" & Format(DTPicker1(1).value, "yyyy-mm-dd 23:59") & "')"

rs.Open QRY, adoConn

If IsNull(rs!시작일시최소값) Then
    MaskEdBox2.Text = "00:00"
    MaskEdBox4.Text = "23:59"
Else
    MaskEdBox2.Text = Mid(rs!시작일시최소값, 12, 5)
    MaskEdBox4.Text = Mid(rs!종료일시최대값, 12, 5)
End If
End Sub


Private Sub Command1_Click(index As Integer)
 Dim SelectionFormula$
 Dim Start_Date As String
 Dim Start_Time As String
 Dim End_Date As String
 Dim End_Time As String
 Dim SQL As String
 Dim rs As Recordset

Report1.Connect = AdoConn_Str
Report1.Formulas(0) = ""
Select Case index
         Case 0
            Report1.ReportFileName = Report_Path_Name$ & "Sail.rpt"
            Start_Date = Format(DTPicker1(0).value, "yyyy-mm-dd")
            Start_Time = MaskEdBox2
            End_Date = Format(DTPicker1(1).value, "yyyy-mm-dd")
            End_Time = MaskEdBox4.Text
            'SelectionFormula$ = "{charge_amt.시작일시} >= '" & Start_Date & " " & Start_Time & "' AND " & "{charge_amt.종료일시} <= '" & End_Date & " " & End_Time & "'"
            On Error GoTo PrintReportError
            DoEvents
            Report1.SelectionFormula = SelectionFormula$
            Report1.CopiesToPrinter = 1
            Report1.Action = 1
            Exit Sub
         Case 1
            Report1.ReportFileName = Report_Path_Name$ & "Time.rpt"
            Record_Source = "charge_time"
            Start_Date = Format(DTPicker1(0).value, "yyyy-mm-dd")
            Start_Time = MaskEdBox2
            End_Date = Format(DTPicker1(1).value, "yyyy-mm-dd")
            End_Time = MaskEdBox4.Text
            SelectionFormula$ = "{charge_time.시작일시} >= '" & Start_Date & " " & Start_Time & "' AND " & "{charge_time.종료일시} <= '" & End_Date & " " & End_Time & "'"
            'On Error GoTo PrintReportError
            DoEvents
            Report1.Formulas(0) = "StartDate=' 자료검색일 : " & Format(DTPicker1(0).value, "yyyy년 mm월 dd일 ") & MaskEdBox2.Text & " ~ " & Format(DTPicker1(1).value, "yyyy년 mm월 dd일 ") & MaskEdBox4.Text & "'"
            Report1.SelectionFormula = SelectionFormula$
            Report1.CopiesToPrinter = 1
            Report1.Action = 1
            Exit Sub
         Case 2
            Report1.ReportFileName = Report_Path_Name$ & "disc.rpt"
            Record_Source = "charge_dic"
            Start_Date = Format(DTPicker1(0).value, "yyyy-mm-dd")
            Start_Time = MaskEdBox2
            End_Date = Format(DTPicker1(1).value, "yyyy-mm-dd")
            End_Time = MaskEdBox4.Text
            SelectionFormula$ = "{charge_dic.시작일시} >= '" & Start_Date & " " & Start_Time & "' AND " & "{charge_dic.종료일시} <= '" & End_Date & " " & End_Time & "'"
            'Debug.Print SelectionFormula$
            On Error GoTo PrintReportError
            DoEvents
            Report1.SelectionFormula = SelectionFormula$
            Report1.CopiesToPrinter = 1
            Report1.Action = 1
            Exit Sub
         Case 3
            Unload Me
            Exit Sub
End Select
PrintReportError:
MsgBox Err.Description
Msg_Box.Label2.Caption = "프린트 작업 오류"
Msg_Box.Label1.Caption = "프린터가 정상작동 하지않습니다."
Msg_Box.Show 1
End Sub

