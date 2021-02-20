VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form IlINList 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   ClientHeight    =   14700
   ClientLeft      =   26100
   ClientTop       =   600
   ClientWidth     =   19185
   LinkTopic       =   "Form1"
   Picture         =   "IlINList.frx":0000
   ScaleHeight     =   14700
   ScaleWidth      =   19185
   Begin Threed.SSCommand Command1 
      Height          =   570
      Left            =   12000
      TabIndex        =   8
      Top             =   13410
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
      Picture         =   "IlINList.frx":FF1E
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
      Height          =   510
      Left            =   10485
      MaxLength       =   4
      TabIndex        =   5
      Top             =   13440
      Width           =   1275
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   780
      Left            =   15870
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   13320
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
      Picture         =   "IlINList.frx":1026F
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   10950
      Top             =   10860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Winpark\Data\IL.RPT"
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
      Left            =   17280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   13305
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
      Picture         =   "IlINList.frx":105C0
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   8100
      Top             =   10890
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
      Bindings        =   "IlINList.frx":10911
      Height          =   10320
      Left            =   345
      TabIndex        =   2
      Top             =   1590
      Width           =   18480
      _ExtentX        =   32597
      _ExtentY        =   18203
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   25
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "포스번호"
         Caption         =   " 포스번호"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "키코드"
         Caption         =   " 키코드"
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
         DataField       =   "정산소구분"
         Caption         =   " 정산소구분"
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
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   2849.953
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   2729.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2954.835
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   3225.26
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   2595
      TabIndex        =   3
      Top             =   13530
      Width           =   2070
      _ExtentX        =   3651
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
      TabIndex        =   4
      Top             =   13530
      Width           =   2070
      _ExtentX        =   3651
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
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   11460
      TabIndex        =   6
      Top             =   885
      Width           =   3720
      _Version        =   65536
      _ExtentX        =   6562
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
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   1740
         TabIndex        =   7
         Top             =   30
         Width           =   1275
      End
   End
End
Attribute VB_Name = "IlINList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectionFormula$

Private Sub Command1_Click()

Me.MousePointer = 11

If (Text1.Text = "") Then
    Data1.RecordSource = "SELECT * FROM ilbanin WHERE (입차일자>='" & Format(DTPicker1.value, "yyyy-mm-dd") & "') AND (입차일자<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "')"
    SelectionFormula$ = "{ilbanin.입차일자} >= '" & Format(DTPicker1.value, "yyyy-mm-dd") & "' AND {ilbanin.입차일자}<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "'"
Else
    Data1.RecordSource = "SELECT * FROM ilbanin WHERE (입차일자>='" & Format(DTPicker1.value, "yyyy-mm-dd") & "') AND (입차일자<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "') AND (주차권번호 = '" & Text1.Text & "')"
    SelectionFormula$ = "{ilbanin.입차일자} >= '" & Format(DTPicker1.value, "yyyy-mm-dd") & "' AND {ilbanin.입차일자}<='" & Format(DTPicker2.value, "yyyy-mm-dd") & "'" & " AND {ilbanin.주차권번호}='" & Text1.Text & "'"
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

'Data1.ConnectionString = AdoConn_Str
'Report1.ReportFileName = Report_Path_Name$ & "ilban_in.rpt"

'Report1.Connect = AdoConn_Str
''SSPanel1.FontSize = 16
''SSPanel1.FontBold = True
'Report1.Connect = AdoConn_Str

'성훈
Data1.RecordSource = "SELECT * FROM ilbanin WHERE 입차일자='" & Format(Now, "yyyy-mm-dd") & " ORDER BY 처리일시'"
'Data1.RecordSource = "SELECT * FROM ilbanin WHERE 입차일자='" & Format(Now, "yyyy-mm-dd") & "'"
Data1.Refresh

If (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    LblRecordCount = 0
Else
    Data1.Recordset.MoveLast
    LblRecordCount = Data1.Recordset.RecordCount
End If
DTPicker1.value = Now
DTPicker2.value = Now

End Sub

Private Sub SSCommand1_Click()
' Dim tmp%
'
' On Error GoTo PrintReportError
' DoEvents
' Report1.WindowTitle = "입차보고서(입차)"
' 'Report1.ReportFileName = Report_Path_Name$ & "il.rpt"
' Report1.ReportFileName = Report_Path_Name$ & "ilban_IN.rpt"
' '{ilban.입차일자} >= '2008-03-31' AND {ilban.입차일자}<='2008-04-30'
'
' Report1.Formulas(0) = "StartDate=' 자료검색일 : " & Format(DTPicker1.value, "yyyy년 mm월 dd일") & " ~ " & Format(DTPicker2.value, "yyyy년 mm월 dd일") & "'"
' Report1.SortFields(0) = "+{ilban.키코드}"
' Report1.SortFields(1) = "+{ilban.입차시간}"
' Report1.SortFields(2) = "+{ilban.주차권번호}"
' Report1.SelectionFormula = SelectionFormula$
' Report1.CopiesToPrinter = 1
' Report1.Action = 1
' Exit Sub
'PrintReportError:
'Msg_Box.Label2.Caption = "프린트 작업 오류"
'Msg_Box.Label1.Caption = "프린터가 정상작동 하지않습니다."
'Msg_Box.Show 1

 ' On Error GoTo PrintReportError
  
  Report1.Formulas(0) = "StartDate=' 자료검색일 : " & Format(DTPicker1.value, "yyyy년 mm월 dd일") & " ~ " & Format(DTPicker2.value, "yyyy년 mm월 dd일") & "'"
'  Select Case Combo3.ListIndex
'            Case 0
'                     Report1.SortFields(0) = "+{regcarinout.차량번호}"
'                     Report1.SortFields(1) = "+{regcarinout.처리일시}"
'            Case 1
'                     Report1.SortFields(0) = "+{regcarinout.이름}"
'                     Report1.SortFields(1) = "+{regcarinout.처리일시}"
'            Case 2
'                     Report1.SortFields(0) = "+{regcarinout.소속}"
'                     Report1.SortFields(1) = "+{regcarinout.처리일시}"
'            Case 3
'                     Report1.SortFields(0) = "+{regcarinout.입출상태}"
'                     Report1.SortFields(1) = "+{regcarinout.처리일시}"
'            Case 4
'                     Report1.SortFields(0) = "+{regcarinout.인식상태}"
'                     Report1.SortFields(1) = "+{regcarinout.처리일시}"
'  End Select
  DoEvents
  
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

