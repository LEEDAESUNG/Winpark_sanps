VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form JungList2 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   ClientHeight    =   14895
   ClientLeft      =   25500
   ClientTop       =   360
   ClientWidth     =   19155
   LinkTopic       =   "Form1"
   Picture         =   "JungList2.frx":0000
   ScaleHeight     =   14895
   ScaleWidth      =   19155
   Begin Threed.SSCommand Command2 
      Cancel          =   -1  'True
      Height          =   780
      Left            =   17220
      TabIndex        =   3
      Top             =   13335
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "종료"
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
      RoundedCorners  =   0   'False
      Picture         =   "JungList2.frx":FDEF
   End
   Begin Threed.SSCommand Command1 
      Height          =   780
      Left            =   15840
      TabIndex        =   2
      Top             =   13335
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "인쇄"
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
      Picture         =   "JungList2.frx":10140
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "JungList2.frx":10491
      Left            =   10350
      List            =   "JungList2.frx":104A1
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   13470
      Width           =   2115
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "JungList2.frx":104C5
      Left            =   19590
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Winpark\Data\jlist.rpt"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   1024
      WindowHeight    =   768
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   19560
      Top             =   300
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   6000
      TabIndex        =   4
      Top             =   13500
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
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "JungList2.frx":104D9
      Height          =   10260
      Left            =   360
      TabIndex        =   6
      Top             =   1620
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   18098
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "차량번호"
         Caption         =   " 차량번호"
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
         DataField       =   "차종"
         Caption         =   " 차량모델"
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
         DataField       =   "이름"
         Caption         =   " 이름"
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
      BeginProperty Column04 
         DataField       =   "전화번호"
         Caption         =   " 전화번호"
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
         DataField       =   "발급일"
         Caption         =   " 발급일"
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
         DataField       =   "발급시간"
         Caption         =   " Upadate"
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
         DataField       =   "시작일"
         Caption         =   " 시작일"
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
         DataField       =   "종료일"
         Caption         =   " 종료일"
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
         DataField       =   "비고"
         Caption         =   " 비 고"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   3825.071
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "날짜선택 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   4365
      TabIndex        =   5
      Top             =   13485
      Width           =   1905
   End
   Begin VB.Label LblRecordCount 
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11640
      TabIndex        =   0
      Top             =   1005
      Width           =   1275
   End
End
Attribute VB_Name = "JungList2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Order$

Private Sub Combo1_Click()
Order$ = "  ORDER BY " & Combo1.List(Combo1.ListIndex)
Adodc1.RecordSource = "SELECT * FROM regcar WHERE 발급일 = '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'" & Order$
'Adodc1.RecordSource = Record_Source
Adodc1.Refresh
If (Adodc1.Recordset.RecordCount <> 0) Then
    Adodc1.Recordset.MoveLast
End If
LblRecordCount.Caption = Adodc1.Recordset.RecordCount
End Sub

Private Sub DTPicker1_Change()
Adodc1.RecordSource = "SELECT * FROM regcar WHERE 발급일 = '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'" & Order$
Adodc1.Refresh
If (Adodc1.Recordset.RecordCount <> 0) Then
    Adodc1.Recordset.MoveLast
End If
LblRecordCount.Caption = Adodc1.Recordset.RecordCount

End Sub

Private Sub Form_Activate()
If (Adodc1.Recordset.RecordCount <> 0) Then
    Adodc1.Recordset.MoveLast
End If
LblRecordCount.Caption = Adodc1.Recordset.RecordCount
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
Adodc1.ConnectionString = AdoConn_Str
Report1.Connect = AdoConn_Str
'Adodc1.RecordSource = "SELECT * FROM regcar WHERE 종료일 >= '" & Format(Now, "yyyy-mm-dd") & "'"
'Adodc1.RecordSource = "SELECT * FROM regcar"
DTPicker1.value = Now
Adodc1.RecordSource = "SELECT * FROM regcar WHERE 발급일 = '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'" & Order$
Adodc1.Refresh
Combo1.ListIndex = 0
End Sub

Private Sub Command1_Click()
 Dim tmp%
 Dim SelectionFormula$
 On Error GoTo PrintReportError
 SelectionFormula$ = "{regcar.발급일} = '" & Format(DTPicker1.value, "yyyy-mm-dd") & "'"
 DoEvents
 Report1.ReportFileName = Report_Path_Name$ & "jlist3.rpt"
 
 Select Case Combo1.ListIndex
        Case 0
             Report1.SortFields(0) = "+{regcar.차량번호}"
        Case 1
             Report1.SortFields(0) = "+{regcar.이름}"
        Case 2
             '성훈
             Report1.SortFields(0) = "+{regcar.구분}"
             'Report1.SortFields(0) = "+{regcar.소속}"
        Case 2
             Report1.SortFields(0) = "+{regcar.발급시간}"
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

Private Sub Command2_Click()
Unload Me
End Sub

