VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form T_OutList 
   ClientHeight    =   14925
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   ScaleHeight     =   14925
   ScaleWidth      =   19035
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Caption         =   "자료검색"
      Height          =   855
      Left            =   135
      TabIndex        =   5
      Top             =   13950
      Width           =   11850
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7410
         MaxLength       =   4
         TabIndex        =   11
         Top             =   330
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command1 
         Caption         =   "확인"
         Height          =   405
         Left            =   10320
         TabIndex        =   9
         Top             =   300
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1635
         TabIndex        =   6
         Top             =   315
         Width           =   1920
         _ExtentX        =   3387
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
         Format          =   52756480
         CurrentDate     =   36927
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   4050
         TabIndex        =   8
         Top             =   315
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   52756480
         CurrentDate     =   36927
      End
      Begin VB.Label Label1 
         Caption         =   "검색날짜 선택 :                                      ~                                       주차권번호 : "
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
         Left            =   135
         TabIndex        =   7
         Top             =   390
         Visible         =   0   'False
         Width           =   7440
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   780
      Left            =   12225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   14025
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "인쇄(&P)"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   135
      TabIndex        =   0
      Top             =   0
      Width           =   18960
      _Version        =   65536
      _ExtentX        =   33443
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "출차 보고서 (교통카드 및 T머니카드)"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Begin Threed.SSPanel PnlOut 
         Height          =   390
         Index           =   7
         Left            =   15600
         TabIndex        =   3
         Top             =   75
         Width           =   3210
         _Version        =   65536
         _ExtentX        =   5662
         _ExtentY        =   688
         _StockProps     =   15
         Caption         =   "  레코드 건수"
         ForeColor       =   0
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
         Begin VB.Label LblRecordCount 
            BackColor       =   &H00008000&
            BeginProperty Font 
               Name            =   "궁서"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   1665
            TabIndex        =   4
            Top             =   45
            Width           =   1275
         End
      End
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "T_OutList.frx":0000
      Left            =   0
      Top             =   0
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
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   780
      Left            =   13785
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   14025
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "종 료(X)"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Winpark\Data\ParkDb.MDB"
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   0   'False
      Height          =   345
      Left            =   15330
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  '스냅샷
      RecordSource    =   "T_머니출차"
      Top             =   14430
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "T_OutList.frx":0014
      Height          =   13275
      Left            =   135
      OleObjectBlob   =   "T_OutList.frx":0028
      TabIndex        =   10
      Top             =   540
      Width           =   18960
   End
End
Attribute VB_Name = "T_OutList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Data1.RecordSource = "SELECT * FROM t-out WHERE (출차일자>='" & Format(DTPicker1.Value, "yyyy-mm-dd") & "') AND (출차일자<='" & Format(DTPicker2.Value, "yyyy-mm-dd") & "') ORDER BY 포스번호, 출차일자, 출차시간, 카드번호"
Data1.Refresh
If (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    LblRecordCount = 0
Else
    Data1.Recordset.MoveLast
    LblRecordCount = Data1.Recordset.RecordCount
End If
End Sub

Private Sub Form_Load()
Dim InsSQL As String
Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.

SSPanel1.FontSize = 16
SSPanel1.FontBold = True
Record_Source = "SELECT * FROM t-out WHERE 출차일자 = '" & Format(Now, "yyyy-mm-dd") & "' ORDER BY 포스번호, 출차일자, 출차시간, 카드번호"
Data1.RecordSource = Record_Source
Data1.DatabaseName = ParkDb_Path
Data1.Refresh
If (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    LblRecordCount = 0
Else
    'Data1.Recordset.MoveLast
    LblRecordCount = Data1.Recordset.RecordCount
End If
DTPicker1.Value = Now
DTPicker2.Value = Now

End Sub

Private Sub SSCommand1_Click()
 Dim tmp%
 Dim SelectionFormula$
 On Error GoTo PrintReportError
 DoEvents
 Report1.WindowTitle = SSPanel1.Caption
 Report1.ReportFileName = Report_Path_Name$ & "tmpil.rpt"

 Report1.Formulas(0) = "StartDate=' 자료검색일 : " & Format(DTPicker1.Value, "yyyy년 mm월 dd일") & " ~ " & Format(DTPicker2.Value, "yyyy년 mm월 dd일") & "'"
 Report1.SortFields(0) = "+{임시일반권.포스번호}"
 Report1.SortFields(1) = "+{임시일반권.입차일자}"
 Report1.SortFields(2) = "+{임시일반권.입차시간}"
 Report1.SortFields(3) = "+{임시일반권.주차권번호}"
 Report1.SelectionFormula = SelectionFormula$
 Report1.CopiesToPrinter = 1
 Report1.Action = 1
 Exit Sub
PrintReportError:
Msg_Box.Caption = "프린트 작업 오류"
Msg_Box.Label1.Caption = "프린터가 정상작동 하지않습니다."
Msg_Box.Show 1
End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub

