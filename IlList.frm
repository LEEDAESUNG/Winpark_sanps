VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form IlList 
   ClientHeight    =   11085
   ClientLeft      =   705
   ClientTop       =   1665
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   15210
   StartUpPosition =   1  '소유자 가운데
   Begin Threed.SSCommand SSCommand1 
      Height          =   870
      Left            =   6210
      TabIndex        =   2
      Top             =   10080
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "인쇄(&P)"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   135
      TabIndex        =   1
      Top             =   0
      Width           =   14925
      _Version        =   65536
      _ExtentX        =   26326
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "입차 보고서 (일반권)"
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
         Left            =   11475
         TabIndex        =   4
         Top             =   90
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
            TabIndex        =   5
            Top             =   45
            Width           =   1275
         End
      End
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "IlList.frx":0000
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
      Height          =   870
      Left            =   7740
      TabIndex        =   3
      Top             =   10080
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "종 료(&X)"
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\WinPark\DATA\il.MDB"
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   -1  'True
      Height          =   345
      Left            =   180
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  '스냅샷
      RecordSource    =   "일반권"
      Top             =   9810
      Visible         =   0   'False
      Width           =   2220
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "IlList.frx":0014
      Height          =   9240
      Left            =   135
      OleObjectBlob   =   "IlList.frx":0028
      TabIndex        =   0
      Top             =   540
      Width           =   14925
   End
End
Attribute VB_Name = "IlList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
SSPanel1.FontSize = 16
SSPanel1.FontBold = True
Record_Source = "SELECT * FROM 일반권 ORDER BY 포스번호, 입차일자, 입차시간, 주차권번호"
Data1.RecordSource = Record_Source
Data1.DatabaseName = ParkDb_Path
Data1.Refresh
If (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    LblRecordCount = 0
Else
    'Data1.Recordset.MoveLast
    LblRecordCount = Data1.Recordset.RecordCount
End If
End Sub

Private Sub SSCommand1_Click()
 Dim tmp%
 Dim SelectionFormula$
 On Error GoTo PrintReportError
 DoEvents
 Report1.WindowTitle = SSPanel1.Caption
 Report1.ReportFileName = Report_Path_Name$ & "il.rpt"
 Report1.SortFields(0) = "+{일반권.포스번호}"
 Report1.SortFields(1) = "+{일반권.입차일자}"
 Report1.SortFields(2) = "+{일반권.입차시간}"
 Report1.SortFields(3) = "+{일반권.주차권번호}"
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
