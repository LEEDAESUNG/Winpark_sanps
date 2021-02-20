VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJungSearch 
   BorderStyle     =   1  '단일 고정
   Caption         =   "정기권 발급대장"
   ClientHeight    =   14655
   ClientLeft      =   27855
   ClientTop       =   765
   ClientWidth     =   19125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJungSearch.frx":0000
   ScaleHeight     =   14655
   ScaleWidth      =   19125
   Begin ComctlLib.ListView ListView1 
      Height          =   10005
      Left            =   510
      TabIndex        =   0
      Top             =   1740
      Width           =   18195
      _ExtentX        =   32094
      _ExtentY        =   17648
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   0
   End
   Begin Threed.SSCommand Command1 
      Height          =   570
      Left            =   8370
      TabIndex        =   1
      Top             =   13380
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "검 색"
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
      Picture         =   "frmJungSearch.frx":F651
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1950
      TabIndex        =   2
      Top             =   13470
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
      Format          =   16580608
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   5055
      TabIndex        =   3
      Top             =   13470
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
      Format          =   16580608
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   660
      Left            =   16920
      TabIndex        =   4
      Top             =   13380
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "종 료(&X)"
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
      Picture         =   "frmJungSearch.frx":F9A2
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   660
      Left            =   15210
      TabIndex        =   5
      Top             =   13380
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "엑셀저장"
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
      Picture         =   "frmJungSearch.frx":FCF3
   End
   Begin VB.Label LblRecordCount 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   10440
      TabIndex        =   8
      Top             =   960
      Width           =   1785
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
      Left            =   7290
      TabIndex        =   7
      Top             =   13530
      Width           =   705
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
      Left            =   4200
      TabIndex        =   6
      Top             =   13530
      Width           =   705
   End
End
Attribute VB_Name = "frmJungSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim excel_sql_str As String


Private Sub Form_Load()
Dim Record_Source As String
Dim i As Integer

'On Error GoTo err_P

'Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
'Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
Left = 0
Top = 0

DTPicker1.value = Now
DTPicker2.value = Now

'오늘날짜 데이터만
Glo_JungSearch = "SELECT * FROM regcar WHERE (발급시간 >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (발급시간 <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY 발급시간"
'Debug.Print Glo_JungSearch

Call ListView_Draw

Exit Sub

err_P:
        MsgBox "데이터 베이스 연결실패" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "네트웍 관리자에게 문의 바랍니다." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "데이터 베이스 연결전에는 자료검색 기능을 수행할수 없습니다.", vbCritical

End Sub


Public Sub ListView_Draw()
Dim Column_to_size As Integer
Dim rs As Recordset
Dim QRY As String
Dim itmX As ListItem
Dim INDEX_NO As Long
    
    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No   "
    ListView1.ColumnHeaders.Add , , " 차량번호           "
    ListView1.ColumnHeaders.Add , , " 차 종            "
    ListView1.ColumnHeaders.Add , , " 구 분            "
    ListView1.ColumnHeaders.Add , , " 이 름       "
    ListView1.ColumnHeaders.Add , , " 전화번호             "
    ListView1.ColumnHeaders.Add , , " 월정요금  "
    ListView1.ColumnHeaders.Add , , " 발 급 일             "
    ListView1.ColumnHeaders.Add , , " 시 작 일             "
    ListView1.ColumnHeaders.Add , , " 종 료 일             "
    ListView1.ColumnHeaders.Add , , " 등록일시             "
    
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
 
    Set rs = New ADODB.Recordset
    rs.Open Glo_JungSearch, adoConn
    LblRecordCount = rs.RecordCount & " 건"

    INDEX_NO = 1

    Do While Not (rs.EOF)
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        itmX.SubItems(1) = "" & rs!차량번호
        itmX.SubItems(2) = "" & rs!차종
        itmX.SubItems(3) = "" & rs!구분
        itmX.SubItems(4) = "" & rs!이름
        itmX.SubItems(5) = "" & rs!전화번호
        itmX.SubItems(6) = "" & rs!월정요금
        itmX.SubItems(7) = "" & rs!발급일
        itmX.SubItems(8) = "" & rs!시작일
        itmX.SubItems(9) = "" & rs!종료일
        itmX.SubItems(10) = "" & rs!발급시간
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    INDEX_NO = 0
    Set rs = Nothing

End Sub

Private Sub SSCommand1_Click()
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
    
tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
tmpFileName = App.Path & "\Excel\" & tmpFileName & "_발급대장 검색내역" & ".xls"
'Call makeexcel(ListView1, tmpFileName, "차량입출차현황")
Call makeexcel(ListView1, tmpFileName, "발급대장 검색내역")

Exit Sub

End Sub

'종료
Private Sub SSCommand2_Click()
Unload Me
End Sub

'검색 실행
Private Sub Command1_Click()
Dim i As Integer
Dim Cnt As Integer
Dim Current_Date As String

Dim TmpPath As String
Dim Tmp_File As String
Dim InsSQL As String
Dim Now_Flag As Boolean
Dim sql_str As String
Dim Sort_Order As String

Me.MousePointer = 11

Glo_JungSearch = ""

'쿼리 구성
sql_str = "SELECT * FROM regcar WHERE (발급시간 >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (발급시간 <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY 발급시간"

'Debug.Print sql_str

Glo_JungSearch = sql_str

Call ListView_Draw

Me.MousePointer = 0

'On Error Resume Next

End Sub
