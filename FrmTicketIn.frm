VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTicketIn 
   BorderStyle     =   1  '단일 고정
   Caption         =   " 일반권 입차 현황"
   ClientHeight    =   14715
   ClientLeft      =   300
   ClientTop       =   1800
   ClientWidth     =   19200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmTicketIn.frx":0000
   ScaleHeight     =   14715
   ScaleWidth      =   19200
   Begin VB.TextBox txt_CarNo 
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
      Left            =   12930
      TabIndex        =   28
      Top             =   2130
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   13680
      Width           =   18975
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   5325
      Left            =   360
      TabIndex        =   1
      Top             =   2910
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   9393
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
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   0
      Left            =   17580
      TabIndex        =   2
      Top             =   690
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "닫 기"
      ForeColor       =   14737632
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
      Picture         =   "FrmTicketIn.frx":F4A3
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   1
      Left            =   15450
      TabIndex        =   3
      Top             =   1950
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "삭 제"
      ForeColor       =   255
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
      Picture         =   "FrmTicketIn.frx":F7F4
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   2
      Left            =   16920
      TabIndex        =   4
      Top             =   1950
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "검 색"
      ForeColor       =   14737632
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
      Picture         =   "FrmTicketIn.frx":FB45
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   6030
      TabIndex        =   5
      Top             =   2130
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   8820
      TabIndex        =   6
      Top             =   2130
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
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
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   3
      Left            =   16080
      TabIndex        =   29
      Top             =   690
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "엑 셀"
      ForeColor       =   14737632
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
      Picture         =   "FrmTicketIn.frx":FE96
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
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
      Height          =   285
      Index           =   1
      Left            =   11790
      TabIndex        =   27
      Top             =   2175
      Width           =   1035
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "까지"
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
      Index           =   9
      Left            =   10950
      TabIndex        =   26
      Top             =   2175
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Total : "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Index           =   0
      Left            =   19890
      TabIndex        =   25
      Top             =   8580
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
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   8160
      TabIndex        =   24
      Top             =   2175
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회기간 :"
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
      Index           =   5
      Left            =   4920
      TabIndex        =   23
      Top             =   2175
      Width           =   1035
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '투명
      Caption         =   "일반권 입차 현황 검색"
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
      Left            =   510
      TabIndex        =   22
      Top             =   2010
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "조회 건수 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   15840
      TabIndex        =   21
      Top             =   8550
      Width           =   1215
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
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   17100
      TabIndex        =   20
      Top             =   8550
      Width           =   1425
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '투명
      Caption         =   "# 차량 입차 내역"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   0
      Left            =   12300
      TabIndex        =   19
      Top             =   9480
      Width           =   4035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "일반권 입차 현황"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   1470
      TabIndex        =   18
      Top             =   870
      Width           =   3525
   End
   Begin VB.Label lbl_Search 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "원"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   21000
      TabIndex        =   17
      Top             =   8550
      Width           =   2025
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
      Height          =   4290
      Index           =   0
      Left            =   240
      Picture         =   "FrmTicketIn.frx":101E7
      Stretch         =   -1  'True
      Top             =   9300
      Width           =   5730
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "입차일시 :"
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
      Height          =   435
      Index           =   0
      Left            =   12510
      TabIndex        =   16
      Top             =   10500
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "인식번호 :"
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
      Height          =   435
      Index           =   1
      Left            =   12510
      TabIndex        =   15
      Top             =   11025
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "인식상태 :"
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
      Height          =   435
      Index           =   2
      Left            =   12510
      TabIndex        =   14
      Top             =   12075
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "차량번호 :"
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
      Height          =   435
      Index           =   3
      Left            =   12510
      TabIndex        =   13
      Top             =   11550
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      Caption         =   "처리일시 :"
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
      Height          =   435
      Index           =   4
      Left            =   12510
      TabIndex        =   12
      Top             =   12600
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
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
      Height          =   435
      Index           =   5
      Left            =   14070
      TabIndex        =   11
      Top             =   10500
      Width           =   4665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
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
      Height          =   435
      Index           =   6
      Left            =   14070
      TabIndex        =   10
      Top             =   11025
      Width           =   4665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
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
      Height          =   435
      Index           =   7
      Left            =   14070
      TabIndex        =   9
      Top             =   12075
      Width           =   4665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
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
      Height          =   435
      Index           =   8
      Left            =   14070
      TabIndex        =   8
      Top             =   11550
      Width           =   4665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   435
      Index           =   9
      Left            =   14070
      TabIndex        =   7
      Top             =   12600
      Width           =   4665
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   240
      Top             =   1770
      Width           =   18735
   End
End
Attribute VB_Name = "FrmTicketIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer

Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
'Left = 0
'Top = 0
'Me.cmb_Gubun = Me.cmb_Gubun.List(0)
DTPicker1.value = Now
DTPicker2.value = Now
Glo_SQL_REG = "SELECT * FROM ilbancarin WHERE (처리일시 >= '" & Format(DTPicker1, "yyyymmdd") & "000000') AND (처리일시 <= '" & Format(DTPicker2, "yyyymmdd") & "235959') ORDER BY 처리일시"
'Glo_SQL_REG = "SELECT * From TB_FEE WHERE ORDER BY REG_DATE ASC"
Call ListView_REG_Draw
Call ListView_REG_SQL
List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    일반권 입차현황 시작...!!", 0
Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    일반권 입차현황 시작...!!")


End Sub

Public Sub ListView_REG_SQL()
Dim rs As Recordset
Dim QRY As String
Dim itmX As ListItem
Dim INDEX_NO As Long
Dim TOTAL_FEE As Single

INDEX_NO = 1
TOTAL_FEE = 0
Set rs = New ADODB.Recordset
rs.Open Glo_SQL_REG, adoConn
lbl_COUNT = rs.RecordCount
Do While Not (rs.EOF)
    Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
    itmX.SubItems(1) = "" & rs!입차일자
    itmX.SubItems(2) = "" & rs!입차시간
    itmX.SubItems(3) = "" & rs!인식번호
    itmX.SubItems(4) = "" & rs!차량번호
    itmX.SubItems(5) = "" & rs!인식상태
    itmX.SubItems(6) = "" & rs!처리일시
    itmX.SubItems(7) = "" & rs!구분
    itmX.SubItems(8) = "" & rs!게이트구분
    itmX.SubItems(9) = "" & rs!이미지명
'    TOTAL_FEE = TOTAL_FEE + Val(rs!실수령금액)
'    lbl_Search.Caption = TOTAL_FEE & "원"
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
    .ListView_REG.ColumnHeaders.Add , , " 입차일자         "
    .ListView_REG.ColumnHeaders.Add , , " 입차시간     "
    .ListView_REG.ColumnHeaders.Add , , " 인식번호         "
    .ListView_REG.ColumnHeaders.Add , , " 차량번호         "
    .ListView_REG.ColumnHeaders.Add , , " 인식상태     "
    .ListView_REG.ColumnHeaders.Add , , " 처리일시                           "
    .ListView_REG.ColumnHeaders.Add , , " 할인적용         "
    .ListView_REG.ColumnHeaders.Add , , " 등록부서         "
    .ListView_REG.ColumnHeaders.Add , , " 이미지명                           "
    .ListView_REG.ColumnHeaders.Add , , " "
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim Tmp_Path1, Tmp_Path2 As String

ListView_REG.SetFocus

lbl_Name(5) = ListView_REG.SelectedItem.SubItems(1) & " " & ListView_REG.SelectedItem.SubItems(2)
lbl_Name(6) = ListView_REG.SelectedItem.SubItems(3)
lbl_Name(7) = ListView_REG.SelectedItem.SubItems(5)
lbl_Name(8) = ListView_REG.SelectedItem.SubItems(4)
lbl_Name(9) = ListView_REG.SelectedItem.SubItems(6)

Tmp_Path1 = Dir(ListView_REG.SelectedItem.SubItems(9))
If (Tmp_Path1 = "") Then
    ImageIn(0).Picture = Nothing
Else
    ImageIn(0).Picture = LoadPicture(ListView_REG.SelectedItem.SubItems(9))
End If
'Tmp_Path2 = Dir(ListView_REG.SelectedItem.SubItems(9))
'If (Tmp_Path2 = "") Then
'    ImageIn(1).Picture = Nothing
'Else
'    ImageIn(1).Picture = LoadPicture(ListView_REG.SelectedItem.SubItems(9))
'End If

End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim sql_str As String

On Error Resume Next

Select Case Index
    Case 0  '종료
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    일반권 입차현황 종료", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    일반권 입차현황 종료")
        Unload Me
        Exit Sub

    Case 1  '삭제
        If (Len(lbl_Name(9)) = 0) Then
            Msg_Box.Label2.Caption = "데이터 선택 오류"
            Msg_Box.Label1.Caption = "삭제할 데이터를 선택해 주십시요."
            Msg_Box.Show 1
            Exit Sub
        End If
        
        MBox.Label3.Caption = lbl_Name(8)
        MBox.Label1.Caption = "위 일반차량의 입차자료를 삭제합니다. 삭제하시겠습니까?"
        MBox.Label2.Caption = "일반권 자료 삭제"
        MBox.Show 1
        
        If (Glo_MsgRet = True) Then
            '삭제쿼리
            adoConn.Execute "Delete from ilbancarin where 처리일시 = '" & lbl_Name(9) & "'"
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & lbl_Name(5) & "  " & lbl_Name(8) & " 일반차량의 입차내역을 삭제", 0
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & lbl_Name(5) & "  " & lbl_Name(8) & " 일반차량의 입차내역을 삭제")
        Else
        End If
        Call cmd_Button_Click(2)
        Exit Sub
        
    Case 2
        '차량결제내역 검색
        Me.MousePointer = 11
        Glo_SQL_SEARCH = ""
        '쿼리 구성
        '조회기간
        'Glo_SQL_REG = "SELECT * FROM TB_FEE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY REG_DATE ASC"
        sql_str = "SELECT * FROM ilbancarin WHERE (처리일시 >= '" & Format(DTPicker1, "yyyymmdd") & "000000" & "') AND (처리일시 <= '" & Format(DTPicker2, "yyyymmdd") & "235959" & "')"
        '차량번호 검색
        If (txt_CarNo.Text <> "") Then
            If IsNumeric(txt_CarNo) And Len(txt_CarNo) = 4 Then
            Else
                MsgBox "차량번호 끝4자리를 확인해주세요."
                Me.MousePointer = 0
                Exit Sub
            End If
            sql_str = sql_str & " AND (차량번호 Like '%" & txt_CarNo.Text & "')"
        End If
'        '운전자 이름 검색
'        If (txt_Name.Text <> "") Then
'            sql_str = sql_str & " AND (DRIVER_NAME Like '%" & txt_Name.Text & "%')"
'        End If
        sql_str = sql_str & " ORDER BY 처리일시"
        Glo_SQL_REG = sql_str
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sql_str, 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sql_str)
        
        'Debug.Print Glo_SQL_REG
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Me.MousePointer = 0
        Exit Sub
    Case 3
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_일반권_입차내역" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "일반권_입차내역")
        Exit Sub
End Select


End Sub

'엔터키 입력시 탭 실행
'폼속성 keypreview = true 설정
Private Sub Form_KeyPRESS(KeyAscii As Integer)
    
If KeyAscii = vbKeyReturn Then
    Call cmd_Button_Click(2)
    KeyAscii = 0
    'SendKeys "{TAB}"
End If

End Sub

