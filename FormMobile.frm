VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormMobile 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ComboBox Combo3 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FormMobile.frx":0000
      Left            =   7005
      List            =   "FormMobile.frx":0002
      Style           =   2  '드롭다운 목록
      TabIndex        =   17
      Top             =   2880
      Width           =   1830
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FormMobile.frx":0004
      Left            =   7005
      List            =   "FormMobile.frx":0006
      Style           =   2  '드롭다운 목록
      TabIndex        =   15
      Top             =   2445
      Width           =   1830
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "기종"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FormMobile.frx":0008
      Left            =   7005
      List            =   "FormMobile.frx":000A
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   2010
      Width           =   1830
   End
   Begin ComctlLib.ListView ListView_Mobile 
      Height          =   4770
      Left            =   120
      TabIndex        =   0
      Top             =   4275
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   8414
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   10920
      TabIndex        =   2
      Top             =   120
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "닫기"
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
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormMobile.frx":000C
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   7005
      TabIndex        =   4
      Top             =   1185
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9
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
      Format          =   138870784
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   9765
      TabIndex        =   5
      Top             =   1185
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9
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
      Format          =   138870784
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   7005
      TabIndex        =   6
      Top             =   1560
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9
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
      Format          =   138870787
      UpDown          =   -1  'True
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   9765
      TabIndex        =   7
      Top             =   1560
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   9
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
      Format          =   138870787
      UpDown          =   -1  'True
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   7005
      TabIndex        =   19
      Top             =   3570
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "검 색"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   15
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormMobile.frx":035D
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   9690
      TabIndex        =   20
      Top             =   3795
      Width           =   2520
      _Version        =   65536
      _ExtentX        =   4445
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   "  검색 건수"
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
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1170
         TabIndex        =   21
         Top             =   60
         Width           =   1275
      End
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "알림확인"
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
      Left            =   5775
      TabIndex        =   18
      Top             =   2895
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "발생위치"
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
      Left            =   5775
      TabIndex        =   16
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장치구분"
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
      Left            =   5775
      TabIndex        =   14
      Top             =   2025
      Width           =   900
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
      Left            =   5775
      TabIndex        =   13
      Top             =   1605
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
      Left            =   5775
      TabIndex        =   12
      Top             =   1215
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
      Left            =   8985
      TabIndex        =   11
      Top             =   1605
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
      Left            =   8985
      TabIndex        =   10
      Top             =   1215
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
      Left            =   11745
      TabIndex        =   9
      Top             =   1605
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
      Left            =   11745
      TabIndex        =   8
      Top             =   1215
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   12225
      Y1              =   780
      Y2              =   795
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '투명
      Caption         =   " 모바일 알림"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   4185
   End
End
Attribute VB_Name = "FormMobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Selected_CCTV_URL As String

Private Sub Form_Load()

    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    DTPicker3.Format = dtpCustom
    DTPicker3.CustomFormat = "HH:mm:ss"
    DTPicker3.Refresh
    
    DTPicker4.Format = dtpCustom
    DTPicker4.CustomFormat = "HH:mm:ss"
    DTPicker4.Refresh

    DTPicker1.value = Now
    DTPicker2.value = Now
    DTPicker3.value = Format("00:00:00")
    DTPicker4.value = Format("23:59:59")
    
    
    Combo1.AddItem "전체": Combo1.ListIndex = 0
    Combo2.AddItem "전체": Combo2.ListIndex = 0
    Combo3.AddItem "전체": Combo3.ListIndex = 0
    
    Call Clear_Field
    Call ListView_Mobile_Draw
    'Call ListView_Mobile_SQL
    
End Sub

Private Function GetSQL()
    Dim sql_str As String
    
    sql_str = "SELECT * FROM tb_event WHERE (DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & ".000') AND (DATE <='" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & ".999')"
    
    If (Combo1.text <> "전체") Then '장치구분
        sql_str = sql_str & " AND (TYPE = '" & Combo1.text & "')"
    End If
    
    If (Combo2.text <> "전체") Then '발생위치
        sql_str = sql_str & " AND (TYPE = '" & Combo2.text & "')"
    End If
    
    If (Combo3.text <> "전체") Then '알림확인
        sql_str = sql_str & " AND (TYPE = '" & Combo3.text & "')"
    End If

    GetSQL = sql_str
End Function


Private Sub SSCommand1_Click()
    Call Clear_Field
    Call ListView_Mobile_Draw
    Call ListView_Mobile_SQL(GetSQL)
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub

Private Sub Clear_Field()
    
End Sub

Private Sub ListView_Mobile_SQL(qry As String)
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Integer
    Dim sQry As String
    
    sQry = qry
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
    If (bQryResult = False) Then
        Exit Sub
    End If

    INDEX_NO = 0
    Do While Not (rs.EOF)
    
        INDEX_NO = INDEX_NO + 1
        
        Set itmX = ListView_Mobile.ListItems.Add(, , "" & rs!SEQ)
        
        i = 1
        itmX.SubItems(i) = "" & rs!Type: i = i + 1
        itmX.SubItems(i) = "" & rs!Location: i = i + 1
        itmX.SubItems(i) = "" & rs!Date: i = i + 1
        itmX.SubItems(i) = "" & rs!Count: i = i + 1
        itmX.SubItems(i) = "" & rs!CONFIRMDATE: i = i + 1
        itmX.SubItems(i) = "" & rs!DEVICE: i = i + 1

        rs.MoveNext
    Loop
    Set rs = Nothing
    
    LblRecordCount = INDEX_NO
    
End Sub

Private Sub ListView_Mobile_Draw()
    Dim Column_to_size As Integer
    
    With Me
        Call ListViewExtended(.ListView_Mobile)
        .ListView_Mobile.View = lvwReport
        .ListView_Mobile.ListItems.Clear
        .ListView_Mobile.ColumnHeaders.Clear
        .ListView_Mobile.ColumnHeaders.Add , , " No   "
        .ListView_Mobile.ColumnHeaders.Add , , " 장치구분              " '차단기충격, ...
        .ListView_Mobile.ColumnHeaders.Add , , " 발생위치              " 'Lane1, Lane2, ..., 출구무인기, 사전무인기...
        .ListView_Mobile.ColumnHeaders.Add , , " 발생일시              " 'yyyy-mm-dd hh:nn:ss
        .ListView_Mobile.ColumnHeaders.Add , , " 알림횟수              " '1~3
        .ListView_Mobile.ColumnHeaders.Add , , " 확인일시              " 'yyyy-mm-dd hh:nn:ss
        .ListView_Mobile.ColumnHeaders.Add , , " 알림대상              " '모바일 사용자이름
        
        For Column_to_size = 0 To .ListView_Mobile.ColumnHeaders.Count - 2
             SendMessage .ListView_Mobile.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With

End Sub



Private Sub ListView_Mobile_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim i As Integer
    With ListView_Mobile
        For i = 1 To .ColumnHeaders.Count
            If (.ColumnHeaders.Item(i) = ColumnHeader) Then
                .SortKey = i - 1
                .SortOrder = .SortOrder Xor 1
                '.SortOrder = lvwDescending
                .Sorted = True
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub ListView_Mobile_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Error Resume Next
    
    ListView_Mobile.SetFocus
    
End Sub




