VERSION 5.00
Begin VB.Form FrmCSV 
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   8415
   ClientLeft      =   11940
   ClientTop       =   3030
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmCSV.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   6765
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2310
      Left            =   60
      TabIndex        =   13
      Top             =   5985
      Width           =   6690
   End
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5655
      TabIndex        =   12
      Top             =   795
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 정기권 일괄등록 "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4425
      Left            =   75
      TabIndex        =   0
      Top             =   1515
      Width           =   6660
      Begin VB.CommandButton cmb_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등 록"
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5505
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   7
         Top             =   3780
         Width           =   945
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   150
         TabIndex        =   6
         Top             =   1410
         Width           =   2925
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   5
         Top             =   915
         Width           =   2925
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   3165
         TabIndex        =   4
         Top             =   900
         Width           =   3375
      End
      Begin VB.TextBox TxtDuration 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3960
         Width           =   1665
      End
      Begin VB.TextBox TxtStop 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3630
         Width           =   1665
      End
      Begin VB.TextBox TxtOpen 
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3300
         Width           =   1665
      End
      Begin VB.Label lbl_TextPath 
         BackStyle       =   0  '투명
         BeginProperty Font 
            Name            =   "나눔고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   135
         TabIndex        =   11
         Top             =   330
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Proc. Time : "
         Height          =   300
         Index           =   0
         Left            =   225
         TabIndex        =   10
         Top             =   3975
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "End Time  : "
         Height          =   300
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   3630
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "Start Time : "
         Height          =   300
         Index           =   2
         Left            =   225
         TabIndex        =   8
         Top             =   3300
         Width           =   1185
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   " # 주의 # 시스템 관리자 외에 절대 사용 금지...!!"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   135
      Width           =   5520
   End
End
Attribute VB_Name = "FrmCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmb_Reg_Click()
    'Open 함수를 이용 Text File을 Open 한 뒤 읽어들인다.
    Dim LF          As Long             'File Pointer
    Dim strLine     As String           '파일을 한 라인씩 읽어들이는 변수
    Dim sngStart    As Single
    Dim sngStop     As Single
    Dim sngDuration As Single
    Dim rs As ADODB.Recordset
    Dim Qry As String
    Dim i, s, q As String
    
    Dim CarNo As String
    Dim CarModel As String
    Dim CarGubun As String
    Dim DriverName As String
    Dim DriverPhone As String
    Dim Dong As String
    Dim Ho As String
    Dim DtEnd As String
    
On Error GoTo erro_p
    
    If Len(lbl_TextPath) = 0 Then
        MsgBox "등록 파일을 선택해 주세요..!! "
        Exit Sub
    End If
    
    i = 0
    'TxtLineStatus = "파일 읽기 시작"
    sngStart = Timer
    TxtOpen = CStr(sngStart): TxtOpen.Refresh

    LF = FreeFile()
    'Open App.Path & RFileName For Input As LF     '파일 Open 읽기
    Open lbl_TextPath For Input As LF     '파일 Open 읽기
    Do While Not EOF(LF)
        Line Input #LF, strLine           '내용을 읽어들인다.
        If i <> 0 Then
            s = InStr(1, strLine, ",", 1)
            CarNo = Trim(Left(strLine, (s - 1)))
            q = InStr(s + 1, strLine, ",", 1)
            CarModel = Mid(strLine, (s + 1), (q - s - 1))
            
            s = InStr(q + 1, strLine, ",")
            CarGubun = Mid(strLine, (q + 1), (s - q - 1))
            q = InStr(s + 1, strLine, ",")
            DriverName = Trim(Mid(strLine, (s + 1), (q - s - 1)))
            
            s = InStr(q + 1, strLine, ",")
            DriverPhone = Mid(strLine, (q + 1), (s - q - 1))
            q = InStr(s + 1, strLine, ",")
            Dong = Mid(strLine, (s + 1), (q - s - 1))
            
            s = InStr(q + 1, strLine, ",")
            Ho = Mid(strLine, (q + 1), (s - q - 1))
            q = InStr(s + 1, strLine, ",")
            DtEnd = Mid(strLine, s + 1, 8)
        
            Qry = "SELECT * From tb_reg Where CAR_NO = '" & Trim(CarNo) & "'"
            Set rs = New ADODB.Recordset
            rs.Open Qry, adoConn
            If (rs.EOF) Then
                'adoConn.Execute "Insert Into tb_member VALUES ('', '', '정기권', '0', '', '', '', '', '" & Format(Now, "YYYYMMDD") & "', '" & Format(Now, "YYYYMMDD") & "', '', '" & Format(Now, "YYYYMMDD") & "', '', '', '0', '0', '" & Mid(strLine, 18, 10) & "', '" & Mid(strLine, 1, 8) & "', '0', '', '', '1')"
                'adoConn.Execute "Insert Into tb_reg VALUES ('" & Trim(CarNo) & "', '" & Trim(CarModel) & "','정기권','0', '" & Trim(DriverName) & "','" & Trim(DriverPhone) & "','" & Trim(Dong) & "','" & Trim(Ho) & "', '" & Format(Now, "yyyymmdd") & "', '" & Trim(DtEnd) & "', '', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "', '', '', '', '')"
                adoConn.Execute "Insert Into tb_reg (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, REG_DATE) VALUES ('" & Trim(CarNo) & "', '" & Trim(CarModel) & "','" & Trim(CarGubun) & "','0', '" & Trim(DriverName) & "','" & Trim(DriverPhone) & "','" & Trim(Dong) & "','" & Trim(Ho) & "', '" & Format(Now, "yyyymmdd") & "', '" & Trim(DtEnd) & "', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            Else
                'AdoConn.Execute "UPDATE tb_member SET RF_NO = '" & Mid(strLine, 1, 8) & "' WHERE RF_CODE = '" & Mid(strLine, 18, 10) & "'"
                Call DataLogger(" [CSV Proc]    " & "CAR_NO 중복등록 시도 : " & CarNo)
                'Call Err_doc("CAR_NO 중복등록 시도 : " & CarNo)
            End If
            DoEvents
            Set rs = Nothing
        End If
        i = i + 1
    Loop
    Close #LF

    sngStop = Timer
    TxtStop = CStr(sngStop): TxtStop.Refresh

    'TxtLineStatus = "파일 오픈 종료"

    TxtDuration = sngStop - sngStart
    TxtDuration.Refresh

    MsgBox "파일의 마지막 Line : " & strLine
    
erro_p:
    Call DataLogger(" [CSV Proc]    " & Err.Description & "    " & "Insert Into tb_reg (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, REG_DATE) VALUES ('" & Trim(CarNo) & "', '" & Trim(CarModel) & "','" & Trim(CarGubun) & "','0', '" & Trim(DriverName) & "','" & Trim(DriverPhone) & "','" & Trim(Dong) & "','" & Trim(Ho) & "', '" & Format(Now, "yyyymmdd") & "', '" & Trim(DtEnd) & "', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')")
    
End Sub

Private Sub cmd_Exit_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    
    If Right(File1.Path, 1) = "\" Then
        lbl_TextPath.Caption = File1.Path & File1.filename
    Else
        lbl_TextPath.Caption = File1.Path & "\" & File1.filename
    End If
End Sub

Private Sub Form_Load()
    
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.

End Sub



