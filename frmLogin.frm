VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  '없음
   Caption         =   "ParkingManager™"
   ClientHeight    =   3885
   ClientLeft      =   40230
   ClientTop       =   6090
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2295.387
   ScaleMode       =   0  '사용자
   ScaleWidth      =   5619.59
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_Login 
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
      IMEMode         =   3  '사용 못함
      Index           =   1
      Left            =   2640
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1830
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확 인"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1350
      TabIndex        =   2
      Top             =   2880
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취 소"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3180
      TabIndex        =   3
      Top             =   2880
      Width           =   1500
   End
   Begin VB.TextBox txt_Login 
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  '사용 못함
      Index           =   0
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1380
      Width           =   2325
   End
   Begin VB.Shape Shape1 
      Height          =   3885
      Left            =   0
      Top             =   0
      Width           =   5985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   390
      TabIndex        =   6
      Top             =   210
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "아이디 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   1410
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "암    호 :"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   1890
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Call ShowMenu(txt_Login(0), txt_Login(1))

End Sub


Private Sub txt_Login_Change(Index As Integer)
On Error GoTo Err_P
    If (Len(txt_Login(0).text) >= txt_Login(0).MaxLength) Then
        txt_Login(0).text = Left(txt_Login(0).text, txt_Login(0).MaxLength)
        If (txt_Login(1).Enabled = True) Then
            txt_Login(1).SetFocus
        End If
    End If
    
    If (Len(txt_Login(0).text) = txt_Login(0).MaxLength) Then
        If (Len(txt_Login(1).text) = txt_Login(1).MaxLength) Then
            Call ShowMenu(txt_Login(0), txt_Login(1))
        End If
    End If
    
Exit Sub
    
Err_P:
    Call DataLogger(" [txt_Login_Change] " & Err.Description)
End Sub

Private Sub Form_Load()

Me.Caption = Me.Caption & "  " & "Version " & App.Major & "." & App.Minor & "." & App.Revision

Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.

Call Check_ID_Table 'ID 테이블 데이터 없을경우 기본값 저장

End Sub


Private Sub Check_ID_Table()
    Dim rs As Recordset
    Dim qry As String
    Dim bQryResult As Boolean

On Error GoTo Err_P

    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, "Select * from tb_id", False)
    If (bQryResult = False) Then
        Call DataLogger("[FrmLogin Check_ID_Table]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    If (rs.EOF) Then

        Set rs = New ADODB.Recordset
        
        qry = "INSERT INTO tb_id VALUES ('11111111', '11111111', '총괄관리자', '입출차조회', '정기권관리', '정기권이력', '근무자관리', '환경설정', '무인정산기', '','','','', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        bQryResult = DataBaseQueryExec(adoConn, qry, False)
        qry = "INSERT INTO tb_id VALUES ('22222222', '22222222', '관리자', '입출차조회', '정기권관리', '정기권이력', '근무자관리', '환경설정', '무인정산기', '','','','', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        bQryResult = DataBaseQueryExec(adoConn, qry, False)
        qry = "INSERT INTO tb_id VALUES ('33333333', '33333333', '운영자', '입출차조회', '정기권관리', '정기권이력', '근무자관리', '환경설정', '', '','','','', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        bQryResult = DataBaseQueryExec(adoConn, qry, False)
        If (bQryResult = False) Then
            Call DataLogger("[Check_ID_Table]    " & "네트워크 및 DB 점검바랍니다")
            Set rs = Nothing
            Exit Sub
        End If
    End If
    Set rs = Nothing
Exit Sub
    
Err_P:
    Set rs = Nothing
    Call DataLogger("[Check_ID_Table] Err : " & Err.Description)
End Sub

Public Sub ShowMenu(sID As String, sPW As String)
    Dim rs As Recordset
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim sPasswordEncode As String
    Dim myForm As Form
    
    If (sID <> "") Then
    
        sPasswordEncode = EncodeNDE01(sPW, "www.jawootek.com")  '복호화
        
        Set rs = New ADODB.Recordset
        'rs.Open sQry, adoConn
        'sQry = "SELECT * FROM tb_id where ID='" & sID & "' AND PASSWORD = '" & sPasswordEncode & "' "
        sQry = "SELECT * FROM tb_id where ID='" & sID & "' "
        bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
        If (bQryResult = False) Then
            Call DataLogger("[FrmLogin]    " & "네트워크 및 DB 점검바랍니다")
            Exit Sub
        End If
        
        
        If Not (rs.EOF) Then
        
            If (sPasswordEncode = "" & rs!PassWord) Then
    
                Glo_Login_ID = sID
                Glo_Login_PW = sPW
                Glo_Login_GUBUN = rs!Gubun

                Select Case Glo_Screen_No
                    Case 6
                        Set myForm = FrmG6_23
                    Case 4
                        Set myForm = FrmG4Mini
                    Case 2
                        Set myForm = Jung
                    Case 1
                        Set myForm = FrmG1
                End Select
                
                Dim aMenuList As Variant
                If (Glo_Screen_No = 6) Then
                    aMenuList = Array(rs!MENU1, rs!MENU2, rs!MENU3, rs!MENU4, rs!MENU5, rs!MENU6, rs!MENU7, rs!MENU8, rs!MENU9, rs!MENU10)
                    Call ReleaseMainMenuButton6Form(myForm, aMenuList)
                    Set myForm = Nothing
                    Unload Me
                Else
                    aMenuList = Array(rs!MENU1, rs!MENU2, rs!MENU3, rs!MENU4, rs!MENU5, rs!MENU6, rs!MENU7, rs!MENU8, rs!MENU9, rs!MENU10)
                    Call ReleaseMainMenuButton(myForm, aMenuList)
                    Set myForm = Nothing
                    Unload Me
                End If
                
            Else
                MsgBox "정확히 입력해주세요."
                txt_Login(0) = ""
                txt_Login(1) = ""
                txt_Login(0).SetFocus
            End If
                
        Else
            MsgBox "정확히 입력해주세요."
            txt_Login(0) = ""
            txt_Login(1) = ""
            txt_Login(0).SetFocus
        End If
        
        Set rs = Nothing
    Else
        If (Glo_Screen_No = 4) Then
                FrmG4Mini.Lblbutton(0).Enabled = False
                FrmG4Mini.Imgbutton(0).Enabled = False
                FrmG4Mini.Lblbutton(2).Enabled = False
                FrmG4Mini.Imgbutton(2).Enabled = False
                FrmG4Mini.Lblbutton(3).Enabled = False
                FrmG4Mini.Imgbutton(3).Enabled = False
                FrmG4Mini.Lblbutton(4).Enabled = False
                FrmG4Mini.Imgbutton(4).Enabled = False
                FrmG4Mini.Lblbutton(5).Enabled = False
                FrmG4Mini.Imgbutton(5).Enabled = False
                FrmG4Mini.Lblbutton(8).Enabled = False
                FrmG4Mini.Imgbutton(8).Enabled = False
                FrmG4Mini.Lblbutton(6).Enabled = True
                FrmG4Mini.Imgbutton(6).Enabled = True
        ElseIf (Glo_Screen_No = 2) Then
                Jung.Lblbutton(0).Enabled = False
                Jung.Imgbutton(0).Enabled = False
                Jung.Lblbutton(2).Enabled = False
                Jung.Imgbutton(2).Enabled = False
                Jung.Lblbutton(3).Enabled = False
                Jung.Imgbutton(3).Enabled = False
                Jung.Lblbutton(4).Enabled = False
                Jung.Imgbutton(4).Enabled = False
                Jung.Lblbutton(5).Enabled = False
                Jung.Imgbutton(5).Enabled = False
                Jung.Lblbutton(8).Enabled = False
                Jung.Imgbutton(8).Enabled = False
                Jung.Lblbutton(6).Enabled = True
                Jung.Imgbutton(6).Enabled = True
        ElseIf (Glo_Screen_No = 1) Then
                FrmG1.Lblbutton(0).Enabled = False
                FrmG1.Imgbutton(0).Enabled = False
                FrmG1.Lblbutton(2).Enabled = False
                FrmG1.Imgbutton(2).Enabled = False
                FrmG1.Lblbutton(3).Enabled = False
                FrmG1.Imgbutton(3).Enabled = False
                FrmG1.Lblbutton(4).Enabled = False
                FrmG1.Imgbutton(4).Enabled = False
                FrmG1.Lblbutton(5).Enabled = False
                FrmG1.Imgbutton(5).Enabled = False
                FrmG1.Lblbutton(8).Enabled = False
                FrmG1.Imgbutton(8).Enabled = False
                FrmG1.Lblbutton(6).Enabled = True
                FrmG1.Imgbutton(6).Enabled = True
        End If
    End If
End Sub






