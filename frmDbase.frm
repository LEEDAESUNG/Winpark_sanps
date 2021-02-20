VERSION 5.00
Begin VB.Form frmDbase 
   BorderStyle     =   1  '단일 고정
   Caption         =   " DB 관리"
   ClientHeight    =   9585
   ClientLeft      =   14295
   ClientTop       =   3240
   ClientWidth     =   4485
   Icon            =   "frmDbase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDbase.frx":000C
   ScaleHeight     =   9585
   ScaleWidth      =   4485
   Begin VB.CommandButton Command1 
      Caption         =   "일반권 입차 기록 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   2100
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "일반권 결제 내역 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   135
      TabIndex        =   9
      Top             =   2820
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "정기권 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   4260
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "전체자료 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   135
      TabIndex        =   7
      Top             =   7860
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "정기권 결제 내역 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   135
      TabIndex        =   6
      Top             =   3540
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "정기권 입출 내역 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   135
      TabIndex        =   5
      Top             =   4980
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "할인권 판매 내역 초기화"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   135
      TabIndex        =   4
      Top             =   5700
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   135
      TabIndex        =   3
      Top             =   6420
      Width           =   4200
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   8
      Left            =   135
      TabIndex        =   2
      Top             =   7140
      Width           =   4200
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
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
      Left            =   6390
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1155
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "닫    기"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2760
      TabIndex        =   1
      Top             =   8700
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  '단일 고정
      Caption         =   $"frmDbase.frx":4A7C
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   825
      Left            =   150
      TabIndex        =   11
      Top             =   1080
      Width           =   4215
   End
End
Attribute VB_Name = "frmDbase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private NText(1) As New clsNtext
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Put_Ini "System Config", "자료보관기간", Text1.Text
    Set NText(1) = Nothing
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Response
Dim msg As String
Dim Style


Style = vbYesNo + vbCritical + vbDefaultButton2

Select Case Index
            Case 0
                     msg = "이기능은 모든 일반권 입차기록을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 1
                     msg = "이기능은 모든 일반권 결제내역을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 2
                     msg = "이기능은 모든 정기권 등록내역을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 3
                     msg = "이기능은 모든 데이터베이스 자료를 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 4
                     msg = "이기능은 모든 정기권 결제 기록을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 5
                     msg = "이기능은 모든 정기권 입출 기록을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 6
                     msg = "이기능은 모든 할인권 판매 기록을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 7
                     msg = "이기능은 모든 T_머니출차 기록을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
            Case 8
                     msg = "이기능은 모든 후불교통 기록을 삭제합니다." & Chr$(13) & Chr$(13) & "실행하면 복구할 수 없으므로 주의 하시기 바랍니다." & Chr$(13) & Chr$(13) & "계속하시겠습니까?"
End Select

'성훈
Response = MsgBox(msg, Style, " Parking Manager™")

If Response = vbNo Then
    Exit Sub
End If

'Exit Sub

Me.MousePointer = 11
Select Case Index
       Case 0
            adoConn.Execute "DELETE  FROM ilbacarnin"
            Call Err_doc("호스트 : 일반권 입차 기록 초기화 완료")
       Case 1
            adoConn.Execute "DELETE  FROM ilbancarinout"
            Call Err_doc("호스트 : 출차 기록 초기화 완료")
       Case 2
            adoConn.Execute "DELETE  FROM regcar"
            Call Err_doc("호스트 : 정기권 등록내역 초기화 완료")
       Case 3   '전체
            adoConn.Execute "DELETE FROM ilbancarin"
            adoConn.Execute "DELETE FROM ilbancarinout"
            adoConn.Execute "DELETE FROM regcarinout"
            adoConn.Execute "DELETE FROM regcar"
            adoConn.Execute "DELETE FROM tb_fee"
            adoConn.Execute "DELETE FROM tb_coupon_sale"
            'adoConn.Execute "DELETE FROM charge_dic"
            'adoConn.Execute "DELETE FROM cancleout"
            '성훈 카드결제부분
            'adoConn.Execute "DELETE FROM t_in"
            'adoConn.Execute "DELETE FROM t_out"
            'adoConn.Execute "DELETE FROM after_money"
            Call Err_doc("호스트 : 모든자료 초기화 완료")
       Case 4
            adoConn.Execute "DELETE FROM tb_fee"
            Call Err_doc("호스트 : 정기권 판매내역 초기화 완료")
       Case 5
            adoConn.Execute "DELETE FROM regcarinout"
            Call Err_doc("호스트 : 정기권 입출내역 초기화 완료")
       Case 6
            adoConn.Execute "DELETE FROM tb_coupon_sale"
            Call Err_doc("호스트 : 할인권 판매내역 초기화 완료")
'       Case 7
'           adoConn.Execute "DELETE FROM t_out"
'       Case 8
'            adoConn.Execute "DELETE FROM after_money"
End Select
Me.MousePointer = 0

End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    '폼을 가운데에 놓습니다.
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

Set NText(1).NText = Me.Text1
    Text1.Text = Get_Ini("System Config", "자료보관기간", "12")

End Sub

