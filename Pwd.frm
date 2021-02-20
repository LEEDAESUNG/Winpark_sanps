VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form Pwd 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   3900
   ClientLeft      =   4200
   ClientTop       =   8160
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Pwd.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   6000
   Begin Threed.SSCommand cmd_cancel 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2265
      TabIndex        =   1
      Top             =   2865
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "취  소"
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
      Picture         =   "Pwd.frx":25F7
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1845
      TabIndex        =   0
      Top             =   2145
      Width           =   2355
   End
End
Attribute VB_Name = "Pwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_cancel_Click()
Pwd_Cancel = True
Unload Me
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
Pwd_Cancel = False

End Sub

Private Sub Text1_Change()
Dim PassWord As String
Dim Dest_Date As String

PassWord = Get_Ini("System Config", "비밀번호", "")
If (Len(Text1.Text) = Text1.MaxLength And (PassWord = Trim(Text1.Text))) Then
        Unload Me
End If
End Sub
