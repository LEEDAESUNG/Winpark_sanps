VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form MBox 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   BorderStyle     =   1  '���� ����
   Caption         =   " Parking Manager�� ���"
   ClientHeight    =   3885
   ClientLeft      =   13185
   ClientTop       =   6750
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Msg1.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   5985
   Begin Threed.SSCommand Command1 
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   2850
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Ȯ ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "Msg1.frx":1D4B
   End
   Begin Threed.SSCommand Command1 
      Cancel          =   -1  'True
      Height          =   615
      Index           =   1
      Left            =   3165
      TabIndex        =   2
      Top             =   2850
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "Msg1.frx":209C
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   26.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   1410
      TabIndex        =   4
      Top             =   840
      Width           =   4245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   435
      TabIndex        =   3
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   1470
      TabIndex        =   0
      Top             =   1875
      Width           =   4005
   End
End
Attribute VB_Name = "MBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click(Index As Integer)
    If (Index = 1) Then
        Glo_MsgRet = False
    Else
        Glo_MsgRet = True
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()
Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Glo_MsgRet = False

End Sub

