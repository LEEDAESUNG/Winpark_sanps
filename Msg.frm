VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form Msg_Box 
   Appearance      =   0  '���
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  '���� ����
   Caption         =   " Parking Manager�� ���"
   ClientHeight    =   3885
   ClientLeft      =   11370
   ClientTop       =   4515
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Msg.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   5970
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   600
      Left            =   2280
      TabIndex        =   1
      Top             =   2865
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1058
      _StockProps     =   78
      Caption         =   "Ȯ ��"
      ForeColor       =   16777215
      RoundedCorners  =   0   'False
      Outline         =   0   'False
      Picture         =   "Msg.frx":1C5C
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   390
      TabIndex        =   2
      Top             =   255
      Width           =   5280
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Height          =   1440
      Left            =   1530
      TabIndex        =   0
      Top             =   1185
      Width           =   3765
   End
End
Attribute VB_Name = "Msg_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
End Sub

Private Sub SSCommand1_Click()
    Unload Me
End Sub
