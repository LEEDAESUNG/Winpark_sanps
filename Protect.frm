VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form protect 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   BorderStyle     =   1  '���� ����
   Caption         =   "Password"
   ClientHeight    =   3795
   ClientLeft      =   6495
   ClientTop       =   3465
   ClientWidth     =   5985
   DrawStyle       =   5  '����
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Protect.frx":0000
   ScaleHeight     =   3795
   ScaleWidth      =   5985
   Begin VB.TextBox Text1 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      IMEMode         =   3  '��� ����
      Left            =   1845
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2145
      Width           =   2355
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   2250
      TabIndex        =   1
      Top             =   2850
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "��  ��"
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
      Picture         =   "Protect.frx":25F7
   End
End
Attribute VB_Name = "protect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = &H1B) Then
        Text1.Text = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SSCommand1.ForeColor = vbWhite
End Sub

Private Sub SSCommand1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SSCommand1.ForeColor = vbYellow
End Sub

Private Sub Text1_Change()
    Dim PassWord As String
    Dim tmp As Long
    Dim i, dummy%
    
    PassWord = Get_Ini("System Config", "��й�ȣ", "")
    If (Len(Text1.Text) = Text1.MaxLength And (PassWord = Trim(Text1.Text))) Then
        Select Case HostType
            Case 0
                 Jung.cmd_menu(4).Caption = "��ȣ���"
            Case 1
                Jung.cmd_menu(4).Caption = "��ȣ���"
            Case 2, 3
                FrmG4Mini.cmd_menu(2).Caption = "��ȣ���"
                Put_Ini "System Config", "��ȣ���", "False"
                Put_Ini "System Config", "��й�ȣ", ""
                For i = 0 To 7
                    'If (i <> 1) Then
                        FrmG4Mini.cmd_menu(i).Enabled = True
                    'End If
                Next i
                
            Case Else
                 Jung.cmd_menu(4).Caption = "��ȣ���"
        End Select
        Text1.Text = ""
        Me.Hide
    End If
End Sub

Private Sub SSCommand1_Click()
    Unload Me
End Sub
