VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Begin VB.Form FrmId 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   12660
   ClientLeft      =   5640
   ClientTop       =   2010
   ClientWidth     =   15345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmId.frx":0000
   ScaleHeight     =   12660
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1410
      Left            =   0
      TabIndex        =   44
      Top             =   11250
      Width           =   15360
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�̿��� ID ��� ����"
      Height          =   4110
      Left            =   -15
      TabIndex        =   66
      Top             =   7125
      Width           =   15360
      Begin VB.TextBox txt_PaidMoney 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12975
         MaxLength       =   8
         TabIndex        =   123
         Text            =   "0000"
         Top             =   2580
         Width           =   1005
      End
      Begin VB.CommandButton cmd_FreeCharge 
         BackColor       =   &H0080C0FF&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   14085
         Style           =   1  '�׷���
         TabIndex        =   122
         Top             =   2070
         Width           =   915
      End
      Begin VB.CommandButton cmd_PaidCharge 
         BackColor       =   &H0080C0FF&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   14085
         Style           =   1  '�׷���
         TabIndex        =   121
         Top             =   2580
         Width           =   915
      End
      Begin VB.TextBox txt_PaidCount 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11715
         MaxLength       =   8
         TabIndex        =   119
         Text            =   "0000"
         Top             =   2580
         Width           =   705
      End
      Begin VB.TextBox txt_FreeCount 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11715
         MaxLength       =   8
         TabIndex        =   117
         Text            =   "0000"
         Top             =   2115
         Width           =   705
      End
      Begin VB.CommandButton cmd_InitPassword 
         BackColor       =   &H0080C0FF&
         Caption         =   "��й�ȣ �ʱ�ȭ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7950
         Style           =   1  '�׷���
         TabIndex        =   116
         ToolTipText     =   """1234""���� ��й�ȣ �ʱ�ȭ �մϴ�"
         Top             =   2145
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Code 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   102
         Top             =   2280
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Partner 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5295
         MaxLength       =   8
         TabIndex        =   101
         Top             =   2280
         Width           =   1545
      End
      Begin VB.TextBox txt_DC 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   13365
         MaxLength       =   8
         TabIndex        =   100
         Text            =   "���ΰ�5"
         ToolTipText     =   "���ΰ��� �Է����ּ���"
         Top             =   3465
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Desc 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   13350
         MaxLength       =   8
         TabIndex        =   99
         Text            =   "���θ�Ī5"
         ToolTipText     =   "���θ�Ī�� �Է����ּ���"
         Top             =   3060
         Width           =   1545
      End
      Begin VB.TextBox txt_DC 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   10650
         MaxLength       =   8
         TabIndex        =   98
         Text            =   "���ΰ�4"
         ToolTipText     =   "���ΰ��� �Է����ּ���"
         Top             =   3465
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Desc 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   10650
         MaxLength       =   8
         TabIndex        =   97
         Text            =   "���θ�Ī4"
         ToolTipText     =   "���θ�Ī�� �Է����ּ���"
         Top             =   3060
         Width           =   1545
      End
      Begin VB.TextBox txt_DC 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7950
         MaxLength       =   8
         TabIndex        =   96
         Text            =   "���ΰ�3"
         ToolTipText     =   "���ΰ��� �Է����ּ���"
         Top             =   3465
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Desc 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7950
         MaxLength       =   8
         TabIndex        =   95
         Text            =   "���θ�Ī3"
         ToolTipText     =   "���θ�Ī�� �Է����ּ���"
         Top             =   3060
         Width           =   1545
      End
      Begin VB.TextBox txt_DC 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5295
         MaxLength       =   8
         TabIndex        =   94
         Text            =   "���ΰ�2"
         ToolTipText     =   "���ΰ��� �Է����ּ���"
         Top             =   3465
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Desc 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5295
         MaxLength       =   8
         TabIndex        =   93
         Text            =   "���θ�Ī2"
         ToolTipText     =   "���θ�Ī�� �Է����ּ���"
         Top             =   3060
         Width           =   1545
      End
      Begin VB.TextBox txt_DC_Desc 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   91
         Text            =   "���θ�Ī1"
         ToolTipText     =   "���θ�Ī�� �Է����ּ���"
         Top             =   3060
         Width           =   1545
      End
      Begin VB.TextBox txt_DC 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   92
         Text            =   "���ΰ�1"
         ToolTipText     =   "���ΰ��� �Է����ּ���"
         Top             =   3465
         Width           =   1545
      End
      Begin VB.ComboBox cmb_DC_Gubun 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmId.frx":4AF2
         Left            =   7830
         List            =   "FrmId.frx":4AF4
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   90
         Top             =   2625
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�޴�10"
         Height          =   315
         Index           =   9
         Left            =   8070
         TabIndex        =   84
         Top             =   1425
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�޴�9"
         Height          =   315
         Index           =   8
         Left            =   6510
         TabIndex        =   83
         Top             =   1425
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������"
         Height          =   315
         Index           =   7
         Left            =   4920
         TabIndex        =   82
         Top             =   1425
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������"
         Height          =   315
         Index           =   6
         Left            =   3270
         TabIndex        =   81
         Top             =   1425
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���������"
         Height          =   315
         Index           =   5
         Left            =   1560
         TabIndex        =   80
         Top             =   1425
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ȯ�漳��"
         Height          =   315
         Index           =   4
         Left            =   8070
         TabIndex        =   79
         Top             =   1110
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ٹ��ڰ���"
         Height          =   315
         Index           =   3
         Left            =   6510
         TabIndex        =   78
         Top             =   1110
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�湮����"
         Height          =   315
         Index           =   2
         Left            =   4920
         TabIndex        =   77
         Top             =   1110
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����ǰ���"
         Height          =   315
         Index           =   1
         Left            =   3270
         TabIndex        =   76
         Top             =   1110
         Width           =   1485
      End
      Begin VB.CheckBox chk_Menu 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��������ȸ"
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   75
         Top             =   1110
         Width           =   1485
      End
      Begin VB.TextBox txt_password 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  '��� ����
         Left            =   4395
         MaxLength       =   8
         TabIndex        =   69
         Top             =   510
         Width           =   1545
      End
      Begin VB.TextBox txt_id 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         MaxLength       =   8
         TabIndex        =   68
         Top             =   510
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmId.frx":4AF6
         Left            =   7560
         List            =   "FrmId.frx":4AF8
         TabIndex        =   67
         Text            =   "Combo1"
         Top             =   510
         Width           =   2325
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   8
         Left            =   14025
         TabIndex        =   85
         Top             =   405
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmId.frx":4AFA
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   9
         Left            =   12915
         TabIndex        =   86
         Top             =   405
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmId.frx":4E4B
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   10
         Left            =   11805
         TabIndex        =   87
         Top             =   405
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmId.frx":519C
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   11
         Left            =   10695
         TabIndex        =   88
         Top             =   405
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�ʱ�ȭ"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmId.frx":54ED
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   615
         Index           =   12
         Left            =   10695
         TabIndex        =   115
         ToolTipText     =   "����Ʈ�� �α��� ����ڿ��� ��� �޼��� �����մϴ�."
         Top             =   1155
         Visible         =   0   'False
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "�޼�������"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmId.frx":583E
      End
      Begin VB.Label lbl_NowPaidPoint 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   10695
         TabIndex        =   127
         Top             =   2625
         Width           =   720
      End
      Begin VB.Label lbl_NowFreePoint 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   10695
         TabIndex        =   126
         Top             =   2205
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   9690
         X2              =   15000
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '����
         Caption         =   "�ݾ�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   12525
         TabIndex        =   124
         Top             =   2625
         Width           =   540
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '����
         Caption         =   "��������Ʈ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   9705
         TabIndex        =   120
         Top             =   2625
         Width           =   900
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         Caption         =   "��������Ʈ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   9705
         TabIndex        =   118
         Top             =   2175
         Width           =   900
      End
      Begin VB.Label lbl_DC 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ΰ�5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   4
         Left            =   12285
         TabIndex        =   114
         Top             =   3525
         Width           =   1020
      End
      Begin VB.Label lbl_DC_Desc 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���θ�Ī5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   4
         Left            =   12285
         TabIndex        =   113
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lbl_DC 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ΰ�4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   3
         Left            =   9585
         TabIndex        =   112
         Top             =   3525
         Width           =   1020
      End
      Begin VB.Label lbl_DC_Desc 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���θ�Ī4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   3
         Left            =   9585
         TabIndex        =   111
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lbl_DC 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ΰ�3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   2
         Left            =   6900
         TabIndex        =   110
         Top             =   3525
         Width           =   1020
      End
      Begin VB.Label lbl_DC_Desc 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���θ�Ī3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   2
         Left            =   6900
         TabIndex        =   109
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lbl_DC 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ΰ�2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   1
         Left            =   4245
         TabIndex        =   108
         Top             =   3525
         Width           =   1020
      End
      Begin VB.Label lbl_DC_Desc 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���θ�Ī2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   1
         Left            =   4245
         TabIndex        =   107
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lbl_DC 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���ΰ�1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   0
         Left            =   1470
         TabIndex        =   106
         Top             =   3525
         Width           =   1020
      End
      Begin VB.Label lbl_DC_Desc 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���θ�Ī1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Index           =   0
         Left            =   1470
         TabIndex        =   105
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Line Line1 
         X1              =   285
         X2              =   15075
         Y1              =   1965
         Y2              =   1965
      End
      Begin VB.Label lbl_DC_Gubun 
         BackStyle       =   0  '����
         Caption         =   "���α���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   7020
         TabIndex        =   104
         Top             =   2685
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lbl_PName 
         BackStyle       =   0  '����
         Caption         =   "��ü��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   4695
         TabIndex        =   103
         Top             =   2340
         Width           =   1020
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   360
         TabIndex        =   89
         Top             =   2310
         Width           =   1020
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '����
         Caption         =   "�޴� ����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   360
         TabIndex        =   74
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '����
         Caption         =   "�̿��� ID"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   360
         TabIndex        =   73
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "��й�ȣ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   3435
         TabIndex        =   72
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label lbl_PCode 
         BackStyle       =   0  '����
         Caption         =   "��ü�ڵ�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1725
         TabIndex        =   71
         Top             =   2340
         Width           =   1020
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "��        ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   6510
         TabIndex        =   70
         Top             =   510
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " �����˻�"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   16320
      TabIndex        =   62
      Top             =   3600
      Width           =   7455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��/ȣ �˻�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cmbDong 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.ComboBox cmbHo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3525
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txt_tmpCarNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   3270
         TabIndex        =   18
         Top             =   390
         Width           =   1845
      End
      Begin VB.ComboBox cmb_GB 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "FrmId.frx":5B8F
         Left            =   1680
         List            =   "FrmId.frx":5BA2
         TabIndex        =   17
         Text            =   "������ȣ"
         Top             =   390
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�󼼰˻�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin Threed.SSCommand cmd_Search 
         Height          =   705
         Left            =   6045
         TabIndex        =   22
         Top             =   360
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   1244
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Picture         =   "FrmId.frx":5BDC
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   3015
         TabIndex        =   64
         Top             =   1125
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "ȣ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   11.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   4860
         TabIndex        =   63
         Top             =   1125
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin VB.Frame frm_Week 
      Appearance      =   0  '���
      BackColor       =   &H00404040&
      Caption         =   " ���� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   16320
      TabIndex        =   43
      Top             =   1695
      Width           =   6405
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   6
         Left            =   5430
         TabIndex        =   32
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   5
         Left            =   4590
         TabIndex        =   31
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   4
         Left            =   3765
         TabIndex        =   30
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   3
         Left            =   2925
         TabIndex        =   29
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   2
         Left            =   2085
         TabIndex        =   28
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "ȭ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   27
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   26
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
   End
   Begin VB.Frame frm_Rotation 
      Appearance      =   0  '���
      BackColor       =   &H00404040&
      Caption         =   " ���� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   16320
      TabIndex        =   38
      Top             =   2670
      Width           =   7185
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "10 ����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   3
         Left            =   5550
         TabIndex        =   42
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "5 ����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   2
         Left            =   3900
         TabIndex        =   41
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "2 ����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   1
         Left            =   2250
         TabIndex        =   40
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00404040&
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   39
         Top             =   360
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.ComboBox cmb_Search 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "FrmId.frx":5F2D
      Left            =   16320
      List            =   "FrmId.frx":5F2F
      TabIndex        =   37
      Text            =   "�˻�����"
      Top             =   1230
      Width           =   2715
   End
   Begin VB.TextBox txt_Dong 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   16320
      TabIndex        =   46
      Top             =   150
      Width           =   2325
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   5670
      Left            =   -15
      TabIndex        =   25
      Top             =   1410
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   10001
      View            =   3
      Arrange         =   2
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
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
      Height          =   585
      Index           =   0
      Left            =   13905
      TabIndex        =   24
      Top             =   765
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmId.frx":5F31
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   5
      Left            =   10950
      TabIndex        =   23
      Top             =   765
      Visible         =   0   'False
      Width           =   1065
      _Version        =   65536
      _ExtentX        =   1879
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "����"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmId.frx":6282
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   540
      Index           =   6
      Left            =   19320
      TabIndex        =   33
      Top             =   1170
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   952
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   14737632
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
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   7
      Left            =   16335
      TabIndex        =   45
      Top             =   600
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ���� ��� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   16320
      TabIndex        =   47
      Top             =   5310
      Width           =   15255
      Begin VB.ComboBox cmb_Rotation 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmId.frx":65D3
         Left            =   9705
         List            =   "FrmId.frx":65DD
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   11
         Top             =   1500
         Width           =   2325
      End
      Begin VB.CommandButton cmd_Month 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1���� ����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7890
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   49
         Top             =   2415
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox cmb_Gubun 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmId.frx":65EF
         Left            =   9690
         List            =   "FrmId.frx":65F1
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   9
         Top             =   480
         Width           =   2325
      End
      Begin VB.TextBox txt_CarNo 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1365
         TabIndex        =   0
         Top             =   975
         Width           =   2325
      End
      Begin VB.TextBox txt_Object 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9690
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   975
         Width           =   5385
      End
      Begin VB.TextBox txt_Ho 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5490
         TabIndex        =   6
         Top             =   1440
         Width           =   2325
      End
      Begin VB.TextBox txt_Phone 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1365
         TabIndex        =   2
         Top             =   1905
         Width           =   2325
      End
      Begin VB.TextBox txt_Name 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1365
         TabIndex        =   1
         Top             =   1440
         Width           =   2325
      End
      Begin VB.TextBox txt_CarModel 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1365
         TabIndex        =   3
         Top             =   2385
         Width           =   2325
      End
      Begin VB.TextBox txt_Num 
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1365
         TabIndex        =   48
         Top             =   495
         Width           =   2865
      End
      Begin VB.ComboBox cmb_Dong 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmId.frx":65F3
         Left            =   5490
         List            =   "FrmId.frx":65F5
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   990
         Width           =   2340
      End
      Begin MSMask.MaskEdBox MaskEdBox_Start 
         Height          =   375
         Left            =   5490
         TabIndex        =   7
         Top             =   1920
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   2
         Left            =   13950
         TabIndex        =   14
         Top             =   2235
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
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
         Picture         =   "FrmId.frx":65F7
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   4
         Left            =   12840
         TabIndex        =   13
         Top             =   2235
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
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
         Picture         =   "FrmId.frx":6948
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   1
         Left            =   11730
         TabIndex        =   12
         Top             =   2235
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   14737632
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
         Picture         =   "FrmId.frx":6C99
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   3
         Left            =   10620
         TabIndex        =   15
         Top             =   2235
         Width           =   1110
         _Version        =   65536
         _ExtentX        =   1958
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�ʱ�ȭ"
         ForeColor       =   14737632
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
         Picture         =   "FrmId.frx":6FEA
      End
      Begin MSMask.MaskEdBox MaskEdBox_End 
         Height          =   375
         Left            =   5490
         TabIndex        =   8
         Top             =   2400
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox_Fee 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """\""#,##0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   2
         EndProperty
         Height          =   375
         Left            =   5490
         TabIndex        =   4
         Top             =   495
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "�����뺸"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   8595
         TabIndex        =   65
         Top             =   1515
         Width           =   1185
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "��     ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   8610
         TabIndex        =   61
         Top             =   525
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "��     ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   4530
         TabIndex        =   60
         Top             =   540
         Width           =   960
      End
      Begin VB.Label lbl_dept 
         BackStyle       =   0  '����
         Caption         =   "����1 / ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Index           =   2
         Left            =   4290
         TabIndex        =   59
         Top             =   1005
         Width           =   1200
      End
      Begin VB.Label lbl_clas 
         BackStyle       =   0  '����
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Index           =   0
         Left            =   255
         TabIndex        =   58
         Top             =   2385
         Width           =   1020
      End
      Begin VB.Label lbl_Phone 
         BackStyle       =   0  '����
         Caption         =   "��ȭ��ȣ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   255
         TabIndex        =   57
         Top             =   1905
         Width           =   1020
      End
      Begin VB.Label lbl_StartDate 
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   4530
         TabIndex        =   56
         Top             =   1935
         Width           =   960
      End
      Begin VB.Label lbl_Object 
         BackStyle       =   0  '����
         Caption         =   "��     ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   8610
         TabIndex        =   55
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label lbl_EndDate 
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   4530
         TabIndex        =   54
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label lbl_dept 
         BackStyle       =   0  '����
         Caption         =   "����2 / ȣ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Index           =   3
         Left            =   4290
         TabIndex        =   53
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lbl_Num 
         BackStyle       =   0  '����
         Caption         =   "����Ͻ�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   255
         TabIndex        =   52
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label lbl_Name 
         BackStyle       =   0  '����
         Caption         =   "��      ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   255
         TabIndex        =   51
         Top             =   1425
         Width           =   1020
      End
      Begin VB.Label lbl_CarNo 
         BackStyle       =   0  '����
         Caption         =   "������ȣ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   255
         TabIndex        =   50
         Top             =   975
         Width           =   1020
      End
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   13
      Left            =   12075
      TabIndex        =   125
      Top             =   765
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "������"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmId.frx":733B
   End
   Begin VB.Label lbl_title 
      BackColor       =   &H00404040&
      Caption         =   "�̿��� ���̵� ��� ����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   315
      TabIndex        =   34
      Top             =   120
      Width           =   5160
   End
   Begin VB.Label lbl_COUNT 
      BackStyle       =   0  '����
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1470
      TabIndex        =   36
      Top             =   1005
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��ϰǼ� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   435
      TabIndex        =   35
      Top             =   1005
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "FrmId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TXT_ID_TMP As String
Dim TXT_PASSWORD_TMP As String
Dim CAR_NO_TMP As String
Dim PART_NAME_TMP As String
Dim RegQry As String
Const WebDC_COUNT As Integer = 5 '������ ������




Private Sub chk_Menu_Click(Index As Integer)
    Dim i As Integer
    If (chk_Menu(Index).Caption = "���������" And chk_Menu(Index).value = 1) Then
        For i = 0 To 9
            If (chk_Menu(i).Caption = "��������") Then
                chk_Menu(i).value = 0
                Exit For
            End If
        Next
    ElseIf (chk_Menu(Index).Caption = "��������" And chk_Menu(Index).value = 1) Then
        For i = 0 To 9
            If (chk_Menu(i).Caption = "���������") Then
                chk_Menu(i).value = 0
                Exit For
            End If
        Next
    End If
    
    
    Call Disable_WebDC
    
On Error Resume Next
    For i = 0 To chk_Menu.Count - 1
        If (chk_Menu(i).Caption = "������" And chk_Menu(i).value = 1) Then
            Call Enable_WebDC
            Exit For
        End If
    Next i
    
End Sub

'�׽�Ʈ
Private Sub AllDeviceSendMsg()
    
    Dim rsID As ADODB.Recordset
    Dim bQryResult As Boolean
    
On Error GoTo Err_P
    
    Set rsID = New ADODB.Recordset
    bQryResult = DataBaseQuery(rsID, adoConn, "SELECT * FROM tb_id", False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[AllDeviceSendMsg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    Do While Not (rsID.EOF)
        If (rsID!MENU1 = "�ۻ��" Or rsID!MENU2 = "�ۻ��" Or rsID!MENU3 = "�ۻ��" Or rsID!MENU4 = "�ۻ��" Or rsID!MENU5 = "�ۻ��" Or rsID!MENU6 = "�ۻ��" Or rsID!MENU7 = "�ۻ��" Or rsID!MENU8 = "�ۻ��" Or rsID!MENU9 = "�ۻ��" Or rsID!MENU10 = "�ۻ��") Then
            Call OneDeviceSendMsg(rsID!ID)
        End If
        rsID.MoveNext
    Loop
    Set rsID = Nothing
    
    Exit Sub
    
Err_P:
    Call DataLogger("[AllDeviceSendMsg] Err:" & Err.Description)
End Sub


Private Sub OneDeviceSendMsg(sID As String)
    Dim rs As ADODB.Recordset
    Dim bQryResult As Boolean
    Dim sMsg As String
    Dim sMsg_UTF8() As Byte
    Dim Title As String
    Dim Body As String
    
On Error GoTo Err_P

    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM tb_devices WHERE ID = '" & sID & "' ", False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[OneDeviceSendMsg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    
    Do While Not (rs.EOF)
        sMsg = "{" & Chr(34) & "target" & Chr(34) & ":[" '  {"target":[
        sMsg = sMsg & Chr(34) & rs!token & Chr(34) & ","
        sMsg = Left(sMsg, Len(sMsg) - 1)
        
        Title = "test"
        Body = "test"
        
        sMsg = sMsg & "]," & Chr(34) & "title" & Chr(34) & ":" & Chr(34) & Title & Chr(34) & "," & Chr(34) & "body" & Chr(34) & ":" & Chr(34) & Body & Chr(34) & "}"
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    If (Len(sMsg) > 0) Then
        sMsg_UTF8 = StringToUTF8BytesArray(sMsg)

        FrmTcpServer.WinsockS_Devices.SendData sMsg_UTF8
        Call DataLogger("[DeviceSendMsg] sID " & "[Title] :" & Title & "[Body] :" & Body)
    End If
    
    Exit Sub
    
Err_P:
    Call DataLogger("[OneDeviceSendMsg] Err:" & Err.Description)
    
End Sub


Private Sub cmd_OneDeviceSendMsg_Click()
    
    'FrmTcpServer.WinsockS_Devices.SendData
End Sub


'������ ��������
Private Sub cmd_FreeCharge_Click()

    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim sPcode As String
    Dim nFreePoint, nAddFreePoint, nSumFreePoint As Integer
    Dim nPaidPoint As Integer
    Dim nPaidPoint_Money As Long
    Dim sStoreID As String
    Dim sLog As String

On Error GoTo Err_P

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ǽ� �� �ݾ� �� üũ ����
    txt_FreeCount.text = Trim(txt_FreeCount.text)
    If (txt_FreeCount.text = "") Then txt_FreeCount.text = "0"
    
    If IsNumeric(txt_FreeCount.text) = False Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "���ڸ� �Է��ϼ���."
        Msg_Box.Show 1
        
        txt_FreeCount.text = "0"
        txt_FreeCount.SetFocus
        Exit Sub
    End If

    If txt_FreeCount.text = "0" Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "������ ��������Ʈ�� �Է��ϼ���."
        Msg_Box.Show 1
        txt_FreeCount.SetFocus
        Exit Sub
    End If
    '����Ǽ� �� �ݾ� �� üũ ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Set rs = New ADODB.Recordset
    sQry = "SELECT * FROM tb_id WHERE ID = '" & txt_id & "' LIMIT 1"
    bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId FreeCharge]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Set rs = Nothing
        Exit Sub
    End If
    
    If (Not rs.EOF) Then
            
            MBox.Label2.Caption = "������"
            MBox.Label3.Caption = txt_id.text
            MBox.Label1.Caption = "�������� �����Ͻðڽ��ϱ�?"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
               
                sPcode = "" & rs!SEQ
                sStoreID = "" & rs!ID
               
                Set rs2 = New ADODB.Recordset
                sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sPcode & "'"
                rs2.Open sQry, adoConn
                If Not (rs2.EOF) Then
                
                    nFreePoint = rs2!FREE_POINT
                    nAddFreePoint = CInt(txt_FreeCount.text)
                    nSumFreePoint = nFreePoint + nAddFreePoint
                    
                    
                    sQry = "UPDATE  tb_partner  SET  FREE_POINT = " & nSumFreePoint & " WHERE SEQ = '" & sPcode & "' "
                    adoConn.Execute sQry
                    
                    
                    sLog = "[������ ��������]" & sPcode & "." & sStoreID & ":" & nAddFreePoint & "(��)"
                    
                    sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, PAID_POINT, PAID_POINT_CHARGEMONEY, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sPcode & "', " & nAddFreePoint & ", 0,0,'" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sPcode & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                    adoConn.Execute sQry
                    
                    
                    '��������Ʈ ���
                    lbl_NowFreePoint.Caption = "[" & nSumFreePoint & "]"
                    
                    
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sLog, 0
                    Call DataLogger("[FrmId FreeCharge]    " & sLog)
                End If
                
                Set rs2 = Nothing
            Else
                
            End If
            
    Else
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":������ �˻� �����Դϴ�. �ٽ� �õ����ּ���(E00001)", 0
        Call DataLogger("[FrmId FreeCharge]    " & txt_id & ":������ �˻� �����Դϴ�. �ٽ� �õ����ּ���(E00001)")
        Set rs = Nothing
        Exit Sub
    End If
    Set rs = Nothing
    
    
    txt_FreeCount.text = "0"
    
    Exit Sub
    
Err_P:
    Set rs = Nothing
    Set rs2 = Nothing
    
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":�����߻�. �ٽ� �õ����ּ���(E00002)" & " " & Err.Description, 0
    Call DataLogger("[FrmId FreeCharge]    " & txt_id & ":�����߻�. �ٽ� �õ����ּ���(E00002)" & " " & Err.Description)
    
End Sub

'������ ��������
Private Sub cmd_PaidCharge_Click()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim sPcode As String
    Dim nPaidPoint, nAddPaidPoint, nSumPaidPoint As Integer
    Dim nPaidPoint_Money As Long
    Dim sStoreID As String
    Dim sLog As String

'On Error GoTo Err_p
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����Ǽ� �� �ݾ� �� üũ ����
    txt_PaidCount.text = Trim(txt_PaidCount.text)
    txt_PaidMoney.text = Trim(txt_PaidMoney.text)
    If (txt_PaidCount.text = "") Then txt_PaidCount.text = "0"
    If (txt_PaidMoney.text = "") Then txt_PaidMoney.text = "0"
    
    If IsNumeric(txt_PaidCount.text) = False Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "���ڸ� �Է��ϼ���."
        Msg_Box.Show 1
        
        txt_PaidCount.text = "0"
        txt_PaidCount.SetFocus
        Exit Sub
    End If
    If IsNumeric(txt_PaidMoney.text) = False Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "���ڸ� �Է��ϼ���."
        Msg_Box.Show 1
        
        txt_PaidMoney.text = "0"
        txt_PaidMoney.SetFocus
        Exit Sub
    End If
    If txt_PaidCount.text = "0" Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "������ ��������Ʈ�� �Է��ϼ���."
        Msg_Box.Show 1
        txt_PaidCount.SetFocus
        Exit Sub
    End If
    '����Ǽ� �� �ݾ� �� üũ ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Set rs = New ADODB.Recordset
    sQry = "SELECT * FROM tb_id WHERE ID = '" & txt_id & "' LIMIT 1"
    bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId PaidCharge]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Set rs = Nothing
        Exit Sub
    End If
    
    If (Not rs.EOF) Then
            
            MBox.Label2.Caption = "������"
            MBox.Label3.Caption = txt_id.text
            MBox.Label1.Caption = "�������� �����Ͻðڽ��ϱ�?"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
               
                sPcode = "" & rs!SEQ
                sStoreID = "" & rs!ID
               
                Set rs2 = New ADODB.Recordset
                sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sPcode & "'"
                
                bQryResult = DataBaseQuery(rs2, adoConn, sQry, False)
                If (bQryResult = False) Then
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                    Call DataLogger("[FrmId PaidCharge]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                    Set rs = Nothing
                    Exit Sub
                End If
                
                nPaidPoint = rs2!PAID_POINT
                nAddPaidPoint = txt_PaidCount.text
                nSumPaidPoint = nPaidPoint + nAddPaidPoint
                nPaidPoint_Money = txt_PaidMoney
                
                sQry = "UPDATE  tb_partner  SET  PAID_POINT = " & nSumPaidPoint & " WHERE SEQ = '" & sPcode & "' "
                adoConn.Execute sQry
                

                sLog = "[������ ��������]" & sPcode & "." & sStoreID & ":" & nAddPaidPoint & "(��)"
                
                sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, PAID_POINT, PAID_POINT_CHARGEMONEY, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sPcode & "', 0, " & nAddPaidPoint & ", " & nPaidPoint_Money & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                adoConn.Execute sQry
                
                sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sPcode & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                adoConn.Execute sQry
                
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sLog, 0
                Call DataLogger("[FrmId PaidCharge]    " & sLog)


                '��������Ʈ ���
                lbl_NowPaidPoint.Caption = "[" & nSumPaidPoint & "]"

                Set rs2 = Nothing
            Else
                
            End If
            
    Else
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":������ �˻� �����Դϴ�. �ٽ� �õ����ּ���(E00003)", 0
        Call DataLogger("[FrmId PaidCharge]    " & txt_id & ":������ �˻� �����Դϴ�. �ٽ� �õ����ּ���(E00003)")
        Set rs = Nothing
        Exit Sub
    End If
    Set rs = Nothing
    
    
    txt_PaidCount.text = "0"
    txt_PaidMoney.text = "0"
    
    Exit Sub
    
Err_P:
    Set rs = Nothing
    Set rs2 = Nothing
    
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description, 0
    Call DataLogger("[FrmId PaidCharge]    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description)
End Sub

Private Sub cmd_InitPassword_Click()
    
    Dim qry As String
    Dim bQryResult As Boolean
    Dim sInitPW As String
    Dim sPWEncode  As String
    
    MBox.Label3.Caption = TXT_ID_TMP
    MBox.Label1.Caption = "�����Ͻ� ��й�ȣ�� '1234' �� �ʱ�ȭ �մϴ�." & vbCrLf & vbCrLf & " �����Ͻðڽ��ϱ�?"
    MBox.Label2.Caption = "��й�ȣ �ʱ�ȭ"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
       If (TXT_ID_TMP <> "") Then
            sInitPW = "1234"
            sPWEncode = EncodeNDE01(sInitPW, "www.jawootek.com")   '��ȣȭ
            
            qry = "UPDATE  tb_id  SET  PASSWORD = '" & sPWEncode & "', MENU10 = '" & sInitPW & "' WHERE ID = '" & TXT_ID_TMP & "' "
            'adoConn.Execute Qry
            bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
            If (bQryResult = False) Then
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                Call DataLogger("[FrmId InitPassword]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                Exit Sub
            Else
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & TXT_ID_TMP & ":��й�ȣ�� �ʱ�ȭ �߽��ϴ�", 0
                Call DataLogger("[FrmId InitPassword]    " & TXT_ID_TMP & ":��й�ȣ�� �ʱ�ȭ �߽��ϴ�")
            End If
            
            Call Combo_Gubun
            Call ListView_REG_Draw
            Call ListView_REG_SQL
        End If
    End If
    
    
End Sub

Private Sub Combo1_Click()
    If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��(��й�ȣ ��������)
        txt_password = ""
        txt_password.Enabled = False
        txt_password.BackColor = &HC0C0C0
        'Call MsgBox("�����ȣ�� �����Ҽ� �����ϴ�", vbInformation Or vbMsgBoxSetForeground, "��й�ȣ ����")
    Else
        txt_password.Enabled = True
        txt_password.BackColor = &H80000005
    End If
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rs As Recordset
    Dim qry As String
    Dim bView As Boolean

    
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.


    
    
    'cmd_Button(8).Enabled = True
    
'''    RegQry = "SELECT * From tb_id"
'''
'''    bView = Able_WebDC
'''
'''    If (Glo_Login_GUBUN = "�Ѱ�������") Then
'''        Combo1.AddItem ("�Ѱ�������")
'''        Combo1.AddItem ("������")
'''        Combo1.AddItem ("���")
'''        If (bView = True) Then
'''            Combo1.AddItem ("��Ʈ��")
'''        End If
'''
'''    ElseIf (Glo_Login_GUBUN = "������") Then
'''        Combo1.AddItem ("������")
'''        Combo1.AddItem ("���")
'''        If (bView = True) Then
'''            Combo1.AddItem ("��Ʈ��")
'''        End If
'''        RegQry = RegQry + " WHERE GUBUN = '������' OR GUBUN = '���' "
'''
'''    ElseIf (Glo_Login_GUBUN = "���") Then
'''        Combo1.AddItem ("���")
'''        For i = 0 To 9
'''            chk_Menu(i).Enabled = False
'''        Next
'''        RegQry = RegQry + " WHERE ID = '" & Glo_Login_ID & "' "
'''    End If
'''
    Call Clear_Field
    Call Clear_WebDC
    Call Disable_WebDC
    Call View_WebDC
    Call View_GuestReg 'üũ�ڽ� Enable/Disable
    Call Combo_Gubun
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    'cmb_GB.ListIndex = 0
    
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ���̵� ���/���� ����...!!", 0
    Call DataLogger("[ID Formload]    " & "���̵� ���/���� ����...!!")

End Sub

Private Sub Combo_Gubun()
    Dim i As Integer
    Dim rs As Recordset
    Dim qry As String
    Dim bView As Boolean
    
    bView = False
    Combo1.Clear
    
    RegQry = "SELECT * From tb_id "
    
    If (Glo_Login_GUBUN = "�Ѱ�������") Then
        Combo1.AddItem ("�Ѱ�������")
        Combo1.AddItem ("������")
        Combo1.AddItem ("���")
        
        bView = Able_WebDC
    
    ElseIf (Glo_Login_GUBUN = "������") Then
        Combo1.AddItem ("������")
        Combo1.AddItem ("���")
        
        bView = Able_WebDC

        RegQry = RegQry + " WHERE GUBUN = '������' OR GUBUN = '���' "
        
    ElseIf (Glo_Login_GUBUN = "���") Then
        Combo1.AddItem ("���")
        For i = 0 To 9
            chk_Menu(i).Enabled = False
        Next
        RegQry = RegQry + " WHERE ID = '" & Glo_Login_ID & "' "
    End If
    
    If (bView = True) Then
        'Combo1.AddItem ("��Ʈ��")
        Set rs = New ADODB.Recordset
        rs.Open "SELECT GUBUN From tb_id group by GUBUN", adoConn
        Do While Not (rs.EOF)
            If (rs!Gubun <> "�Ѱ�������" And rs!Gubun <> "������" And rs!Gubun <> "���") Then
                Combo1.AddItem rs!Gubun
            End If
            'Debug.Print rs!Gubun
            rs.MoveNext
        Loop
        Set rs = Nothing
    End If
End Sub


'�湮�������� üũ��ư enable/disable
Private Sub View_GuestReg()
'''    On Error Resume Next
'''
'''    Dim bCheck As Boolean
'''    bCheck = False
'''
'''    Set rs = New ADODB.Recordset
'''    rs.Open "SELECT Content from tb_config WHERE NAME = 'GuestCarReg'", adoConn
'''    If (Not rs.EOF) Then
'''        If (rs!Content = "Y") Then
'''            bCheck = True
'''        End If
'''    End If
'''    Set rs = Nothing
'''
'''    If (bCheck = True) Then
'''        chk_Menu(2).Enabled = True
'''    Else
'''        chk_Menu(2).Enabled = False
'''    End If
    
    If (Glo_GuestReg_YN = "Y") Then
        chk_Menu(2).Enabled = True
    Else
        chk_Menu(2).Enabled = False
    End If
    
End Sub


'������ ��� ����� ��쿡�� ������â�� �����ش�
Private Sub View_WebDC()
'    Dim bView As Boolean
'
'    bView = Able_WebDC
'
'    If (bView = True) Then
'        '������ ��ɻ���� ��
'        Me.height = 13095
'        List1.Top = 11250
'
'        chk_Menu(7).Enabled = True '������ üũ�ڽ� enable
'        cmd_Button(13).Enabled = True '�����ι�ư
'        cmd_Button(13).Visible = True
'    Else
'        '������ ��� ������ ��
'        Me.height = 10935
'        List1.Top = 9104
'
'        chk_Menu(7).Enabled = False '������ üũ�ڽ� disable
'        cmd_Button(13).Enabled = False '�����ι�ư
'        cmd_Button(13).Visible = False
'    End If
    If (Glo_WebDC_YN = "Y") Then
        Me.height = 13095
        List1.Top = 11250
        
        chk_Menu(7).Enabled = True '������ üũ�ڽ� enable
        cmd_Button(13).Enabled = True '�����ι�ư
        cmd_Button(13).Visible = True
    Else
        Me.height = 10935
        List1.Top = 9104

        chk_Menu(7).Enabled = False '������ üũ�ڽ� disable
        cmd_Button(13).Enabled = False '�����ι�ư
        cmd_Button(13).Visible = False
    End If
End Sub

Private Function Able_WebDC() As Boolean
    Dim rs As Recordset
    Dim qry As String

    Able_WebDC = False
    
    On Error Resume Next

    Set rs = New ADODB.Recordset
    qry = "SELECT Content FROM tb_config WHERE (NAME = 'WebDC' AND CONTENT = 'Y') "
    rs.Open qry, adoConn
    
    If (Not (rs.EOF)) Then
        Able_WebDC = True
    End If
    
    Set rs = Nothing
End Function

Private Sub Enable_WebDC()
    Dim i As Integer
    
    'txt_DC_Code.Enabled = True
    txt_DC_Partner.Enabled = True
    cmb_DC_Gubun.Enabled = True
    lbl_PCode.Enabled = True
    lbl_PName.Enabled = True
    lbl_DC_Gubun.Enabled = True
    cmd_InitPassword.Enabled = True
    cmd_FreeCharge.Enabled = True
    cmd_PaidCharge.Enabled = True
    txt_FreeCount.Enabled = True
    txt_PaidCount.Enabled = True
    txt_PaidMoney.Enabled = True
    
    For i = 0 To WebDC_COUNT - 1
        lbl_DC(i).Enabled = True
        lbl_DC_Desc(i).Enabled = True
        txt_DC_Desc(i).Enabled = True
        txt_DC(i).Enabled = True
    Next i
End Sub

Private Sub Disable_WebDC()
    Dim i As Integer
    
    'txt_DC_Code.Enabled = False
    txt_DC_Partner.Enabled = False
    cmb_DC_Gubun.Enabled = False
    lbl_PCode.Enabled = False
    lbl_PName.Enabled = False
    lbl_DC_Gubun.Enabled = False
    cmd_InitPassword.Enabled = False
    cmd_FreeCharge.Enabled = False
    cmd_PaidCharge.Enabled = False
    txt_FreeCount.Enabled = False
    txt_PaidCount.Enabled = False
    txt_PaidMoney.Enabled = False
    
    For i = 0 To WebDC_COUNT - 1
        lbl_DC(i).Enabled = False
        lbl_DC_Desc(i).Enabled = False
        txt_DC_Desc(i).Enabled = False
        txt_DC(i).Enabled = False
    Next i
End Sub

Private Sub Clear_WebDC()
    Dim i As Integer
    
    cmb_DC_Gubun.Clear
    cmb_DC_Gubun.AddItem "�ð�(��)"
    cmb_DC_Gubun.AddItem "�ݾ�(��)"
    cmb_DC_Gubun.ListIndex = 0
    
    txt_DC_Code.text = ""
    txt_DC_Partner.text = ""
    
    For i = 0 To WebDC_COUNT - 1
        txt_DC(i).text = ""
        txt_DC_Desc(i).text = ""
    Next i
    
    
'    bChk = False
'    For i = 0 To 9
'        If (chk_Menu(i).Caption = "������" And chk_Menu(i).value = 1) Then
'            bChk = True
'            Exit For
'        End If
'    Next i
    
    
    txt_DC_Code.Enabled = False
    txt_DC_Partner.Enabled = False
    cmb_DC_Gubun.Enabled = False
    For i = 0 To WebDC_COUNT - 1
        txt_DC_Desc(i).Enabled = False
        txt_DC(i).Enabled = False
    Next i
    
    
    txt_FreeCount.text = "0"
    txt_PaidCount.text = "0"
    txt_PaidMoney.text = "0"
    
    lbl_NowFreePoint = "" '���� ��������Ʈ ���
    lbl_NowPaidPoint = "" '���� ��������Ʈ ���
        
End Sub


Public Sub ListView_REG_SQL()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim bQryResult As Boolean
    Dim iIdx As Integer
    Dim sPasswordEncode As String
    Dim bWebDC As Boolean
    
    On Error GoTo Err_P

    bWebDC = Able_WebDC
    
    INDEX_NO = 1
    Set rs = New ADODB.Recordset
    'rs.Open RegQry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, RegQry, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    lbl_COUNT = rs.RecordCount
    
Repeat:

    Do While Not (rs.EOF)
    

        If (bWebDC = False) Then
            If (Not (rs!Gubun = "�Ѱ�������" Or rs!Gubun = "������" Or rs!Gubun = "���")) Then
                rs.MoveNext
                GoTo Repeat
            End If
        End If
        
        
        
        Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
        
        iIdx = 1
        itmX.SubItems(iIdx) = "" & rs!ID: iIdx = iIdx + 1
        
        'itmX.SubItems(iIdx) = "" & rs!PassWord: iIdx = iIdx + 1
        If (rs!Gubun = "�Ѱ�������" Or rs!Gubun = "������" Or rs!Gubun = "���") Then
            sPasswordEncode = DecodeNDE01(rs!PassWord, "www.jawootek.com")  '��ȣȭ
            itmX.SubItems(iIdx) = "" & sPasswordEncode: iIdx = iIdx + 1
        Else
            itmX.SubItems(iIdx) = "": iIdx = iIdx + 1
        End If
        
        
        itmX.SubItems(iIdx) = "" & rs!Gubun: iIdx = iIdx + 1
        
        
        '��Ʈ�� ���̺�
        Set rs2 = New ADODB.Recordset
        'rs2.Open "SELECT * FROM tb_partner WHERE ID='" & rs!ID & "' ", adoConn
        rs2.Open "SELECT * FROM tb_partner WHERE SEQ='" & rs!SEQ & "' ", adoConn
        If (Not (rs2.EOF)) Then
            itmX.SubItems(iIdx) = "" & rs2!PNAME: iIdx = iIdx + 1 '��ü��
        Else
            itmX.SubItems(iIdx) = "": iIdx = iIdx + 1
        End If
        
        If (rs!Gubun = "�Ѱ�������" Or rs!Gubun = "������" Or rs!Gubun = "���") Then
            itmX.SubItems(iIdx) = "" & rs!MENU1: iIdx = iIdx + 1
            itmX.SubItems(iIdx) = "" & rs!MENU2: iIdx = iIdx + 1
            itmX.SubItems(iIdx) = "" & rs!MENU3: iIdx = iIdx + 1
        Else
            itmX.SubItems(iIdx) = "����:" & rs2!FREE_POINT: iIdx = iIdx + 1 'partner
            itmX.SubItems(iIdx) = "����:" & rs2!PAID_POINT: iIdx = iIdx + 1
            itmX.SubItems(iIdx) = "�ڵ�����:" & rs2!FREE_AUTOPOINT: iIdx = iIdx + 1
        End If
        itmX.SubItems(iIdx) = "" & rs!MENU4: iIdx = iIdx + 1
        itmX.SubItems(iIdx) = "" & rs!MENU5: iIdx = iIdx + 1
        itmX.SubItems(iIdx) = "" & rs!MENU6: iIdx = iIdx + 1
        itmX.SubItems(iIdx) = "" & rs!MENU7: iIdx = iIdx + 1
        itmX.SubItems(iIdx) = "" & rs!MENU8: iIdx = iIdx + 1
        'itmX.SubItems(iIdx) = "" & rs!MENU9: iIdx = iIdx + 1
        'itmX.SubItems(iIdx) = "" & rs!MENU10: iIdx = iIdx + 1
        itmX.SubItems(iIdx) = "" & rs!REG_DATE: iIdx = iIdx + 1

        Set rs2 = Nothing
        
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    Set rs = Nothing

Exit Sub

Err_P:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Description, 0
    Call DataLogger("[FrmId ListView_REG_SQL]    " & Err.Description)
    
End Sub

Public Sub ListView_REG_Draw()
Dim Column_to_size As Integer

With Me
    Call ListViewExtended(.ListView_REG)
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No   "
    .ListView_REG.ColumnHeaders.Add , , " ���̵�      "
    .ListView_REG.ColumnHeaders.Add , , " ��й�ȣ    "
    .ListView_REG.ColumnHeaders.Add , , " ����                  "
    .ListView_REG.ColumnHeaders.Add , , " ��ü��      "
    .ListView_REG.ColumnHeaders.Add , , " �޴�1       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�2       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�3       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�4       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�5       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�6       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�7       "
    .ListView_REG.ColumnHeaders.Add , , " �޴�8       "
    '.ListView_REG.ColumnHeaders.Add , , " �޴�9       "
    '.ListView_REG.ColumnHeaders.Add , , " �޴�10      "
    .ListView_REG.ColumnHeaders.Add , , " ��ϳ�¥                      "
    .ListView_REG.ColumnHeaders.Add , , "    "
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With

End Sub

Private Sub ListView_REG_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim i As Integer
    With ListView_REG
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

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
    
    ListView_REG.SetFocus
    txt_id = ListView_REG.SelectedItem.SubItems(1)
    Call Search_Record
    
    If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��(��й�ȣ ��������)
        txt_password = ""
        txt_password.Enabled = False
        txt_password.BackColor = &HC0C0C0
    Else
        txt_password.Enabled = True
        txt_password.BackColor = &H80000005
    End If
    
End Sub

Public Sub Clear_Field()
Dim i As Long

    cmd_Button(8).Enabled = False   '����
    cmd_Button(9).Enabled = False    '����
    cmd_Button(10).Enabled = True  '���
    cmd_Button(11).Enabled = True   '�ʱ�ȭ

    
    txt_id.text = ""
    txt_password.text = ""

    TXT_ID_TMP = ""
    TXT_PASSWORD_TMP = ""
    For i = 0 To 9
        chk_Menu(i).value = 0
    Next i

    On Error Resume Next
    txt_id.SetFocus
    Combo1.ListIndex = 0
    
    
    txt_FreeCount = ""
    txt_PaidCount = ""
    txt_PaidMoney = ""
End Sub

'������ ����
Sub Delete_Record()
    Dim sQry As String
    Dim bQryResult As Boolean
    
On Error GoTo Err_P
    'adoConn.Execute "DELETE FROM tb_id WHERE ID = '" & txt_id & "'"
    sQry = "DELETE FROM tb_id WHERE ID = '" & txt_id & "'"
    bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmID Delete_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    adoConn.Execute "DELETE FROM tb_partner  WHERE ID = '" & txt_id & "'"
    
    
    
    '�Ʒ��� ���� ������
    'adoConn.Execute "INSERT INTO tb_reg_log VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun.Text & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & txt_Object & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.Text & "', '" & Glo_PartName & "', '����', '" & Glo_Login_ID & "')"
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & "    �α��� ���̵� ���� �Ϸ�", 0
    Call DataLogger("[Delete Button]    " & txt_id & "    �α��� ���̵� ���� �Ϸ�")

    Call Combo_Gubun
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    '20200601
    '����̽� ����
    '����̽�(����Ʈ��)���� tb_id �� ID/password �α����Ұ�� tb_devices �� Insert ��
    'ȣ��Ʈ���α׷��� tb_id �� ID������ ��� tb_device �Բ� ����ó����
    adoConn.Execute "DELETE FROM tb_devices WHERE ID = '" & txt_id & "'"
    
    
    Exit Sub

Err_P:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Description, 0
    Call DataLogger("[FrmId Delete_Record]    " & Err.Description)
End Sub

Sub Insert_Record()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim qry As String
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim sPasswordEncode As String
    Dim sPartnerPasswordEncode As String
    Dim sPW As String
    
    Dim sMenu1 As String
    Dim sMenu2 As String
    Dim sMenu3 As String
    Dim sMenu4 As String
    Dim sMenu5 As String
    Dim sMenu6 As String
    Dim sMenu7 As String
    Dim sMenu8 As String
    Dim sMenu9 As String
    Dim sMenu10 As String
    
    Dim i As Integer
    Dim sDC_Code As String
    Dim sDC_Partner As String
    Dim sDC_Gubun As String
    Dim iDC(5) As Long
    Dim iDC_De(5) As String
    

On Error GoTo Err_P

    Set rs = New ADODB.Recordset
    qry = "SELECT * FROM tb_id WHERE ID = '" & txt_id & "' LIMIT 1"
    'rs.Open Qry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    If (Not rs.EOF) Then
        Msg_Box.Label2.Caption = "������ ���̽� ����"
        Msg_Box.Label1.Caption = "�ߺ��� ID�� ��������ʽ��ϴ�."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    
    sMenu1 = ""
    sMenu2 = ""
    sMenu3 = ""
    sMenu4 = ""
    sMenu5 = ""
    sMenu6 = ""
    sMenu7 = ""
    sMenu8 = ""
    sMenu9 = ""
    sMenu10 = ""
    If (chk_Menu(0).value = 1) Then
        sMenu1 = chk_Menu(0).Caption
    End If
    If (chk_Menu(1).value = 1) Then
        sMenu2 = chk_Menu(1).Caption
    End If
    If (chk_Menu(2).value = 1) Then
        sMenu3 = chk_Menu(2).Caption
    End If
    If (chk_Menu(3).value = 1) Then
        sMenu4 = chk_Menu(3).Caption
    End If
    If (chk_Menu(4).value = 1) Then
        sMenu5 = chk_Menu(4).Caption
    End If
    If (chk_Menu(5).value = 1) Then
        sMenu6 = chk_Menu(5).Caption
    End If
    If (chk_Menu(6).value = 1) Then
        sMenu7 = chk_Menu(6).Caption
    End If
    If (chk_Menu(7).value = 1) Then
        sMenu8 = chk_Menu(7).Caption
    End If
    If (chk_Menu(8).value = 1) Then
        sMenu9 = chk_Menu(8).Caption
    End If
    If (chk_Menu(9).value = 1) Then
        sMenu10 = chk_Menu(9).Caption
    End If
    
    
    
    sDC_Partner = LeftH(Trim(txt_DC_Partner.text), 16)
    
    If (cmb_DC_Gubun.text = "�ð�(��)") Then
        sDC_Gubun = "T"
    Else
        sDC_Gubun = "M"
    End If
    
    For i = 0 To UBound(iDC) - 1
        iDC_De(i) = txt_DC_Desc(i).text
        iDC(i) = Val(txt_DC(i).text)
    Next i
    
    
    
    If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��
        sPW = "1234"
    Else
        sPW = txt_password
    End If
    sPasswordEncode = EncodeNDE01(sPW, "www.jawootek.com") '��ȣȭ
    

    If (TXT_ID_TMP = "") Then '�űԵ��
        'INSERT
        sQry = "INSERT INTO tb_id (ID, PASSWORD, GUBUN, MENU1, MENU2, MENU3, MENU4, MENU5, MENU6, MENU7, MENU8, MENU9, MENU10, REG_DATE ) VALUES ('" & txt_id & "', '" & sPasswordEncode & "', '" & Combo1.text & "', '" & sMenu1 & "', '" & sMenu2 & "','" & sMenu3 & "','" & sMenu4 & "','" & sMenu5 & "', '" & sMenu6 & "','" & sMenu7 & "','" & sMenu8 & "','" & sMenu9 & "','" & sPW & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
        If (bQryResult = False) Then
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
            Call DataLogger("[FrmID Insert_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
            Exit Sub


        Else
        
            Set rs2 = New ADODB.Recordset
            bQryResult = DataBaseQuery(rs2, adoConn, "Select SEQ as IDSeq from tb_id WHERE ID = '" & txt_id & "' ", False)
            If (bQryResult = False) Then
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                Exit Sub
            End If
            
            If (Not rs2.EOF) Then
                If (Len("" & rs2!IDSeq) > 0) Then
                    sDC_Code = rs2!IDSeq
                End If
            End If
            Set rs2 = Nothing
            
            adoConn.Execute "INSERT INTO tb_partner (SEQ, ID, PCODE, PNAME, PGUBUN, PDC1, PDC1_DESC, PDC2, PDC2_DESC, PDC3, PDC3_DESC, PDC4, PDC4_DESC, PDC5, PDC5_DESC, REG_DATE ) VALUES ('" & sDC_Code & "', '" & txt_id & "', '', '" & sDC_Partner & "', '" & sDC_Gubun & "', " & iDC(0) & ",'" & iDC_De(0) & "', " & iDC(1) & ",'" & iDC_De(1) & "', " & iDC(2) & ",'" & iDC_De(2) & "', " & iDC(3) & ",'" & iDC_De(3) & "', " & iDC(4) & ",'" & iDC_De(4) & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
            
            
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & "    �α��� ���̵� ��� �Ϸ�", 0
            Call DataLogger("[LogIn Button]    " & txt_id & "    �α��� ���̵� ��� �Ϸ�")
        
        End If
        
        
        
    Else
    

        If (TXT_ID_TMP <> txt_id.text) Then '���� �α��� ���̵� �����ϸ�
            If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��(��й�ȣ ��������)
                sQry = "UPDATE tb_id  SET  ID = '" & txt_id & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "'  WHERE ID = '" & TXT_ID_TMP & "' "
            Else
                sQry = "UPDATE tb_id  SET  ID = '" & txt_id & "', PASSWORD = '" & sPasswordEncode & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "', MENU10 = '" & txt_password & "'  WHERE ID = '" & TXT_ID_TMP & "' "
            End If
            bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
            If (bQryResult = False) Then
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                Call DataLogger("[FrmID Insert_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                Exit Sub
            End If
            
        Else
            If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��(��й�ȣ ��������)
                sQry = "UPDATE tb_id  SET  GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "'  WHERE ID = '" & TXT_ID_TMP & "' "
            Else
                sQry = "UPDATE tb_id  SET  PASSWORD = '" & sPasswordEncode & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "', MENU10 = '" & txt_password & "' WHERE ID = '" & TXT_ID_TMP & "' "
            End If
            bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
            If (bQryResult = False) Then
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                Call DataLogger("[FrmID Insert_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                Exit Sub
            End If

        End If
        
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & "    �α��� ���̵� ���� �Ϸ�", 0
        Call DataLogger("[LogIn Button]    " & txt_id & "    �α��� ���̵� ���� �Ϸ�")
    End If
    
    Call Combo_Gubun
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    Exit Sub

Err_P:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Description, 0
    Call DataLogger("[FrmId Insert_Record]    " & Err.Description)

End Sub


Sub Update_Record()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim qry As String
    Dim bQryResult As Boolean
    Dim sPasswordEncode As String
    
    Dim sMenu1 As String
    Dim sMenu2 As String
    Dim sMenu3 As String
    Dim sMenu4 As String
    Dim sMenu5 As String
    Dim sMenu6 As String
    Dim sMenu7 As String
    Dim sMenu8 As String
    Dim sMenu9 As String
    Dim sMenu10 As String
    
    Dim i As Integer
    Dim sDC_Code As String
    Dim sDC_Partner As String
    Dim sDC_Gubun As String
    Dim iDC(5) As Long
    Dim iDC_De(5) As String
    
    
On Error GoTo Err_P

    Set rs = New ADODB.Recordset
    qry = "SELECT * FROM tb_id WHERE ID = '" & TXT_ID_TMP & "' LIMIT 1"
    'rs.Open Qry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    If (rs.EOF) Then
        Msg_Box.Label2.Caption = "������ ���̽� ����"
        Msg_Box.Label1.Caption = "�ٽ� �������ּ���"
        Msg_Box.Show 1
        Exit Sub
    End If


    sMenu1 = ""
    sMenu2 = ""
    sMenu3 = ""
    sMenu4 = ""
    sMenu5 = ""
    sMenu6 = ""
    sMenu7 = ""
    sMenu8 = ""
    sMenu9 = ""
    sMenu10 = ""
    If (chk_Menu(0).value = 1) Then
        sMenu1 = chk_Menu(0).Caption
    End If
    If (chk_Menu(1).value = 1) Then
        sMenu2 = chk_Menu(1).Caption
    End If
    If (chk_Menu(2).value = 1) Then
        sMenu3 = chk_Menu(2).Caption
    End If
    If (chk_Menu(3).value = 1) Then
        sMenu4 = chk_Menu(3).Caption
    End If
    If (chk_Menu(4).value = 1) Then
        sMenu5 = chk_Menu(4).Caption
    End If
    If (chk_Menu(5).value = 1) Then
        sMenu6 = chk_Menu(5).Caption
    End If
    If (chk_Menu(6).value = 1) Then
        sMenu7 = chk_Menu(6).Caption
    End If
    If (chk_Menu(7).value = 1) Then
        sMenu8 = chk_Menu(7).Caption
    End If
    If (chk_Menu(8).value = 1) Then
        sMenu9 = chk_Menu(8).Caption
    End If
    If (chk_Menu(9).value = 1) Then
        sMenu10 = chk_Menu(9).Caption
    End If

    
    sDC_Code = Format(Left(txt_DC_Code.text, 4), "0000")
    sDC_Partner = LeftH(Trim(txt_DC_Partner.text), 16)
    
    If (cmb_DC_Gubun.text = "�ð�(��)") Then
        sDC_Gubun = "T"
    Else
        sDC_Gubun = "M"
    End If
    
    For i = 0 To UBound(iDC) - 1
        iDC_De(i) = "" & txt_DC_Desc(i).text
        iDC(i) = Val(txt_DC(i).text)
    Next i
    
    sPasswordEncode = EncodeNDE01(txt_password, "www.jawootek.com")  '��ȣȭ
    
    If (TXT_ID_TMP <> txt_id.text) Then '���� �α��� ���̵� �����ϸ�
        If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��(��й�ȣ �������)
            qry = "UPDATE tb_id     SET ID = '" & txt_id & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "' WHERE ID = '" & TXT_ID_TMP & "' "
        Else
            'qry = "UPDATE tb_id     SET ID = '" & txt_id & "', PASSWORD = '" & sPasswordEncode & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "', MENU10 = '" & sMenu10 & "' WHERE ID = '" & TXT_ID_TMP & "' "
            qry = "UPDATE tb_id     SET ID = '" & txt_id & "', PASSWORD = '" & sPasswordEncode & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "', MENU10 = '" & txt_password & "' WHERE ID = '" & TXT_ID_TMP & "' "
        End If
        
        'adoConn.Execute Qry
        bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
        If (bQryResult = False) Then
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
            Call DataLogger("[FrmId Update_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
            Exit Sub
            
        Else
        
            Set rs2 = New ADODB.Recordset
            bQryResult = DataBaseQuery(rs2, adoConn, "Select PCode from tb_partner where ID='" & TXT_ID_TMP & "' LIMIT 1", False)
            If (bQryResult = False) Then
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                Exit Sub
            End If
            If (Not rs2.EOF) Then
                'adoConn.Execute "UPDATE tb_partner SET ID='" & txt_id & "',PCODE='" & sDC_Code & "',PNAME='" & sDC_Partner & "',PGUBUN='" & sDC_Gubun & "',PDC1=" & iDC(0) & ",PDC1_DESC='" & iDC_De(0) & "',PDC2=" & iDC(1) & ",PDC2_DESC='" & iDC_De(1) & "',PDC3=" & iDC(2) & ",PDC3_DESC='" & iDC_De(2) & "',PDC4=" & iDC(3) & ",PDC4_DESC='" & iDC_De(3) & "',PDC5=" & iDC(4) & ",PDC5_DESC='" & iDC_De(4) & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "'  WHERE ID = '" & TXT_ID_TMP & "' "
                adoConn.Execute "UPDATE tb_partner SET ID='" & txt_id & "',PNAME='" & sDC_Partner & "',PGUBUN='" & sDC_Gubun & "',PDC1=" & iDC(0) & ",PDC1_DESC='" & iDC_De(0) & "',PDC2=" & iDC(1) & ",PDC2_DESC='" & iDC_De(1) & "',PDC3=" & iDC(2) & ",PDC3_DESC='" & iDC_De(2) & "',PDC4=" & iDC(3) & ",PDC4_DESC='" & iDC_De(3) & "',PDC5=" & iDC(4) & ",PDC5_DESC='" & iDC_De(4) & "', REG_DATE='" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "'  WHERE ID = '" & TXT_ID_TMP & "' "
            Else
                If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then
                    'adoConn.Execute "INSERT INTO tb_partner (ID, PCODE, PNAME, PGUBUN, PDC1, PDC1_DESC, PDC2, PDC2_DESC, PDC3, PDC3_DESC, PDC4, PDC4_DESC, PDC5, PDC5_DESC, REG_DATE ) VALUES ('" & txt_id & "', '" & sDC_Code & "', '" & sDC_Partner & "', '" & sDC_Gubun & "', " & iDC(0) & ",'" & iDC_De(0) & "', " & iDC(1) & ",'" & iDC_De(1) & "', " & iDC(2) & ",'" & iDC_De(2) & "', " & iDC(3) & ",'" & iDC_De(3) & "', " & iDC(4) & ",'" & iDC_De(4) & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') "
                    adoConn.Execute "INSERT INTO tb_partner (SEQ, ID, PCODE, PNAME, PGUBUN, PDC1, PDC1_DESC, PDC2, PDC2_DESC, PDC3, PDC3_DESC, PDC4, PDC4_DESC, PDC5, PDC5_DESC, REG_DATE ) VALUES ('" & sDC_Code & "', '" & txt_id & "', '', '" & sDC_Partner & "', '" & sDC_Gubun & "', " & iDC(0) & ",'" & iDC_De(0) & "', " & iDC(1) & ",'" & iDC_De(1) & "', " & iDC(2) & ",'" & iDC_De(2) & "', " & iDC(3) & ",'" & iDC_De(3) & "', " & iDC(4) & ",'" & iDC_De(4) & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') "
                End If
            End If
            Set rs2 = Nothing
        
        End If
        
        
        
    Else
        If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then '��Ʈ��(��й�ȣ �������)
            qry = "UPDATE tb_id     SET    GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "' WHERE ID = '" & TXT_ID_TMP & "' "
        Else
            'qry = "UPDATE tb_id     SET                        PASSWORD = '" & sPasswordEncode & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "', MENU10 = '" & sMenu10 & "' WHERE ID = '" & TXT_ID_TMP & "' "
            qry = "UPDATE tb_id     SET    PASSWORD = '" & sPasswordEncode & "', GUBUN = '" & Combo1 & "', MENU1 = '" & sMenu1 & "', MENU2 = '" & sMenu2 & "', MENU3 = '" & sMenu3 & "', MENU4 = '" & sMenu4 & "', MENU5 = '" & sMenu5 & "', MENU6 = '" & sMenu6 & "', MENU7 = '" & sMenu7 & "', MENU8 = '" & sMenu8 & "', MENU9 = '" & sMenu9 & "', MENU10 = '" & txt_password & "' WHERE ID = '" & TXT_ID_TMP & "' "
        End If
        
        'adoConn.Execute Qry
        bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
        If (bQryResult = False) Then
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
            Call DataLogger("[FrmId Update_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
            Exit Sub
        
        
        Else
        
        
        
            Set rs2 = New ADODB.Recordset
            bQryResult = DataBaseQuery(rs2, adoConn, "Select PCode from tb_partner where ID='" & TXT_ID_TMP & "' LIMIT 1", False)
            If (bQryResult = False) Then
                List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                Exit Sub
            End If
            If (Not rs2.EOF) Then
                adoConn.Execute "UPDATE tb_partner SET PNAME='" & sDC_Partner & "',PGUBUN='" & sDC_Gubun & "',PDC1=" & iDC(0) & ",PDC1_DESC='" & iDC_De(0) & "',PDC2=" & iDC(1) & ",PDC2_DESC='" & iDC_De(1) & "',PDC3=" & iDC(2) & ",PDC3_DESC='" & iDC_De(2) & "',PDC4=" & iDC(3) & ",PDC4_DESC='" & iDC_De(3) & "',PDC5=" & iDC(4) & ",PDC5_DESC='" & iDC_De(4) & "',REG_DATE='" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "'  WHERE ID = '" & TXT_ID_TMP & "' "
            Else
                If (Combo1 <> "�Ѱ�������" And Combo1 <> "������" And Combo1 <> "���") Then
                    adoConn.Execute "INSERT INTO tb_partner (SEQ, ID, PCODE, PNAME, PGUBUN, PDC1, PDC1_DESC, PDC2, PDC2_DESC, PDC3, PDC3_DESC, PDC4, PDC4_DESC, PDC5, PDC5_DESC, REG_DATE ) VALUES ('" & sDC_Code & "', '" & txt_id & "', '', '" & sDC_Partner & "', '" & sDC_Gubun & "', " & iDC(0) & ",'" & iDC_De(0) & "', " & iDC(1) & ",'" & iDC_De(1) & "', " & iDC(2) & ",'" & iDC_De(2) & "', " & iDC(3) & ",'" & iDC_De(3) & "', " & iDC(4) & ",'" & iDC_De(4) & "','" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "') "
                End If
            End If
            Set rs2 = Nothing
        
        End If
        
        
    End If
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & "    �α��� ���̵� ���� �Ϸ�", 0
    Call DataLogger("[LogIn Button]    " & txt_id & "    �α��� ���̵� ���� �Ϸ�")
    
    Call Combo_Gubun
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    Exit Sub

Err_P:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Description, 0
    Call DataLogger("[FrmId UpdateRecord]    " & Err.Description)
    If (InStr(1, Err.Description, "Duplicate") > 0) Then
        Msg_Box.Label2.Caption = "������ ���̽� ����"
        Msg_Box.Label1.Caption = "�ߺ��� ID�� ��������ʽ��ϴ�."
        Msg_Box.Show 1
    End If
    Call Clear_Field
End Sub


Private Sub cmd_Button_Click(Index As Integer)
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim qry As String
Dim bQryResult As String

Select Case Index
    Case 0  '����
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ���̵� ���/���� ����", 0
        Call DataLogger("[REG Button]    " & txt_CarNo & "    ���̵���/���� ����")
        Unload Me
        'Me.Hide
        Exit Sub
       
    Case 10  '�ű��Է�
        If (TXT_ID_TMP = "" Or TXT_PASSWORD_TMP = "") Then
            If (Data_Error_Check = False) Then
                Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                Msg_Box.Label1.Caption = "�߿��� �׸��� �Է����� �ʾҽ��ϴ�."
                Msg_Box.Show 1
            Else
                Call Insert_Record
                Call Clear_Field
                Call Clear_WebDC
                Call Enable_WebDC
            End If
        Else
            Msg_Box.Label2.Caption = "�ű� ������ �Է� ����"
            Msg_Box.Label1.Caption = "�ű� �����Ͱ� �ƴմϴ�." & vbCrLf & vbCrLf & " �ٽ� �ѹ� Ȯ���ϼ���."
            Msg_Box.Show 1
            Call Clear_Field
        End If
        Exit Sub
    
    Case 8  '����
        If (TXT_ID_TMP = "") Then
           Call Clear_Field
           Exit Sub
        End If
        If (TXT_ID_TMP <> Me.txt_id) Then
            Msg_Box.Label2.Caption = "������ ���� ����"
            Msg_Box.Label1.Caption = "������ �����͸� �ٽ� ������ �ֽʽÿ�."
            Msg_Box.Show 1
            Exit Sub
        End If
        MBox.Label3.Caption = txt_CarNo.text
        MBox.Label1.Caption = "�� �α��� ���̵� ������ �����մϴ�." & vbCrLf & vbCrLf & " �����Ͻðڽ��ϱ�?"
        MBox.Label2.Caption = "�α��� ���̵� ���� ����"
        MBox.Show 1
        If (Glo_MsgRet = True) Then
            '���ID, ��Ʈ��ID ���� ����
           Call Delete_Record
        End If
        Call Clear_Field
        Call Clear_WebDC
        Call Enable_WebDC
        Exit Sub
        
    Case 11   '�ʱ�ȭ
        Call Clear_WebDC
        Call Enable_WebDC
        Call Clear_Field
        Exit Sub
            
    Case 9  '����
        If (TXT_ID_TMP = "") Then
            Msg_Box.Label2.Caption = "�ʵ� ����"
            Msg_Box.Label1.Caption = "�ű� �α��� ���̵� ����ڷ� �Դϴ�." & vbCrLf & vbCrLf & " �ٽ� Ȯ�� �ϼ���."
            Msg_Box.Show 1
            Exit Sub
        Else
            If (txt_id.text = TXT_ID_TMP) Then
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                    Msg_Box.Label1.Caption = "�߿��� �׸��� ���� �Ǵ� �߸� �Է��Ͽ����ϴ�."
                    Msg_Box.Show 1
                Else
                    MBox.Label3.Caption = txt_CarNo.text
                    MBox.Label1.Caption = "�����Ͻ� �α��� ���̵� ������ ����˴ϴ�." & vbCrLf & vbCrLf & " ���� �Ͻðڽ��ϱ�?"
                    MBox.Label2.Caption = "�α��� ���̵� ����"
                    MBox.Show 1
                    If (Glo_MsgRet = True) Then
                       'Call Insert_Record
                       Call Update_Record
                       Call Clear_Field
                       Call Clear_WebDC
                       Call Enable_WebDC
                       'txt_CarNo.SetFocus
                    End If
                End If
            Else
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                    Msg_Box.Label1.Caption = "�߿��� �׸��� ���� �Ǵ� �߸� �Է��Ͽ����ϴ�."
                    Msg_Box.Show 1
                Else
                    MBox.Label3.Caption = txt_CarNo.text
                    MBox.Label1.Caption = "�����Ͻ� �α��� ���̵� ������ ����˴ϴ�." & vbCrLf & vbCrLf & " ���� �Ͻðڽ��ϱ�?"
                    MBox.Label2.Caption = "�α��� ���̵� ����"
                    MBox.Show 1
                    If (Glo_MsgRet = True) Then
                       Call Update_Record
                       Call Clear_Field
                       'txt_CarNo.SetFocus
                    End If
                End If
            End If
        End If
        Exit Sub

    Case 5
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�������_" & cmb_Search.text
        'Call makeexcel(ListView_REG, tmpFileName, "�˻�����")
        Call MakeCSV(ListView_REG, tmpFileName)
        Exit Sub
        
    Case 6
        '����������� �˻�
        Select Case cmb_Search.text
            Case "��ü"
                RegQry = "SELECT * From tb_reg ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"
            Case "�Ⱓ�ʰ�"
                '�Ⱓ�ʰ������˻�
                RegQry = "SELECT * From tb_reg WHERE END_DATE < " & Format(Now, "YYYYMMDD") & " ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"
            Case Else
                RegQry = "SELECT * From tb_reg WHERE CAR_GUBUN = '" & cmb_Search.text & "' ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"
        End Select
        'Lbl_search.Caption = cmb_Search.Text
        Call Clear_Field
        
        Call Combo_Gubun
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Exit Sub
        
    Case 7  '����
        If (CAR_NO_TMP <> "") Then
            If (MaskEdBox_Fee <> "0") Then
                '��ȭ���� ó���ؾߵ�...!!!
                MBox.Label3.Caption = txt_CarNo.text & vbCrLf & MaskEdBox_Fee.text & "��"
                MBox.Label3.FontSize = 20
                MBox.Label1.Caption = "�� ������ ���������� ����մϴ�." & vbCrLf & vbCrLf & " ����Ͻðڽ��ϱ�?"
                MBox.Label2.Caption = "�������� ���� ���"
                MBox.Show 1
                If (Glo_MsgRet = True) Then
                    'adoConn.Execute "UPDATE tb_reg SET FEE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
                    'adoConn.Execute "INSERT INTO TB_FEE VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.Text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                    
                    qry = "UPDATE tb_reg SET FEE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
                    bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
                    If (bQryResult = False) Then
                        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                        Call DataLogger("[FrmId Update_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                        Exit Sub
                    End If
                    
                    qry = "INSERT INTO TB_FEE VALUES ('" & txt_CarNo & "', '" & txt_CarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.text & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                    bQryResult = DataBaseQueryExec(adoConn, qry, NWERR_GATE_STAY)
                    If (bQryResult = False) Then
                        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                        Call DataLogger("[FrmId Update_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                        Exit Sub
                    End If
        
        
        
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.text & "��    �������� �Ϸ�", 0
                    Call DataLogger("[REG Button]    " & txt_CarNo & "    " & MaskEdBox_Fee.text & "��    �������� �Ϸ�")
                    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    " & MaskEdBox_Fee.Text & "��    �������� �Ϸ�")
                End If
            Else
                MsgBox "�߸��� �ݾ��Դϴ�. Ȯ���ϼ���."
            End If
        Else
            MsgBox "�߸��� ����Դϴ�. Ȯ���ϼ���."
        End If
        Call Clear_Field
        
        Call Combo_Gubun
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Exit Sub
        
    Case 12
        Call AllDeviceSendMsg '�޼��� �߼�(�׽�Ʈ)
        
    Case 13 '�����γ���
        FrmWebdc.Show 1
End Select

On Error Resume Next

End Sub


'�ʼ� �Է� ������ Ȯ��
Private Function Data_Error_Check()
    Dim Error_Flag As Boolean
        
    Error_Flag = True
    
'''    If (LenH(txt_id.text) < 8) Then
'''        Error_Flag = False
'''    End If
'''    If (LenH(txt_password.text) < 8) Then
'''        Error_Flag = False
'''    End If
    
'    If (IsDate(MaskEdBox_Start.Text) = False) Then
'        Error_Flag = False
'    End If
'    If (IsDate(MaskEdBox_End.Text) = False) Then
'        Error_Flag = False
'    End If

    Data_Error_Check = Error_Flag

End Function

Private Sub txt_CarNo_Change()
'    If (LenH(txt_CarNo) > 7 Or LenH(txt_CarNo) = 4) Then
'        Call Search_Record
'    End If
End Sub

Sub Search_Record()
    Dim rs As Recordset
    Dim SQL_SEARCH As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim bQryResult As String
    Dim sPasswordDecode As String

On Error GoTo Err_P

    SQL_SEARCH = "SELECT * From tb_id WHERE ID = '" & txt_id & "' "

    Set rs = New ADODB.Recordset
    'rs.Open SQL_SEARCH, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, SQL_SEARCH, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If

    If (rs.RecordCount <> 0) Then
        cmd_Button(10).Enabled = False
        cmd_Button(8).Enabled = True
        cmd_Button(9).Enabled = True
        chk_Menu(0).value = 0
        chk_Menu(1).value = 0
        chk_Menu(2).value = 0
        chk_Menu(3).value = 0
        chk_Menu(4).value = 0
        chk_Menu(5).value = 0
        chk_Menu(6).value = 0
        chk_Menu(7).value = 0
'        chk_Menu(8).value = 0
'        chk_Menu(9).value = 0
    
        TXT_ID_TMP = rs!ID
        
        'TXT_PASSWORD_TMP = rs!PassWord
        'txt_password.text = rs!PassWord
        sPasswordDecode = DecodeNDE01(rs!PassWord, "www.jawootek.com") '��ȣȭ
        TXT_PASSWORD_TMP = sPasswordDecode
        txt_password = sPasswordDecode

        Combo1.text = "" & rs!Gubun
        If rs!MENU1 = chk_Menu(0).Caption Then
            chk_Menu(0).value = 1
        End If
        If rs!MENU2 = chk_Menu(1).Caption Then
            chk_Menu(1).value = 1
        End If
        If rs!MENU3 = chk_Menu(2).Caption Then
            chk_Menu(2).value = 1
        End If
        If rs!MENU4 = chk_Menu(3).Caption Then
            chk_Menu(3).value = 1
        End If
        If rs!MENU5 = chk_Menu(4).Caption Then
            chk_Menu(4).value = 1
        End If
        If rs!MENU6 = chk_Menu(5).Caption Then
            chk_Menu(5).value = 1
        End If
        If rs!MENU7 = chk_Menu(6).Caption Then
            chk_Menu(6).value = 1
        End If
        If rs!MENU8 = chk_Menu(7).Caption Then
            chk_Menu(7).value = 1
        End If
'        If rs!menu9 = chk_Menu(8).Caption Then
'            chk_Menu(8).value = 1
'        End If
'        If rs!MENU10 = chk_Menu(9).Caption Then
'            chk_Menu(9).value = 1
'        End If

    Else

    End If
    Set rs = Nothing
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call Clear_WebDC
    Call Disable_WebDC
    
    Dim i As Integer
    Dim bChk  As Boolean
    bChk = False
    For i = 0 To 9
        If (chk_Menu(i).Caption = "������" And chk_Menu(i).value = 1) Then
            Call Enable_WebDC
            Exit For
        End If
    Next i
    
    
    SQL_SEARCH = "SELECT * From tb_partner WHERE ID = '" & txt_id & "' "

    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, SQL_SEARCH, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If

    If (rs.RecordCount <> 0) Then
        txt_DC_Code = rs!PCODE
        txt_DC_Partner = rs!PNAME
        
        If (rs!PGUBUN = "T") Then
            cmb_DC_Gubun.text = "�ð�(��)"
        Else
            cmb_DC_Gubun.text = "�ݾ�(��)"
        End If

        txt_DC_Desc(0).text = rs!PDC1_DESC
        If (txt_DC_Desc(0).text = "") Then txt_DC(0).text = "" Else txt_DC(0).text = rs!PDC1
        txt_DC_Desc(1).text = rs!PDC2_DESC
        If (txt_DC_Desc(1).text = "") Then txt_DC(1).text = "" Else txt_DC(1).text = rs!PDC2
        txt_DC_Desc(2).text = rs!PDC3_DESC
        If (txt_DC_Desc(2).text = "") Then txt_DC(2).text = "" Else txt_DC(2).text = rs!PDC3
        txt_DC_Desc(3).text = rs!PDC4_DESC
        If (txt_DC_Desc(3).text = "") Then txt_DC(3).text = "" Else txt_DC(3).text = rs!PDC4
        txt_DC_Desc(4).text = rs!PDC5_DESC
        If (txt_DC_Desc(4).text = "") Then txt_DC(4).text = "" Else txt_DC(4).text = rs!PDC5
        
        lbl_NowFreePoint.Caption = "[" & rs!FREE_POINT & "]"
        lbl_NowPaidPoint.Caption = "[" & rs!PAID_POINT & "]"
    Else
    End If
    
    Exit Sub
    
Err_P:
    Call DataLogger(" [ID Search Record]  " & Err.Description)
End Sub


Private Sub cmd_Search_Click()

If Option1(0).value = True Then
    If Len(txt_tmpCarNo) <> 0 Then
        Select Case cmb_GB.ListIndex
            Case 0
                RegQry = "SELECT * From tb_reg Where CAR_NO Like '%" & txt_tmpCarNo & "'"
            Case 1
                RegQry = "SELECT * From tb_reg Where DRIVER_NAME Like '%" & txt_tmpCarNo & "%'"
            Case 2
                RegQry = "SELECT * From tb_reg Where DRIVER_DEPT Like '%" & txt_tmpCarNo & "%'"
            Case 3
                RegQry = "SELECT * From tb_reg Where DRIVER_CLASS Like '%" & txt_tmpCarNo & "%'"
            Case Else
                RegQry = "SELECT * From tb_reg Where CAR_GUBUN Like '%" & txt_tmpCarNo & "%'"
        End Select
    Else
        Select Case cmb_GB.ListIndex
            Case 0
                RegQry = "SELECT * From tb_reg Order By CAR_NO"
            Case 1
                RegQry = "SELECT * From tb_reg Order By DRIVER_NAME"
            Case 2
                RegQry = "SELECT * From tb_reg Order By DRIVER_DEPT"
            Case 3
                RegQry = "SELECT * From tb_reg Order By DRIVER_CLASS"
            Case Else
                RegQry = "SELECT * From tb_reg Order By CAR_GUBUN"
        End Select
    End If
Else
    If Len(cmbDong.text) = 0 Then
        If Len(cmbHo.text) = 0 Then
            RegQry = "SELECT * From tb_reg"
        Else
            RegQry = "SELECT * From tb_reg Where DRIVER_CLASS = '" & cmbHo.text & "'"
        End If
    Else
        If Len(cmbHo.text) = 0 Then
            RegQry = "SELECT * From tb_reg Where DRIVER_DEPT = '" & cmbDong.text & "'"
        Else
            RegQry = "SELECT * From tb_reg Where DRIVER_DEPT = '" & cmbDong.text & "' AND DRIVER_CLASS = '" & cmbHo.text & "'"
        End If
    End If
End If

txt_tmpCarNo = ""
Call Clear_Field

Call Combo_Gubun
Call ListView_REG_Draw
Call ListView_REG_SQL

End Sub


'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPress(KeyAscii As Integer)

    Dim Car_Num_Str As String
    Dim qry As String
    Dim rs As Recordset
    Dim rs_Part As Recordset
    Dim itmX As ListItem
        
    If (KeyAscii = 13) Then
        If (Len(txt_tmpCarNo) <> 0) Then
            Call cmd_Search_Click
            Exit Sub
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If


End Sub


Private Sub txt_DC_Desc_GotFocus(Index As Integer)
    Dim i As Integer
    
    For i = 0 To WebDC_COUNT - 1
        If (InStr(txt_DC_Desc(i), "���θ�Ī") > 0) Then
            txt_DC_Desc(i).text = ""
        Else
        End If
    Next i
End Sub

Private Sub txt_DC_GotFocus(Index As Integer)
    Dim i As Integer
    
    For i = 0 To WebDC_COUNT - 1
        If (InStr(txt_DC(i), "���ΰ�") > 0) Then
            txt_DC(i).text = ""
        Else
        End If
    Next i

End Sub


Private Sub txt_id_Change()
    'Call Search_Record
    If (LenH(txt_id) > 8) Then
        txt_id.text = LeftH(txt_id, 8)
    End If
End Sub



Private Sub txt_FreeCount_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_FreeCount = "0") Then
        txt_FreeCount = ""
    End If

    If (KeyAscii = 45) Then
        txt_FreeCount = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt_PaidCount_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_PaidCount = "0") Then
        txt_PaidCount = ""
    End If

    If (KeyAscii = 45) Then
        txt_PaidCount = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt_PaidMoney_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_PaidMoney = "0") Then
        txt_PaidMoney = ""
    End If

    If (KeyAscii = 45) Then
        txt_PaidMoney = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub

