VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmG4Mini 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��19455"
   ClientHeight    =   14775
   ClientLeft      =   4620
   ClientTop       =   615
   ClientWidth     =   19395
   BeginProperty Font 
      Name            =   "�������"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "FrmG4Mini.frx":0000
   ScaleHeight     =   985
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   1293
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   " �ڸ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   540
      TabIndex        =   97
      ToolTipText     =   "��� ����(�����,�̵��,���ν�,�������� ����) ���ܱ� ����"
      Top             =   13710
      Width           =   5835
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "�ڸ���� ����1"
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
         Height          =   270
         Index           =   0
         Left            =   270
         TabIndex        =   101
         ToolTipText     =   "[�ڸ����]üũ�� ���:���ν�����, �������������� ������ ������� ������ ������ϴ�."
         Top             =   210
         Width           =   2655
      End
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "�ڸ���� ����2"
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
         Height          =   270
         Index           =   1
         Left            =   3105
         TabIndex        =   100
         ToolTipText     =   "[�ڸ����]üũ�� ���:���ν�����, �������������� ������ ������� ������ ������ϴ�."
         Top             =   210
         Width           =   2655
      End
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "�ڸ���� ����3"
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
         Height          =   270
         Index           =   2
         Left            =   270
         TabIndex        =   99
         ToolTipText     =   "[�ڸ����]üũ�� ���:���ν�����, �������������� ������ ������� ������ ������ϴ�."
         Top             =   450
         Width           =   2655
      End
      Begin VB.CheckBox chk_NoWork 
         BackColor       =   &H00000000&
         Caption         =   "�ڸ���� ����4"
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
         Height          =   270
         Index           =   3
         Left            =   3105
         TabIndex        =   98
         ToolTipText     =   "[�ڸ����]üũ�� ���:���ν�����, �������������� ������ ������� ������ ������ϴ�."
         Top             =   450
         Width           =   2655
      End
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane1"
      Enabled         =   0   'False
      Height          =   555
      Index           =   0
      Left            =   3420
      TabIndex        =   95
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane2"
      Enabled         =   0   'False
      Height          =   555
      Index           =   1
      Left            =   4260
      TabIndex        =   94
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane3"
      Enabled         =   0   'False
      Height          =   555
      Index           =   2
      Left            =   5100
      TabIndex        =   93
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Lane 
      Caption         =   "Lane4"
      Enabled         =   0   'False
      Height          =   555
      Index           =   3
      Left            =   5940
      TabIndex        =   92
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '���
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
      ForeColor       =   &H00E0E0E0&
      Height          =   1605
      Left            =   19650
      TabIndex        =   90
      Top             =   7260
      Visible         =   0   'False
      Width           =   11670
   End
   Begin VB.CheckBox Chk_Zoom 
      BackColor       =   &H00000000&
      Caption         =   " ���� Ȯ��"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   180
      Left            =   17460
      TabIndex        =   89
      Top             =   13410
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   " �湮���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   12840
      TabIndex        =   84
      ToolTipText     =   "�湮����(�̵������) ���ܱ� ����"
      Top             =   13710
      Width           =   5835
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "�Ϲ����� ����2"
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
         Height          =   270
         Index           =   1
         Left            =   3105
         TabIndex        =   88
         Top             =   210
         Width           =   2655
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "�Ϲ����� ����4"
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
         Height          =   270
         Index           =   3
         Left            =   3105
         TabIndex        =   87
         Top             =   450
         Width           =   2655
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "�Ϲ����� ����3"
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
         Height          =   270
         Index           =   2
         Left            =   270
         TabIndex        =   86
         Top             =   450
         Width           =   2655
      End
      Begin VB.CheckBox Chk_FreePass 
         BackColor       =   &H00000000&
         Caption         =   "�Ϲ����� ����1"
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
         Height          =   270
         Index           =   0
         Left            =   270
         TabIndex        =   85
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   " �������� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   6690
      TabIndex        =   79
      ToolTipText     =   "����������(�ù�,ȭ��) ���ܱ� ����"
      Top             =   13710
      Width           =   5835
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "�������� ����4"
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
         Height          =   270
         Index           =   3
         Left            =   3105
         TabIndex        =   83
         Top             =   450
         Width           =   2655
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "�������� ����3"
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
         Height          =   270
         Index           =   2
         Left            =   270
         TabIndex        =   82
         Top             =   450
         Width           =   2655
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "�������� ����2"
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
         Height          =   270
         Index           =   1
         Left            =   3105
         TabIndex        =   81
         Top             =   210
         Width           =   2655
      End
      Begin VB.CheckBox chk_Taxi 
         BackColor       =   &H00000000&
         Caption         =   "�������� ����1"
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
         Height          =   270
         Index           =   0
         Left            =   270
         TabIndex        =   80
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.TextBox txt_CarNo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�������"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   810
      TabIndex        =   62
      Text            =   "25��5401"
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox Chk_FreePass 
      BackColor       =   &H00000000&
      Caption         =   "����6"
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
      Height          =   360
      Index           =   5
      Left            =   21240
      TabIndex        =   61
      Top             =   14040
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CheckBox Chk_FreePass 
      BackColor       =   &H00000000&
      Caption         =   "����5"
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
      Height          =   360
      Index           =   4
      Left            =   21240
      TabIndex        =   60
      Top             =   13680
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE9CE&
      Height          =   4395
      Index           =   3
      Left            =   14415
      TabIndex        =   54
      Top             =   2055
      Width           =   4755
      Begin Threed.SSCommand SSCommand1 
         Height          =   825
         Index           =   3
         Left            =   3840
         TabIndex        =   75
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3555
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "OPEN"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":3A512A
      End
      Begin Threed.SSCommand cmd_GateClose 
         Height          =   870
         Index           =   3
         Left            =   15
         TabIndex        =   105
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3540
         Width           =   930
         _Version        =   65536
         _ExtentX        =   1640
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "�ݱ�"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":3A7184
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   3
         Left            =   540
         Picture         =   "FrmG4Mini.frx":3A7858
         Top             =   720
         Width           =   3690
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '�������� ����
         Height          =   300
         Index           =   3
         Left            =   150
         Top             =   135
         Width           =   300
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   1185
         TabIndex        =   58
         Top             =   150
         Width           =   3405
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�غ���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   3
         Left            =   780
         TabIndex        =   57
         Top             =   3750
         Width           =   3240
      End
      Begin VB.Label lbl_carno 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���00��0000"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   3
         Left            =   2025
         TabIndex        =   56
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label lbl_GN 
         Appearance      =   0  '���
         BackColor       =   &H00800000&
         BackStyle       =   0  '����
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   3
         Left            =   75
         TabIndex        =   55
         Top             =   3165
         Width           =   2175
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '���
         BorderStyle     =   1  '���� ����
         Height          =   3570
         Index           =   3
         Left            =   0
         Picture         =   "FrmG4Mini.frx":3BFED6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE9CE&
      Height          =   4395
      Index           =   2
      Left            =   9615
      TabIndex        =   49
      Top             =   2055
      Width           =   4755
      Begin Threed.SSCommand SSCommand1 
         Height          =   825
         Index           =   2
         Left            =   3825
         TabIndex        =   74
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3555
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "OPEN"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":3E5609
      End
      Begin Threed.SSCommand cmd_GateClose 
         Height          =   870
         Index           =   2
         Left            =   15
         TabIndex        =   104
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3540
         Width           =   930
         _Version        =   65536
         _ExtentX        =   1640
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "�ݱ�"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":3E7663
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   2
         Left            =   540
         Picture         =   "FrmG4Mini.frx":3E7D37
         Top             =   735
         Width           =   3690
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '�������� ����
         Height          =   300
         Index           =   2
         Left            =   135
         Top             =   135
         Width           =   300
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   1215
         TabIndex        =   53
         Top             =   165
         Width           =   3405
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�غ���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   2
         Left            =   750
         TabIndex        =   52
         Top             =   3750
         Width           =   3240
      End
      Begin VB.Label lbl_carno 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���00��0000"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   2
         Left            =   2055
         TabIndex        =   51
         Top             =   3135
         Width           =   2565
      End
      Begin VB.Label lbl_GN 
         Appearance      =   0  '���
         BackColor       =   &H00800000&
         BackStyle       =   0  '����
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   2
         Left            =   75
         TabIndex        =   50
         Top             =   3165
         Width           =   2175
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '���
         BorderStyle     =   1  '���� ����
         Height          =   3570
         Index           =   2
         Left            =   0
         Picture         =   "FrmG4Mini.frx":4003B5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4725
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE9CE&
      Height          =   4395
      Index           =   1
      Left            =   4830
      TabIndex        =   44
      Top             =   2055
      Width           =   4755
      Begin Threed.SSCommand SSCommand1 
         Height          =   825
         Index           =   1
         Left            =   3825
         TabIndex        =   73
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3555
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "OPEN"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":425AE8
      End
      Begin Threed.SSCommand cmd_GateClose 
         Height          =   870
         Index           =   1
         Left            =   -30
         TabIndex        =   103
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3540
         Width           =   930
         _Version        =   65536
         _ExtentX        =   1640
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "�ݱ�"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":427B42
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   1
         Left            =   525
         Picture         =   "FrmG4Mini.frx":428216
         Top             =   750
         Width           =   3690
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '�������� ����
         Height          =   300
         Index           =   1
         Left            =   135
         Top             =   135
         Width           =   300
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�غ���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   1
         Left            =   735
         TabIndex        =   48
         Top             =   3750
         Width           =   3240
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   1215
         TabIndex        =   47
         Top             =   150
         Width           =   3405
      End
      Begin VB.Label lbl_carno 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���00��0000"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   1
         Left            =   2055
         TabIndex        =   46
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label lbl_GN 
         Appearance      =   0  '���
         BackColor       =   &H00800000&
         BackStyle       =   0  '����
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   1
         Left            =   75
         TabIndex        =   45
         Top             =   3165
         Width           =   2175
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '���
         BorderStyle     =   1  '���� ����
         Height          =   3570
         Index           =   1
         Left            =   0
         Picture         =   "FrmG4Mini.frx":440894
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4725
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   11400
      TabIndex        =   38
      Top             =   10380
      Width           =   945
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1125
      Left            =   13065
      TabIndex        =   2
      Top             =   8535
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1984
      View            =   3
      LabelEdit       =   1
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      IMEMode         =   10  '�ѱ� 
      Left            =   14265
      TabIndex        =   0
      Top             =   7365
      Width           =   2460
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   21240
      TabIndex        =   1
      Top             =   13440
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4860
      Top             =   45
   End
   Begin MSWinsockLib.Winsock APS_Winsock 
      Left            =   5700
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Remote_Winsock 
      Left            =   6120
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   2100
      Left            =   375
      TabIndex        =   22
      Top             =   10800
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   0
      BackColor       =   16771534
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE9CE&
      Height          =   4395
      Index           =   0
      Left            =   45
      TabIndex        =   39
      Top             =   2055
      Width           =   4755
      Begin Threed.SSCommand SSCommand1 
         Height          =   825
         Index           =   0
         Left            =   3825
         TabIndex        =   72
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3555
         Width           =   900
         _Version        =   65536
         _ExtentX        =   1587
         _ExtentY        =   1455
         _StockProps     =   78
         Caption         =   "OPEN"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":465FC7
      End
      Begin Threed.SSCommand cmd_GateClose 
         Height          =   870
         Index           =   0
         Left            =   0
         TabIndex        =   102
         ToolTipText     =   "���ܱ⸦ �����մϴ�.."
         Top             =   3540
         Width           =   930
         _Version        =   65536
         _ExtentX        =   1640
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "�ݱ�"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmG4Mini.frx":468021
      End
      Begin VB.Image Imgshutdown 
         Height          =   2025
         Index           =   0
         Left            =   540
         Picture         =   "FrmG4Mini.frx":4686F5
         Top             =   765
         Width           =   3690
      End
      Begin VB.Shape Shp_Rec 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '�������� ����
         Height          =   300
         Index           =   0
         Left            =   120
         Top             =   135
         Width           =   300
      End
      Begin VB.Label lbl_RecState 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�غ���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   21.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   0
         Left            =   765
         TabIndex        =   43
         Top             =   3720
         Width           =   3240
      End
      Begin VB.Label lbl_time_now 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "lbl_time_now"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   0
         Left            =   1200
         TabIndex        =   42
         Top             =   135
         Width           =   3405
      End
      Begin VB.Label lbl_carno 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "���00��0000"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   18
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   390
         Index           =   0
         Left            =   2040
         TabIndex        =   41
         Top             =   3105
         Width           =   2565
      End
      Begin VB.Label lbl_GN 
         Appearance      =   0  '���
         BackColor       =   &H00800000&
         BackStyle       =   0  '����
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   0
         Left            =   75
         TabIndex        =   40
         Top             =   3165
         Width           =   2175
      End
      Begin VB.Image ImageIn 
         Appearance      =   0  '���
         BorderStyle     =   1  '���� ����
         Height          =   3570
         Index           =   0
         Left            =   0
         Picture         =   "FrmG4Mini.frx":480D73
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4725
      End
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�湮������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   20100
      TabIndex        =   107
      Top             =   5625
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lbl_ParkFull 
      BackStyle       =   0  '����
      Caption         =   "������Ȳ : Now/Full"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8910
      TabIndex        =   96
      Top             =   240
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404040&
      BackStyle       =   0  '����
      Caption         =   "[���ܱ� �ڵ�����]"
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
      Height          =   285
      Left            =   570
      TabIndex        =   91
      Top             =   13350
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillStyle       =   0  '�ܻ�
      Height          =   1530
      Left            =   180
      Top             =   13170
      Width           =   19065
   End
   Begin VB.Label LblTime 
      Alignment       =   1  '������ ����
      BackColor       =   &H00000000&
      BackStyle       =   0  '����
      Caption         =   "�ð�"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   15315
      TabIndex        =   78
      Top             =   465
      Width           =   3915
   End
   Begin VB.Label LblDBInfo 
      Alignment       =   1  '������ ����
      BackColor       =   &H00000000&
      BackStyle       =   0  '����
      Caption         =   "DB���� �޽���"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12540
      TabIndex        =   76
      Top             =   30
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   20010
      TabIndex        =   77
      Top             =   4605
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Index           =   2
      Left            =   4020
      Picture         =   "FrmG4Mini.frx":4A64A6
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Index           =   1
      Left            =   3495
      Picture         =   "FrmG4Mini.frx":4A688F
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   5685
      TabIndex        =   71
      Top             =   1425
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label8 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   17370
      TabIndex        =   70
      Top             =   7530
      Width           =   1170
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Index           =   3
      Left            =   4545
      Picture         =   "FrmG4Mini.frx":4A6C78
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Index           =   3
      Left            =   4545
      Picture         =   "FrmG4Mini.frx":4A7061
      Top             =   1440
      Width           =   465
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Index           =   2
      Left            =   4020
      Picture         =   "FrmG4Mini.frx":4A7448
      Top             =   1440
      Width           =   465
   End
   Begin VB.Image ImgGreen 
      Height          =   495
      Index           =   0
      Left            =   2970
      Picture         =   "FrmG4Mini.frx":4A782F
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Index           =   0
      Left            =   2970
      Picture         =   "FrmG4Mini.frx":4A7C18
      Top             =   1455
      Width           =   465
   End
   Begin VB.Image ImgRed 
      Height          =   450
      Index           =   1
      Left            =   3495
      Picture         =   "FrmG4Mini.frx":4A7FFF
      Top             =   1455
      Width           =   465
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "������̷�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   20070
      TabIndex        =   69
      Top             =   3285
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��ȣ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   10875
      TabIndex        =   68
      Top             =   1425
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "��������ȸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9135
      TabIndex        =   67
      Top             =   1425
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����ǰ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   12630
      TabIndex        =   66
      Top             =   1425
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�ٹ��ڰ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   14370
      TabIndex        =   65
      Top             =   1425
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "ȯ�漳��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   16110
      TabIndex        =   64
      Top             =   1425
      Width           =   1050
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�ý�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   17835
      TabIndex        =   63
      Top             =   1425
      Width           =   1050
   End
   Begin VB.Label LblGubun 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   21
      Top             =   12525
      Width           =   3240
   End
   Begin VB.Label Label6 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "�� ��  �� ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   13185
      TabIndex        =   20
      Top             =   12525
      Width           =   1560
   End
   Begin VB.Label LblName 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   14
      Top             =   10230
      Width           =   3240
   End
   Begin VB.Label Lbl1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "�� ��  �� ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   13185
      TabIndex        =   13
      Top             =   9780
      Width           =   1560
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��        ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   13185
      TabIndex        =   12
      Top             =   10230
      Width           =   1590
   End
   Begin VB.Label Label3 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��        ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   13185
      TabIndex        =   11
      Top             =   10710
      Width           =   1590
   End
   Begin VB.Label Label4 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��   ��   ó "
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   13185
      TabIndex        =   10
      Top             =   11175
      Width           =   1710
   End
   Begin VB.Label Label5 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   " �� / ȣ ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   13170
      TabIndex        =   9
      Top             =   11610
      Width           =   1440
   End
   Begin VB.Label Label6 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��        ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   13185
      TabIndex        =   8
      Top             =   12060
      Width           =   1590
   End
   Begin VB.Label LblCar 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   7
      Top             =   9795
      Width           =   3240
   End
   Begin VB.Label LblId 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   6
      Top             =   10725
      Width           =   3240
   End
   Begin VB.Label LblCarType 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   5
      Top             =   11175
      Width           =   3240
   End
   Begin VB.Label LblTel 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   4
      Top             =   11610
      Width           =   3240
   End
   Begin VB.Label LblDate 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   15660
      TabIndex        =   3
      Top             =   12060
      Width           =   3240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "ī�޶� ����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1395
      TabIndex        =   59
      Top             =   1515
      Width           =   1590
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�������� �ý���"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   18
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   1395
      TabIndex        =   37
      Top             =   975
      Width           =   3120
   End
   Begin VB.Label LblSearch 
      BackColor       =   &H00404040&
      Caption         =   "�˻���� : "
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   13065
      TabIndex        =   15
      Top             =   8190
      Width           =   4065
   End
   Begin VB.Label lbl_info_out 
      BackStyle       =   0  '����
      Caption         =   "lbl_info_out"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   15240
      TabIndex        =   19
      Top             =   17640
      Width           =   3630
   End
   Begin VB.Label lbl_title_out 
      BackStyle       =   0  '����
      Caption         =   "lbl_title_out"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   13275
      TabIndex        =   18
      Top             =   17640
      Width           =   1815
   End
   Begin VB.Label lbl_title_in 
      BackStyle       =   0  '����
      Caption         =   "lbl_title_in"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   285
      TabIndex        =   17
      Top             =   17640
      Width           =   1800
   End
   Begin VB.Label lbl_info_in 
      BackStyle       =   0  '����
      Caption         =   "lbl_info_in"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   2265
      TabIndex        =   16
      Top             =   17640
      Width           =   3615
   End
   Begin VB.Label LblGubun 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   36
      Top             =   9750
      Width           =   2040
   End
   Begin VB.Label Label6 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "ó �� �� ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   7065
      TabIndex        =   35
      Top             =   9735
      Width           =   1815
   End
   Begin VB.Label LblName 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   34
      Top             =   7815
      Width           =   2040
   End
   Begin VB.Label Lbl1 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "�� �� �� ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   7065
      TabIndex        =   33
      Top             =   7380
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��       ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   7065
      TabIndex        =   32
      Top             =   7800
      Width           =   1830
   End
   Begin VB.Label Label3 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��       ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   6
      Left            =   7065
      TabIndex        =   31
      Top             =   8190
      Width           =   1830
   End
   Begin VB.Label Label4 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��   ��  ó "
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   7065
      TabIndex        =   30
      Top             =   8580
      Width           =   1830
   End
   Begin VB.Label Label5 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��   ��  ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   7050
      TabIndex        =   29
      Top             =   8970
      Width           =   1845
   End
   Begin VB.Label Label6 
      Appearance      =   0  '���
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "ó �� �� ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   7065
      TabIndex        =   28
      Top             =   9345
      Width           =   1815
   End
   Begin VB.Label LblCar 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   27
      Top             =   7365
      Width           =   2040
   End
   Begin VB.Label LblId 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   26
      Top             =   8220
      Width           =   2040
   End
   Begin VB.Label LblCarType 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   25
      Top             =   8595
      Width           =   2040
   End
   Begin VB.Label LblTel 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   24
      Top             =   8985
      Width           =   2040
   End
   Begin VB.Label LblDate 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   9060
      TabIndex        =   23
      Top             =   9360
      Width           =   2040
   End
   Begin VB.Image ImageLog 
      Appearance      =   0  '���
      BorderStyle     =   1  '���� ����
      Height          =   3570
      Left            =   375
      Picture         =   "FrmG4Mini.frx":4A83E6
      Stretch         =   -1  'True
      Top             =   7185
      Width           =   4725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   6
      Left            =   17460
      Picture         =   "FrmG4Mini.frx":4B57B3
      Top             =   975
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   5
      Left            =   15735
      Picture         =   "FrmG4Mini.frx":4BAAE1
      Top             =   975
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   4
      Left            =   13995
      Picture         =   "FrmG4Mini.frx":4BFE0F
      Top             =   975
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   3
      Left            =   19695
      Picture         =   "FrmG4Mini.frx":4C513D
      Top             =   2835
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   2
      Left            =   12255
      Picture         =   "FrmG4Mini.frx":4CA46B
      Top             =   975
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   1
      Left            =   10515
      Picture         =   "FrmG4Mini.frx":4CF799
      Top             =   975
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   0
      Left            =   8775
      Picture         =   "FrmG4Mini.frx":4D4AC7
      Top             =   975
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   7
      Left            =   5325
      Picture         =   "FrmG4Mini.frx":4D9DF5
      Top             =   975
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   8
      Left            =   19680
      Picture         =   "FrmG4Mini.frx":4DF123
      Top             =   4155
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   14970
      Picture         =   "FrmG4Mini.frx":4E4451
      Top             =   210
      Width           =   4395
   End
   Begin VB.Label Lblbutton 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�湮����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   7410
      TabIndex        =   106
      Top             =   1425
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   10
      Left            =   7050
      Picture         =   "FrmG4Mini.frx":4E47E7
      Top             =   975
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Imgbutton 
      Height          =   915
      Index           =   9
      Left            =   19710
      Picture         =   "FrmG4Mini.frx":4E9B15
      Top             =   5160
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "FrmG4Mini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyText(1 To 8) As New clsText
Dim DataField_Enabled As Boolean
Dim Save_TagNum, APS_CMD As String
Dim FrmImg_F As Boolean


Private Sub cmd_Menu_Click(Index As Integer)

End Sub


Private Sub chk_NoWork_Click(Index As Integer)
    Dim sNoWork As String
    Dim sSendValue As String
    Dim sLaneName, sGuestUse, sAutoMode As String

    If (chk_NoWork(Index).value = 1) Then
        sNoWork = "�ڸ����"
        sSendValue = "Y"
        chk_Taxi(Index).Enabled = False
        Chk_FreePass(Index).Enabled = False
        chk_NoWork(Index).ForeColor = &HFF&
    Else
        sNoWork = "�ٹ���"
        sSendValue = "N"
        chk_Taxi(Index).Enabled = True
        Chk_FreePass(Index).Enabled = True
        chk_NoWork(Index).ForeColor = &HFFFFFF
    End If
    
    Select Case Index
        Case 0
            Glo_Lane1_NoWork = sNoWork
            sLaneName = LANE1_Name
        Case 1
            Glo_Lane2_NoWork = sNoWork
            sLaneName = LANE2_Name
        Case 2
            Glo_Lane3_NoWork = sNoWork
            sLaneName = LANE3_Name
        Case 3
            Glo_Lane4_NoWork = sNoWork
            sLaneName = LANE4_Name
        Case 4
            Glo_Lane5_NoWork = sNoWork
            sLaneName = LANE5_Name
        Case 5
            Glo_Lane6_NoWork = sNoWork
            sLaneName = LANE6_Name
    End Select
    
    
    
    '�湮�� �ڵ� ó������
    If (chk_NoWork(Index).value = 1 Or Chk_FreePass(Index).value = 1) Then
        sGuestUse = "(�ڵ�ó��)"
        sAutoMode = "Y"
    Else
    End If
    If (Not Glo_FrmGuest(Index) Is Nothing) Then '������� �ִٸ�
        Call Glo_FrmGuest(Index).SetAutoMode(sAutoMode, sLaneName & sGuestUse)
        
    End If
    
    
    
    If (Glo_FreepassS_YN = "Y") Then
        FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_NOWORK_" & sSendValue)
        DataLogger ("FreePass Send : " & Index & "_NOWORK_" & sSendValue)
    End If
    
    
    Dim sLog As String
    sLog = "���ܱ� �ڵ�����[�ڸ����] Lane:" & Index + 1 & ":" & sNoWork
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('���ܱ��ڵ�����', 'HOST','" & sLog & "','�ڸ����'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
End Sub


Private Sub cmd_GateClose_Click(Index As Integer)
    On Error GoTo Err_P
    
    Call DataLogger("[GATE DOWN BTN]  Target Gate = Lane" & Index + 1)
    Call Relay_Close(0, Index)
    
    Exit Sub
    
Err_P:
    Call DataLogger("[GATE DOWN BTN]  Target Gate = Lane" & Index + 1 & ", " & Err.Description)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim SQL As String
    Dim Reg_Addr As String
    Dim sLaneName, sGuestUse, sAutoMode As String

    IniFileName$ = App.Path & "\Winpark.ini"
    Report_Path_Name$ = App.Path & "\Data\"
    Doc_Path_Name$ = App.Path & "\Doc\"

    If App.PrevInstance = True Then
        End
    End If
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = 0
    
    
    
    If (Glo_ParkFull_YN = "Y") Then
        Call ParkFull_Show
    End If
    
    
    
    If (Glo_TestMode = "Y") Then
        txt_CarNo.Enabled = True
        Lane(0).Enabled = True
        Lane(1).Enabled = True
        Lane(2).Enabled = True
        Lane(3).Enabled = True
        txt_CarNo.Visible = True
        Lane(0).Visible = True
        Lane(1).Visible = True
        Lane(2).Visible = True
        Lane(3).Visible = True
    Else
        txt_CarNo.Enabled = False
        Lane(0).Enabled = False
        Lane(1).Enabled = False
        Lane(2).Enabled = False
        Lane(3).Enabled = False
        txt_CarNo.Visible = False
        Lane(0).Visible = False
        Lane(1).Visible = False
        Lane(2).Visible = False
        Lane(3).Visible = False
    End If
    
    
    Call ListView_Init1
    Call ListView_Init2
    
    For i = 0 To 3
        ImageIn(i).Picture = LoadPicture(App.Path & "\NoCar.jpg")
        lbl_GN(0).Caption = ""
        lbl_carno(i).Caption = ""
        lbl_time_now(i).Caption = Format(Now, "YYYY-MM-DD HH:NN:SS")
        Shp_Rec(i).Visible = False
        
        Chk_FreePass(i).Caption = ""
    Next i


    If (Glo_User_Type = "����1/����2") Then
        Label5(0).Caption = "�� ��, �� ��"
    Else
        Label5(0).Caption = "  �� / ȣ ��"
    End If
    
    LblCar(0).Caption = ""
    LblName(0).Caption = ""
    LblId(0).Caption = ""
    LblCarType(0).Caption = ""
    LblTel(0).Caption = ""
    LblDate(0).Caption = ""
    LblGubun(0).Caption = ""
    
    LblCar(1).Caption = ""
    LblName(1).Caption = ""
    LblId(1).Caption = ""
    LblCarType(1).Caption = ""
    LblTel(1).Caption = ""
    LblDate(1).Caption = ""
    LblGubun(1).Caption = ""
    
    Text1.text = ""
    

    lbl_GN(0).Caption = Trim(LANE1_Name)
    lbl_GN(1).Caption = Trim(LANE2_Name)
    lbl_GN(2).Caption = Trim(LANE3_Name)
    lbl_GN(3).Caption = Trim(LANE4_Name)

    
    ' ���������� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ����
    Call Chk_TaxiPassEnable(Me, LANE1_YN, Glo_TAXI1_YN, 0, LANE1_Name)
    Call Chk_TaxiPassEnable(Me, LANE2_YN, Glo_TAXI2_YN, 1, LANE2_Name)
    Call Chk_TaxiPassEnable(Me, LANE3_YN, Glo_TAXI3_YN, 2, LANE3_Name)
    Call Chk_TaxiPassEnable(Me, LANE4_YN, Glo_TAXI4_YN, 3, LANE4_Name)
    ' ���������� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ��
    
    ' �Ϲ����� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ����
    Call Chk_NormalPassEnable(Me, LANE1_YN, Glo_FreePassLane1_YN, 0, LANE1_Name)
    Call Chk_NormalPassEnable(Me, LANE2_YN, Glo_FreePassLane2_YN, 1, LANE2_Name)
    Call Chk_NormalPassEnable(Me, LANE3_YN, Glo_FreePassLane3_YN, 2, LANE3_Name)
    Call Chk_NormalPassEnable(Me, LANE4_YN, Glo_FreePassLane4_YN, 3, LANE4_Name)
    ' �Ϲ����� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ��
    
    chk_NoWork(0).Caption = LANE1_Name
    chk_NoWork(1).Caption = LANE2_Name
    chk_NoWork(2).Caption = LANE3_Name
    chk_NoWork(3).Caption = LANE4_Name
    
    
    If (Glo_Screen_No = 4) Then
    
        '�湮���� �Է� �� ����
        For i = 0 To MAX_LANE_COUNT
            If (Not Glo_FrmGuest(i) Is Nothing) Then '������� �ִٸ�
                Unload Glo_FrmGuest(i)
                Set Glo_FrmGuest(i) = Nothing
            End If
        Next i
        
        If (LANE1_YN = "Y" And Glo_GUEST_LANE1_YN = "Y") Then
            Set Glo_FrmGuest(0) = New FormGuest1
            Glo_FrmGuest(0).Show 0
            Call Glo_FrmGuest(0).SetGateNo(0, Glo_Guest_Print_Model(0), Glo_Guest_Print_Port(0))
            
            If (Glo_FreePassLane1_YN = "Y") Then
                sGuestUse = "(�ڵ�ó��)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            'Call Glo_FrmGuest(0).SetGuestName(LANE1_Name & sGuestUse)
            Call Glo_FrmGuest(0).SetAutoMode(sAutoMode, LANE1_Name & sGuestUse)
        End If
        
        If (LANE2_YN = "Y" And Glo_GUEST_LANE2_YN = "Y") Then
            Set Glo_FrmGuest(1) = New FormGuest1
            Glo_FrmGuest(1).Show 0
            Call Glo_FrmGuest(1).SetGateNo(1, Glo_Guest_Print_Model(1), Glo_Guest_Print_Port(1))
            
            If (Glo_FreePassLane2_YN = "Y") Then
                sGuestUse = "(�ڵ�ó��)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            'Call Glo_FrmGuest(1).SetGuestName(LANE2_Name & sGuestUse)
            Call Glo_FrmGuest(1).SetAutoMode(sAutoMode, LANE2_Name & sGuestUse)
        End If
        
        If (LANE3_YN = "Y" And Glo_GUEST_LANE3_YN = "Y") Then
            Set Glo_FrmGuest(2) = New FormGuest1
            Glo_FrmGuest(2).Show 0
            Call Glo_FrmGuest(2).SetGateNo(2, Glo_Guest_Print_Model(2), Glo_Guest_Print_Port(2))
            
            If (Glo_FreePassLane3_YN = "Y") Then
                sGuestUse = "(�ڵ�ó��)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            'Call Glo_FrmGuest(2).SetGuestName(LANE3_Name & sGuestUse)
            Call Glo_FrmGuest(2).SetAutoMode(sAutoMode, LANE3_Name & sGuestUse)
        End If
        If (LANE4_YN = "Y" And Glo_GUEST_LANE4_YN = "Y") Then
            Set Glo_FrmGuest(3) = New FormGuest1
            Glo_FrmGuest(3).Show 0
            Call Glo_FrmGuest(3).SetGateNo(3, Glo_Guest_Print_Model(3), Glo_Guest_Print_Port(3))
            
            If (Glo_FreePassLane4_YN = "Y") Then
                sGuestUse = "(�ڵ�ó��)"
                sAutoMode = "Y"
            Else
                sGuestUse = ""
                sAutoMode = "N"
            End If
            'Call Glo_FrmGuest(3).SetGuestName(LANE4_Name & sGuestUse)
            Call Glo_FrmGuest(3).SetAutoMode(sAutoMode, LANE4_Name & sGuestUse)
        End If

    End If

    
'''    If (Glo_Login_ID = "") Then
'''
'''        For i = 0 To 8
'''            Lblbutton(i).Enabled = False
'''            Imgbutton(i).Enabled = False
'''        Next i
'''
'''        Lblbutton(1).Enabled = True
'''        Imgbutton(1).Enabled = True
'''        Lblbutton(6).Enabled = True
'''        Imgbutton(6).Enabled = True
'''
'''    Else
'''        Call frmLogin.ShowMenu(Glo_Login_ID, Glo_Login_PW)
'''    End If
    
    Call ProtectMainMenuButton(Me)
    
    Call ShowTitlebarSiteCode
    
    
    Timer1.Enabled = True
    FrmTcpServer.Hide
    gHW = Me.hwnd
    Call Hook
    
End Sub

Private Sub sOutput(strText As String, strIP As String)
    List1.AddItem " " & Format(Now, "yyyy-mm-dd hh:nn:ss") & strText & "     " & strIP, 0
End Sub

Public Sub ReDraw(sKind As String, iIndex As Integer, iValue As Integer)
    If sKind = "FreePass" Then
        Chk_FreePass(iIndex).value = iValue
        Call Chk_FreePass_Click(iIndex)
    ElseIf sKind = "Taxi" Then
        chk_Taxi(iIndex).value = iValue
        Call chk_Taxi_Click(iIndex)
    ElseIf sKind = "NOWORK" Then
        chk_NoWork(iIndex).value = iValue
        Call chk_NoWork_Click(iIndex)
    End If
End Sub

Private Sub Chk_FreePass_Click(Index As Integer)

    Dim sGuestUse, sLaneName, sAutoMode As String
    
    Select Case Index
        
        Case 0
                If Chk_FreePass(0).value = 1 Then
                    Glo_FreePassLane1_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane1_YN", "Y")
                Else
                    Glo_FreePassLane1_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane1_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane1_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane1_YN)
                End If

                sLaneName = LANE1_Name
        
        Case 1
                If Chk_FreePass(1).value = 1 Then
                    Glo_FreePassLane2_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane2_YN", "Y")
                Else
                    Glo_FreePassLane2_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane2_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane2_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane2_YN)
                End If
                
                sLaneName = LANE2_Name
                
        Case 2
                If Chk_FreePass(2).value = 1 Then
                    Glo_FreePassLane3_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane3_YN", "Y")
                Else
                    Glo_FreePassLane3_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane3_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane3_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane3_YN)
                End If
                
                sLaneName = LANE3_Name
        Case 3
                If Chk_FreePass(3).value = 1 Then
                    Glo_FreePassLane4_YN = "Y"
                    Call Put_Ini("System Config", "FreePassLane4_YN", "Y")
                Else
                    Glo_FreePassLane4_YN = "N"
                    Call Put_Ini("System Config", "FreePassLane4_YN", "N")
                End If
                
                If (Glo_FreepassS_YN = "Y") Then
                    FrmTcpServer.FreepassS_sock.SendData (Index & "_FREEPASS_" & Glo_FreePassLane4_YN)
                    DataLogger ("FreePass Send : " & Index & "_FREEPASS_" & Glo_FreePassLane4_YN)
                End If
                
                sLaneName = LANE4_Name
    End Select
    
    
    '�湮�� �ڵ� ó������
    If (Chk_FreePass(Index).value = 1) Then
        sGuestUse = "(�ڵ�ó��)"
        sAutoMode = "Y"
    Else
        sGuestUse = ""
        sAutoMode = "N"
    End If
    If (Not Glo_FrmGuest(Index) Is Nothing) Then '������� �ִٸ�
        'Call Glo_FrmGuest(Index).SetGuestName(sLaneName & sGuestUse)
        Call Glo_FrmGuest(Index).SetAutoMode(sAutoMode, sLaneName & sGuestUse)
    End If

    
    Dim sLog As String
    If (sAutoMode = "Y") Then
        sLog = "Lane" & Index + 1 & ":" & "�湮�����ڵ�����"
    Else
        sLog = "Lane" & Index + 1 & ":" & "�湮�����ڵ����� ����"
    End If
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('���ܱ��ڵ�����', 'HOST','" & sLog & "','�湮����'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"


End Sub

Public Sub Chk_FreePassEnable(Index As Integer, bVal As Boolean)
    If (Index < Glo_Screen_No) Then
        Chk_FreePass(Index).Enabled = bVal
    End If
End Sub

Private Sub chk_Taxi_Click(Index As Integer)
    Select Case Index
        Case 0
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI1_YN = "Y"
            Else
                Glo_TAXI1_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI1_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI1_YN)
            End If
        
        Case 1
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI2_YN = "Y"
            Else
                Glo_TAXI2_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI2_YN", Glo_TAXI2_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI2_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI2_YN)
            End If
        Case 2
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI3_YN = "Y"
            Else
                Glo_TAXI3_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI3_YN", Glo_TAXI3_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI3_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI3_YN)
            End If
        Case 3
            If chk_Taxi(Index).value = 1 Then
                Glo_TAXI4_YN = "Y"
            Else
                Glo_TAXI4_YN = "N"
            End If
            Call Put_Ini("System Config", "TAXI4_YN", Glo_TAXI4_YN)
            
            If (Glo_FreepassS_YN = "Y") Then
                FrmTcpServer.FreepassS_sock.SendData (CStr(Index) & "_TAXI_" & Glo_TAXI4_YN)
                DataLogger ("Taxi Send : " & Index & "_TAXI_" & Glo_TAXI4_YN)
            End If
    End Select
    
    '�������� ���ܱ� �ڵ�����
    Dim sLog As String
    If (chk_Taxi(Index).value = 1) Then
        sLog = "Lane" & Index + 1 & ":" & "���������ڵ�����"
    Else
        sLog = "Lane" & Index + 1 & ":" & "���������ڵ����� ����"
    End If
    Call DataLogger(sLog)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('���ܱ��ڵ�����', 'HOST','" & sLog & "','��������'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"

End Sub


Public Sub Set_GateName(ByVal iIndex As Integer, ByVal sGateName As String)
    If (iIndex < Glo_Screen_No) Then
        lbl_GN(iIndex).Caption = sGateName
        Chk_FreePass(iIndex).Caption = sGateName
    End If
End Sub

Private Sub Command1_Click()
    Call ListView_Init2
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.FontBold = False

If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If

End Sub



Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If

End Sub

Private Sub ImageIn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Chk_Zoom.value = 0) Then
        Exit Sub
    End If
    
    FrmImg.Image1.Picture = ImageIn(Index).Picture
    FrmImg.Show 0
    FrmImg_F = True

End Sub
Private Sub ImageLog_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (FrmImg_F) Then
        FrmImg_F = False
        FrmImg.Hide
    End If

End Sub

Private Sub Imgbutton_Click(Index As Integer)
    Call SelectMenuButton(Me, Index)
    Exit Sub
    
'Dim i As Integer
'
'Call GuestForm_WindowState(vbMinimized)
'
'Me.MousePointer = 11
'Select Case Index
'    Case 0
'         FrmInOut.Show 1
'         Me.MousePointer = 0
'         Call DataLogger("[HOST Button]    " & "������ ���� ȭ�� ����")
'    Case 2
'         FrmReg.Show 1
'         Me.MousePointer = 0
'         Call DataLogger("[HOST Button]    " & "����ǰ��� ȭ�� ����")
'    Case 5
'         If (Glo_Login_GUBUN = "�Ѱ�������") Then
'            Chk_Zoom.value = 0
'            FrmTcpServer.Show 0
'            Me.MousePointer = 0
'            Call DataLogger("[HOST Button]    " & "TCP Server ȭ�� ����")
'         'ElseIf (Glo_Login_GUBUN = "������") Then
'         Else
'            Chk_Zoom.value = 0
'            FrmTcpServer2.Show 0
'            Me.MousePointer = 0
'            Call DataLogger("[HOST Button]    " & "TCP Server2 ȭ�� ����")
'        End If
'    Case 6
'         Call DataLogger("[HOST Button]    " & "�������� �ý��� ����!!!")
'         Unload Me
'    Case 1
''''         If (Lblbutton(1).Caption = "��ȣ���") Then
''''            Call DataLogger("[HOST Button]    " & "���α׷� ��ȣ���� ��ȯ")
''''            Lblbutton(1).Caption = "��ȣ����"
''''            For i = 0 To 8
''''                Lblbutton(i).Enabled = False
''''                Imgbutton(i).Enabled = False
''''            Next i
''''            Lblbutton(6).Enabled = True
''''            Lblbutton(1).Enabled = True
''''            Imgbutton(6).Enabled = True
''''            Imgbutton(1).Enabled = True
''''
''''            Lblbutton(7).Visible = False
''''            Imgbutton(7).Visible = False
''''
''''            Put_Ini "System Config", "��ȣ���", "True"
''''
''''            Glo_Login_ID = ""
''''            Glo_Login_PW = ""
''''            Glo_Login_GUBUN = ""
''''         Else
''''            frmLogin.Show 1
''''            Call DataLogger("[HOST Button]    " & "���α׷� ��ȣ��� ����")
''''            Lblbutton(1).Caption = "��ȣ���"
''''            ListView1.SetFocus
''''         End If
''''         Me.MousePointer = 0
'
'         If (Lblbutton(1).Caption = "��ȣ���") Then
'            'Call UnloadForms(Me) '��� �� ����(Jung, FrmTcpServer ���� ����)
'            Call DataLogger("[HOST Button]    " & "���α׷� ��ȣ���� ��ȯ")
'            Call ProtectMainMenuButton(Me)
'
'            Glo_Login_ID = ""
'            Glo_Login_PW = ""
'            Glo_Login_GUBUN = ""
'            Put_Ini "System Config", "��ȣ���", "True"
'
'         Else
'            Call DataLogger("[HOST Button]    " & "���α׷� ��ȣ��� ����")
'            frmLogin.Show 1
'            'Lblbutton(1).Caption = "��ȣ���"
'            ListView1.SetFocus
'         End If
'         Me.MousePointer = 0
'    Case 3
'         FrmRegHistory.Show 1
'         Me.MousePointer = 0
'         Call DataLogger("[HOST Button]    " & "����� �̷� ȭ�� ����")
'    Case 4
'            FrmId.Show 1
'            Me.MousePointer = 0
'            Call DataLogger("[HOST Button]    " & "���̵� ���� ȭ�� ����")
'    Case 7
'        Me.MousePointer = 0
'        If (Lblbutton(Index).Caption = "���������") Then
'            Call DataLogger("[HOST Button]    " & "��������� ���� ȭ�� ����")
'            FrmAccnt.Show 0
'        ElseIf (Lblbutton(Index).Caption = "��������") Then
'            Call DataLogger("[HOST Button]    " & "�������� ȭ�� ����")
'            frmResult.Show 0
'        End If
'
'    Case 8
'        Me.MousePointer = 0
'        frmResult.Show 0
'        Call DataLogger("[HOST Button]    " & "�������� ȭ�� ����")
'
'    Case 9
'        Me.MousePointer = 1
'        FrmGuestLog.Show 1
'        Call DataLogger("[HOST Button]    " & "�湮������ ȭ�� ����")
'
'    Case 10  '�湮���� �����湮
'        Me.MousePointer = 0
'        FrmGuestRegLog.Show 0
'        Call DataLogger("[HOST Button]    " & "�湮���� ȭ�� ����")
'        Exit Sub
'End Select

End Sub

Private Sub Label8_Click()
    LblCar(0).Caption = ""
    LblName(0).Caption = ""
    LblId(0).Caption = ""
    LblCarType(0).Caption = ""
    LblTel(0).Caption = ""
    LblDate(0).Caption = ""
    LblGubun(0).Caption = ""
    LblSearch = ""
    ListView1.ListItems.Clear
    Text1 = ""
    Text1.SetFocus

End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label8.FontBold = True

End Sub

Private Sub Lane_Click(Index As Integer)
    
    Dim CarnoEnc As String
    Dim carno As String
    
    carno = CStr(Index) & "_" & txt_CarNo & "_\\localhost\Lane1\image\20161229\20161229171357960_25��5401.jpg"
    CarnoEnc = EncodeNDE01(carno, "www.jawootek.com")
    carno = DecodeNDE01(CarnoEnc, "www.jawootek.com")

    If (Glo_TestMode = "Y") Then
        Call UDP_Proc(carno)
    End If
End Sub

Private Sub Lblbutton_Click(Index As Integer)
    Call GuestForm_WindowState(vbMinimized)
    Call Imgbutton_Click(Index)
End Sub

Private Sub Lblbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

For i = 0 To 8

    Lblbutton(i).FontBold = False

Next i
Lblbutton(Index).FontBold = True

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    ListView1.SetFocus
    LblCar(0).Caption = ""
    LblName(0).Caption = ""
    LblId(0).Caption = ""
    LblCarType(0).Caption = ""
    LblTel(0).Caption = ""
    LblDate(0).Caption = ""
    LblGubun(0).Caption = ""
    LblCar(0).Caption = ListView1.SelectedItem.text
    LblName(0).Caption = ListView1.SelectedItem.SubItems(1)
    LblId(0).Caption = ListView1.SelectedItem.SubItems(2)
    LblCarType(0).Caption = ListView1.SelectedItem.SubItems(3)
    LblTel(0).Caption = ListView1.SelectedItem.SubItems(4)
    If (ListView1.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView1.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
        LblDate(0).ForeColor = vbBlack
        LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6)
    Else
        LblDate(0).ForeColor = vbRed
        LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6) & "   " & "[�Ⱓ ����]"
    End If
    LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)
End Sub


Private Sub ListView2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (FrmImg_F) Then
    FrmImg_F = False
    FrmImg.Hide
End If
End Sub

Private Sub SSCommand1_Click(Index As Integer)
On Error GoTo Err_Proc
        
'        Select Case Index
'            Case 0
'                Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane1")
'                Call Relay_Out(0, 0)
'            Case 1
'                Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane2")
'                Call Relay_Out(0, 1)
'            Case 2
'                Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane3")
'                Call Relay_Out(0, 2)
'            Case 3
'                Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane4")
'                Call Relay_Out(0, 3)
'        End Select

'    Debug.Print "[GATE OPEN BTN]  Target Gate = Lane" & Index + 1

'    If (Glo_ParkFull_YN = "Y") Then
'        If (Glo_ParkFull_Count > Glo_ParkNow_Count) Then
'            Glo_ParkNow_Count = Glo_ParkNow_Count + 1
'
'            Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1)
'            Call Relay_Out(0, Index)
'
'        Else
'            Glo_ParkNow_Count = Glo_ParkFull_Count
'            Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1 & "����:���ܱ� �ȿ���")
'        End If
'    Else
'        Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1)
'        Call Relay_Out(0, Index)
'    End If
    
    Call DataLogger("[GATE OPEN BTN]  Target Gate = Lane" & Index + 1)
    Call Relay_Out(0, Index)

    If (Glo_ParkFull_YN = "Y") Then
        Dim sInOut As String
        Select Case Index
            Case 0
                sInOut = LANE1_Inout
            Case 1
                sInOut = LANE2_Inout
            Case 2
                sInOut = LANE3_Inout
            Case 3
                sInOut = LANE4_Inout
            Case 4
                sInOut = LANE5_Inout
            Case 5
                sInOut = LANE6_Inout
        End Select
        
        Call ParkFull_GetState(Index, sInOut)   '�������
        Call ParkFull_PutNMLDisplay(Index)  '���������
        Call ParkFull_Show                  'ȭ�����
    End If
    
    Exit Sub
        
Err_Proc:
    Call DataLogger(" [cmd_GateOpen_Click]  " & Err.Description)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Car_Num_Str As String
    Dim qry As String
    Dim rs As Recordset
    Dim rs_Part As Recordset
    Dim itmX As ListItem
    Dim bQryResult As Boolean
    
    If (KeyAscii = 13) Then
        LblCar(0).Caption = ""
        LblName(0).Caption = ""
        LblId(0).Caption = ""
        LblCarType(0).Caption = ""
        LblTel(0).Caption = ""
        LblDate(0).Caption = ""
        LblGubun(0).Caption = ""
        If ((Len(Text1) <> 4) Or Not (IsNumeric(Text1))) Then
            MsgBox "������ȣ ���� ���ڸ��� ��Ȯ�ϰ� �Է��ϼ���!"
            Text1 = ""
            Exit Sub
        End If
        qry = "Select * From tb_reg Where CAR_NO LIKE CONCAT( '%', '" & Text1 & "','%')"
        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
            Exit Sub
        End If
        
        
        ListView1.ListItems.Clear
        If (rs.EOF) Then
            LblSearch.Caption = "�˻���� : �ڷᰡ ���� �����ʽ��ϴ�.."
        Else
            LblSearch.Caption = "�˻���� : " & (rs.RecordCount) & " ��"
            Do While Not (rs.EOF)
                Set itmX = ListView1.ListItems.Add(, , "" & rs!CAR_NO)
                itmX.SubItems(1) = "" & rs!DRIVER_NAME
                itmX.SubItems(2) = "" & rs!CAR_GUBUN
                itmX.SubItems(3) = "" & rs!DRIVER_PHONE
                'itmX.SubItems(4) = "" & rs!CAR_MODEL
                itmX.SubItems(4) = "" & rs!DRIVER_DEPT & " / " & rs!DRIVER_CLASS
                itmX.SubItems(5) = "" & Left(rs!START_DATE, 10)
                itmX.SubItems(6) = "" & Left(rs!END_DATE, 10)
                itmX.SubItems(7) = "" & Format(rs!REG_DATE, "YYYY-MM-DD HH:NN:SS")
                rs.MoveNext
            Loop
            ListView1.ListItems.Item(1).Selected = True
            If (rs.RecordCount = 1) Then
            Else
                ListView1.SetFocus
            End If
            LblCar(0).Caption = ListView1.SelectedItem.text
            LblName(0).Caption = ListView1.SelectedItem.SubItems(1)
            LblId(0).Caption = ListView1.SelectedItem.SubItems(2)
            LblCarType(0).Caption = ListView1.SelectedItem.SubItems(3)
            LblTel(0).Caption = ListView1.SelectedItem.SubItems(4)
            If (ListView1.SelectedItem.SubItems(5) <= Format(Now, "yyyymmdd") And ListView1.SelectedItem.SubItems(6) >= Format(Now, "yyyymmdd")) Then
                'LblDate(0).ForeColor = vbWhite
                LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6)
            Else
                LblDate(0).ForeColor = vbRed
                LblDate(0).Caption = ListView1.SelectedItem.SubItems(5) & " - " & ListView1.SelectedItem.SubItems(6) & "   " & "[�Ⱓ����]"
            End If
            LblGubun(0).Caption = ListView1.SelectedItem.SubItems(7)
        End If
        Set rs = Nothing
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Public Sub ListView_Init1()
    Dim Column_to_size As Integer

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " ������ȣ    "
    ListView1.ColumnHeaders.Add , , " ��    ��    "
    ListView1.ColumnHeaders.Add , , " ��    ��        "
    ListView1.ColumnHeaders.Add , , " �� �� ó        "
    
    If (Glo_User_Type = "����1/����2") Then
        ListView1.ColumnHeaders.Add , , " �Ҽ�, ����   "
    Else
        ListView1.ColumnHeaders.Add , , " ��, ȣ��   "
    End If
    
    ListView1.ColumnHeaders.Add , , " �� �� ��  "
    ListView1.ColumnHeaders.Add , , " �� �� ��  "
    ListView1.ColumnHeaders.Add , , " ����Ͻ�  "
    ListView1.ColumnHeaders.Add , , "  "
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 2
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End Sub


'���α׷� ����
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg, Style, Title, Response
    Dim Ret As Boolean

    msg = "���α׷��� �����Ͻðڽ��ϱ�?         "
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Parking Manager��  - JWT   "
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then
        Call Err_doc("ȣ��Ʈ : " & "���α׷� ���������� ����")
        Call DataBaseClose(adoConn)
        
        Unload FrmTcpServer
        Unload FrmAccnt
        Unload FormIPCamera
        'Unload FormIPCameraPlayer
        '
    Dim i As Integer
    For i = 0 To UBound(Glo_FrmIPCameraPlayer)
        If (Not Glo_FrmIPCameraPlayer(i) Is Nothing) Then
            Unload Glo_FrmIPCameraPlayer(i)
        End If
    Next i

        Unload frmApsCmd
        Unload FrmCSV
        Unload FrmDeviceCFG
        Unload FrmExtend
        'Unload FrmG1
        'Unload FrmG4Mini
        'Unload FrmG6_23
        Unload FrmId
        Unload FrmImg
        Unload FrmInOut
        Unload frmLogin
        'Unload FrmPhoto
        Unload FrmReg
        Unload FrmRegHistory
        Unload frmResult
        Unload frmSplash
        Unload FrmTcpServer2
        'Unload Jung
        Unload MBox
        Unload Msg_Box
        Unload Pwd
        
        
        If (Glo_Screen_No = 4) Then
            If (Glo_GUEST_LANE1_YN = "Y") Then
                If (Not Glo_FrmGuest(0) Is Nothing) Then
                    Unload Glo_FrmGuest(0)
                    Set Glo_FrmGuest(0) = Nothing
                End If
            End If
            
            If (Glo_GUEST_LANE2_YN = "Y") Then
                If (Not Glo_FrmGuest(1) Is Nothing) Then
                    Unload Glo_FrmGuest(1)
                    Set Glo_FrmGuest(1) = Nothing
                End If
            End If
            If (Glo_GUEST_LANE3_YN = "Y") Then
                If (Not Glo_FrmGuest(2) Is Nothing) Then
                    Unload Glo_FrmGuest(2)
                    Set Glo_FrmGuest(2) = Nothing
                End If
            End If
            If (Glo_GUEST_LANE4_YN = "Y") Then
                If (Not Glo_FrmGuest(3) Is Nothing) Then
                    Unload Glo_FrmGuest(3)
                    Set Glo_FrmGuest(3) = Nothing
                End If
            End If
            
        End If
        
        
        Call Unhook
        End
    End If
    Me.MousePointer = 0
    Cancel = True
End Sub

'
Private Sub Form_G4Mini(Data As String)
Dim i As Integer
Dim gateNo As Integer
Dim GateName As String
Dim carno As String
Dim rs As Recordset
Dim qry As String
Dim Tmp_File As String
Dim bQryResult As Boolean

With FrmG4Mini
        gateNo = Left(Data, 1)
        i = LenH(Data)
        carno = Mid(Data, 3, (i - 2))

        qry = "Select * From tb_inout_ENC Where PASS_GATE = '" & gateNo & "' And CAR_NO = '" & carno & "' And(PASS_DATE >= '" & Format(Now, "yyyy-mm-dd") & " " & "00:00:00" & "' AND PASS_DATE <= '" & Format(Now, "yyyy-mm-dd") & " " & "23:59:59" & "') Order By PASS_DATE Desc"
        Set rs = New ADODB.Recordset
        'rs.Open Qry, adoConn
        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
        If (bQryResult = False) Then
            Call DataLogger("[FrmId]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
            Exit Sub
        End If

        If Not (rs.EOF) Then
            .lbl_carno(gateNo).Caption = "" & rs!CAR_NO
            Tmp_File = Dir(rs!pass_image)
            If (Tmp_File <> "") Then
                .ImageIn(gateNo).Picture = LoadPicture(rs!pass_image)
            End If
            For i = 0 To 3
                .Shp_Rec(i).Visible = False
            Next i
            .Shp_Rec(gateNo).Visible = True
            .lbl_time_now(gateNo).Caption = "" & rs!PASS_DATE
            .lbl_RecState(gateNo).Caption = "" & rs!PASS_RESULT
            If rs!Pass_YN = "Y" Then
                .lbl_RecState(gateNo).ForeColor = vbBlue
            Else
                .lbl_RecState(gateNo).ForeColor = vbRed
            End If
            .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " GateNo : " & gateNo & ", ������ȣ : " & rs!CAR_NO & ", ó����� : " & rs!PASS_RESULT, 0
        Else
            'Beep
        End If
        Set rs = Nothing
End With

On Error Resume Next

End Sub

Private Sub Timer1_Timer()
Dim qry As String
Dim rs As ADODB.Recordset
Dim i As Integer

    If (Glo_Certify = enumCertify.eCertTry And Glo_Cert_NoticeSDate < Format(Now, "yyyy-mm-dd")) Then
        LblTime(0).ForeColor = &HFF&
        LblTime(0).Caption = "[������������] " & "����ð� : " & Format(Now, "yyyy��mm��dd�� hh��nn��ss��")
    Else
        LblTime(0).ForeColor = &H0&
        LblTime(0).ToolTipText = ""
        LblTime(0).Caption = "����ð� : " & Format(Now, "yyyy��mm��dd�� hh��nn��ss��")
    End If
    
    'If (Format(Now, "NNSS") = "0001") Then
    '    '����Ʈ ī��Ʈ �ʱ�ȭ
    '    Qry = "show tables"
    '    Set rs = New ADODB.Recordset
    '    rs.Open Qry, adoConn
    '    Set rs = Nothing
    '    List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "    MySQL Connection Test...!! ", 0
    'End If
    
    If (Abs(Glo_Mon_LastInTime - Timer) >= 5) Then
        Glo_MonStat_Lane(0) = "DEAD"
        Glo_MonStat_Lane(1) = "DEAD"
        Glo_MonStat_Lane(2) = "DEAD"
        Glo_MonStat_Lane(3) = "DEAD"
    End If

    If (LANE1_YN = "Y") Then
        If (Glo_Mon_Lane(0) = True) Then
            If Glo_MonStat_Lane(0) = "LIVE" Then
                Imgshutdown(0).Visible = False
                ImgGreen(0).Visible = True
                ImgRed(0).Visible = False
'                List1.AddItem "lane1 Monitor Live", 0
                Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
            Else
                Imgshutdown(0).Visible = True
                ImgGreen(0).Visible = False
                ImgRed(0).Visible = True
'                List1.AddItem "lane1 Monitor Dead", 0
                Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
                'Call DataLogger("Lane1 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane1.exe")) Then
                Imgshutdown(0).Visible = False
                ImgGreen(0).Visible = True
                ImgRed(0).Visible = False
                Call FrmTcpServer.LPR_Alive_State_Send(0, "LIVE")
            Else
                Imgshutdown(0).Visible = True
                ImgGreen(0).Visible = False
                ImgRed(0).Visible = True
                Call FrmTcpServer.LPR_Alive_State_Send(0, "DEAD")
                'Call DataLogger("Lane1 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(0).Visible = False
        ImgGreen(0).Visible = False
        ImgRed(0).Visible = False
    End If
    
    If (LANE2_YN = "Y") Then
        If (Glo_Mon_Lane(1) = True) Then
            If Glo_MonStat_Lane(1) = "LIVE" Then
                Imgshutdown(1).Visible = False
                ImgGreen(1).Visible = True
                ImgRed(1).Visible = False
'                List1.AddItem "lane2 Monitor Live", 0
                Call FrmTcpServer.LPR_Alive_State_Send(1, "LIVE")
            Else
                Imgshutdown(1).Visible = True
                ImgGreen(1).Visible = False
                ImgRed(1).Visible = True
'                List1.AddItem "lane2 Monitor Dead", 0
                Call FrmTcpServer.LPR_Alive_State_Send(1, "DEAD")
                'Call DataLogger("Lane2 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane2.exe")) Then
                Imgshutdown(1).Visible = False
                ImgGreen(1).Visible = True
                ImgRed(1).Visible = False
                Call FrmTcpServer.LPR_Alive_State_Send(1, "LIVE")
            Else
                Imgshutdown(1).Visible = True
                ImgGreen(1).Visible = False
                ImgRed(1).Visible = True
                Call FrmTcpServer.LPR_Alive_State_Send(1, "DEAD")
                'Call DataLogger("Lane2 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(1).Visible = False
        ImgGreen(1).Visible = False
        ImgRed(1).Visible = False
    End If
    
    If (LANE3_YN = "Y") Then
        If (Glo_Mon_Lane(2) = True) Then
            If Glo_MonStat_Lane(2) = "LIVE" Then
                Imgshutdown(2).Visible = False
                ImgGreen(2).Visible = True
                ImgRed(2).Visible = False
'                List1.AddItem "lane3 Monitor Live", 0
                Call FrmTcpServer.LPR_Alive_State_Send(2, "LIVE")
            Else
                Imgshutdown(2).Visible = True
                ImgGreen(2).Visible = False
                ImgRed(2).Visible = True
'                List1.AddItem "lane3 Monitor Dead", 0
                Call FrmTcpServer.LPR_Alive_State_Send(2, "DEAD")
                'Call DataLogger("Lane3 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane3.exe")) Then
                Imgshutdown(2).Visible = False
                ImgGreen(2).Visible = True
                ImgRed(2).Visible = False
                Call FrmTcpServer.LPR_Alive_State_Send(2, "LIVE")
            Else
                Imgshutdown(2).Visible = True
                ImgGreen(2).Visible = False
                ImgRed(2).Visible = True
                Call FrmTcpServer.LPR_Alive_State_Send(2, "DEAD")
                'Call DataLogger("Lane3 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(2).Visible = False
        ImgGreen(2).Visible = False
        ImgRed(2).Visible = False
    End If
    
    If (LANE4_YN = "Y") Then
        If (Glo_Mon_Lane(3) = True) Then
            If Glo_MonStat_Lane(3) = "LIVE" Then
                Imgshutdown(3).Visible = False
                ImgGreen(3).Visible = True
                ImgRed(3).Visible = False
'                List1.AddItem "lane4 Monitor Live", 0
                Call FrmTcpServer.LPR_Alive_State_Send(3, "LIVE")
            Else
                Imgshutdown(3).Visible = True
                ImgGreen(3).Visible = False
                ImgRed(3).Visible = True
'                List1.AddItem "lane4 Monitor Dead", 0
                Call FrmTcpServer.LPR_Alive_State_Send(3, "DEAD")
                'Call DataLogger("Lane4 Monitor Stat : DEAD")
            End If
        Else
            If (Get_Process("Lane4.exe")) Then
                Imgshutdown(3).Visible = False
                ImgGreen(3).Visible = True
                ImgRed(3).Visible = False
                Call FrmTcpServer.LPR_Alive_State_Send(3, "LIVE")
            Else
                Imgshutdown(3).Visible = True
                ImgGreen(3).Visible = False
                ImgRed(3).Visible = True
                Call FrmTcpServer.LPR_Alive_State_Send(3, "DEAD")
                'Call DataLogger("Lane4 Stat : DEAD")
            End If
        End If
    Else
        Imgshutdown(3).Visible = False
        ImgGreen(3).Visible = False
        ImgRed(3).Visible = False
    End If


End Sub

Private Sub Socket_ConnectRemote(ByVal ip As String, ByVal Port As Long)

    If (Remote_Winsock.State <> sckClosed) Then
        Remote_Winsock.Close
        'DoEvents
    End If
    Remote_Winsock.Connect ip, Port
End Sub

Private Sub Remote_Winsock_Connect()
    Dim bData() As Byte

    ReDim bData(Len(Glo_Remote_Str) - 1) As Byte
    bData = StrConv(Glo_Remote_Str, vbFromUnicode)
    Remote_Winsock.SendData bData

End Sub

Private Sub Remote_Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    Remote_Winsock.GetData strData, , bytesTotal
    Remote_Winsock.Close
    
End Sub


Public Sub ListView_Init2()
Dim Column_to_size As Integer

    Call ListViewExtended(ListView2)
    ListView2.View = lvwReport
    ListView2.ListItems.Clear
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , " ó���Ͻ�                      "         '7
    ListView2.ColumnHeaders.Add , , " ������ȣ     "      '0
    ListView2.ColumnHeaders.Add , , " ��    ��         "  '1
    ListView2.ColumnHeaders.Add , , " ��    ��  "       '2
    ListView2.ColumnHeaders.Add , , " ��ȭ��ȣ     "  '3
    ListView2.ColumnHeaders.Add , , " �� �� ��     "   '4
    ListView2.ColumnHeaders.Add , , " �� �� ��     "        '5
    ListView2.ColumnHeaders.Add , , " ó������     "          '6
'    ListView2.ColumnHeaders.Add , , " ó���Ͻ�     "         '7
    'ListView2.ColumnHeaders.Add , , " ���ⱸ��     "    '8
    ListView2.ColumnHeaders.Add , , ""    '9(�̹������)
    ListView2.ColumnHeaders.Add , , ""    '10 '�̻��
    
    ListView2.ColumnHeaders.Add , , " "
    'ListView2.SortKey = 11
    ListView2.SortOrder = lvwDescending
    ListView2.Sorted = True
    
    For Column_to_size = 0 To ListView2.ColumnHeaders.Count - 2
         SendMessage ListView2.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim Tmp_File As String
    
    ListView2.SetFocus
    LblCar(1).Caption = ""
    LblName(1).Caption = ""
    LblId(1).Caption = ""
    LblCarType(1).Caption = ""
    LblTel(1).Caption = ""
    LblDate(1).Caption = ""
    LblGubun(1).Caption = ""
    
    LblCar(1).Caption = ListView2.SelectedItem.SubItems(1)  'ListView2.SelectedItem.Text
    LblId(1).Caption = ListView2.SelectedItem.SubItems(2) '��������
    LblName(1).Caption = ListView2.SelectedItem.SubItems(3) '�̸�
    LblCarType(1).Caption = ListView2.SelectedItem.SubItems(4) '����ó
    LblTel(1).Caption = Format(ListView2.SelectedItem.SubItems(6), "yyyy-mm-dd")
    LblDate(1).Caption = ListView2.SelectedItem.SubItems(7) '����������������
    LblGubun(1).Caption = ListView2.SelectedItem.text
        
'''    Tmp_File = Dir(Trim(ListView2.SelectedItem.SubItems(8)))
'''    If (Tmp_File <> "") Then
'''        ImageLog.Picture = LoadPicture(Trim(ListView2.SelectedItem.SubItems(8)))
'''    Else
'''        ImageLog.Picture = LoadPicture(App.Path & "\NoCar.jpg")
'''    End If
        If (IsFile(ListView2.SelectedItem.SubItems(8)) = True) Then
            ImageLog.Picture = LoadPicture(ListView2.SelectedItem.SubItems(8))
        Else
            ImageLog.Picture = LoadPicture(App.Path & "\NoCar.jpg")
        End If

End Sub
'���Ȳ ó�� END ===============================================================================================================================

