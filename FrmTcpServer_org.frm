VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmTcpServer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "TCP Server"
   ClientHeight    =   10395
   ClientLeft      =   7050
   ClientTop       =   2850
   ClientWidth     =   16755
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   16755
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " LANE5 "
      ForeColor       =   &H00FF0000&
      Height          =   5730
      Index           =   4
      Left            =   13440
      TabIndex        =   128
      Top             =   2010
      Width           =   3285
      Begin VB.ComboBox cmb_WizMode 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0000
         Left            =   120
         List            =   "FrmTcpServer.frx":000D
         TabIndex        =   148
         Text            =   "TCP"
         Top             =   2550
         Width           =   1530
      End
      Begin VB.ComboBox cmb_RelayComPort 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0020
         Left            =   2325
         List            =   "FrmTcpServer.frx":003F
         TabIndex        =   147
         Top             =   3315
         Width           =   810
      End
      Begin VB.TextBox txt_RelayPort 
         Height          =   330
         Index           =   4
         Left            =   1665
         TabIndex        =   146
         Text            =   "10000"
         Top             =   3315
         Width           =   630
      End
      Begin VB.ComboBox cmb_DispComPort 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":005E
         Left            =   2325
         List            =   "FrmTcpServer.frx":007D
         TabIndex        =   145
         Top             =   2925
         Width           =   810
      End
      Begin VB.TextBox txt_WizIP 
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   144
         Text            =   "192.168.111.111"
         Top             =   2925
         Width           =   1515
      End
      Begin VB.TextBox txt_DispPort 
         Height          =   330
         Index           =   4
         Left            =   1665
         TabIndex        =   143
         Text            =   "10000"
         Top             =   2925
         Width           =   630
      End
      Begin VB.ComboBox cmb_LPRMode 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":009C
         Left            =   120
         List            =   "FrmTcpServer.frx":00A9
         TabIndex        =   142
         Text            =   "TCP"
         Top             =   1440
         Width           =   1530
      End
      Begin VB.CheckBox chk_UseYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmd_OK 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   2295
         TabIndex        =   140
         Top             =   3810
         Width           =   825
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   4
         Left            =   975
         TabIndex        =   139
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_LPRPort 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   4
         Left            =   1650
         TabIndex        =   138
         Text            =   "10000"
         Top             =   1815
         Width           =   630
      End
      Begin VB.TextBox txt_LPRIP 
         Height          =   330
         Index           =   4
         Left            =   105
         TabIndex        =   137
         Text            =   "192.168.111.111"
         Top             =   1815
         Width           =   1515
      End
      Begin VB.TextBox txt_Disp1 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   136
         Text            =   "일단 정지..!!"
         Top             =   4560
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp2 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   135
         Text            =   "주차장내 절대 서행"
         Top             =   4890
         Width           =   2430
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "Gate"
         Height          =   330
         Index           =   4
         Left            =   1125
         TabIndex        =   134
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "Emg"
         Height          =   330
         Index           =   4
         Left            =   1815
         TabIndex        =   133
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "Nml"
         Height          =   330
         Index           =   4
         Left            =   2505
         TabIndex        =   132
         Top             =   5280
         Width           =   630
      End
      Begin VB.ComboBox cmb_Disp1 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":00C0
         Left            =   2550
         List            =   "FrmTcpServer.frx":00CD
         TabIndex        =   131
         Text            =   "녹"
         Top             =   4560
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp2 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":00DD
         Left            =   2550
         List            =   "FrmTcpServer.frx":00EA
         TabIndex        =   130
         Text            =   "녹"
         Top             =   4890
         Width           =   615
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "Cap"
         Height          =   330
         Index           =   4
         Left            =   435
         TabIndex        =   129
         Top             =   5280
         Width           =   630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   90
         X2              =   3165
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Relay"
         Height          =   210
         Index           =   22
         Left            =   135
         TabIndex        =   153
         Top             =   3435
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Device Control"
         Height          =   210
         Index           =   21
         Left            =   135
         TabIndex        =   152
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LPR"
         Height          =   210
         Index           =   20
         Left            =   135
         TabIndex        =   151
         Top             =   1215
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GateName"
         Height          =   210
         Index           =   19
         Left            =   135
         TabIndex        =   150
         Top             =   645
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dispaly"
         Height          =   210
         Index           =   18
         Left            =   135
         TabIndex        =   149
         Top             =   3240
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   90
         X2              =   3165
         Y1              =   4380
         Y2              =   4380
      End
   End
   Begin VB.TextBox txt_Ho 
      Height          =   315
      Left            =   11055
      TabIndex        =   126
      Text            =   "101"
      Top             =   1170
      Width           =   630
   End
   Begin VB.TextBox txt_Dong 
      Height          =   315
      Left            =   10365
      TabIndex        =   125
      Text            =   "102"
      Top             =   1170
      Width           =   630
   End
   Begin VB.CommandButton Command3 
      Caption         =   "세대통보"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11730
      TabIndex        =   124
      Top             =   1125
      Width           =   1260
   End
   Begin MSWinsockLib.Winsock LPR4_sock 
      Left            =   18915
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR3_sock 
      Left            =   18495
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR2_sock 
      Left            =   18075
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate4_sock 
      Left            =   18915
      Top             =   2085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate3_sock 
      Left            =   18495
      Top             =   2085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate2_sock 
      Left            =   18075
      Top             =   2085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSCommGate 
      Index           =   0
      Left            =   17655
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommDisp 
      Index           =   0
      Left            =   17655
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Left            =   17670
      Top             =   2085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Left            =   17655
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Left            =   17655
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " LANE4 "
      ForeColor       =   &H00FF0000&
      Height          =   5730
      Index           =   3
      Left            =   10095
      TabIndex        =   71
      Top             =   2010
      Width           =   3285
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "Cap"
         Height          =   330
         Index           =   3
         Left            =   435
         TabIndex        =   123
         Top             =   5280
         Width           =   630
      End
      Begin VB.ComboBox cmb_Disp2 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":00FA
         Left            =   2550
         List            =   "FrmTcpServer.frx":0107
         TabIndex        =   118
         Text            =   "녹"
         Top             =   4890
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":0117
         Left            =   2550
         List            =   "FrmTcpServer.frx":0124
         TabIndex        =   117
         Text            =   "녹"
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "Nml"
         Height          =   330
         Index           =   3
         Left            =   2505
         TabIndex        =   116
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "Emg"
         Height          =   330
         Index           =   3
         Left            =   1815
         TabIndex        =   115
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "Gate"
         Height          =   330
         Index           =   3
         Left            =   1125
         TabIndex        =   114
         Top             =   5280
         Width           =   630
      End
      Begin VB.TextBox txt_Disp2 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   113
         Text            =   "주차장내 절대 서행"
         Top             =   4890
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   112
         Text            =   "일단 정지..!!"
         Top             =   4560
         Width           =   2430
      End
      Begin VB.TextBox txt_LPRIP 
         Height          =   330
         Index           =   3
         Left            =   105
         TabIndex        =   83
         Text            =   "192.168.111.111"
         Top             =   1815
         Width           =   1515
      End
      Begin VB.TextBox txt_LPRPort 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   3
         Left            =   1650
         TabIndex        =   82
         Text            =   "10000"
         Top             =   1815
         Width           =   630
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   3
         Left            =   975
         TabIndex        =   81
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.CommandButton cmd_OK 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   2295
         TabIndex        =   80
         Top             =   3810
         Width           =   825
      End
      Begin VB.CheckBox chk_UseYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   960
      End
      Begin VB.ComboBox cmb_LPRMode 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":0134
         Left            =   120
         List            =   "FrmTcpServer.frx":0141
         TabIndex        =   78
         Text            =   "TCP"
         Top             =   1440
         Width           =   1530
      End
      Begin VB.TextBox txt_DispPort 
         Height          =   330
         Index           =   3
         Left            =   1665
         TabIndex        =   77
         Text            =   "10000"
         Top             =   2925
         Width           =   630
      End
      Begin VB.TextBox txt_WizIP 
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   76
         Text            =   "192.168.111.111"
         Top             =   2925
         Width           =   1515
      End
      Begin VB.ComboBox cmb_DispComPort 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":0158
         Left            =   2325
         List            =   "FrmTcpServer.frx":0177
         TabIndex        =   75
         Top             =   2925
         Width           =   810
      End
      Begin VB.TextBox txt_RelayPort 
         Height          =   330
         Index           =   3
         Left            =   1665
         TabIndex        =   74
         Text            =   "10000"
         Top             =   3315
         Width           =   630
      End
      Begin VB.ComboBox cmb_RelayComPort 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":0196
         Left            =   2325
         List            =   "FrmTcpServer.frx":01B5
         TabIndex        =   73
         Top             =   3315
         Width           =   810
      End
      Begin VB.ComboBox cmb_WizMode 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":01D4
         Left            =   120
         List            =   "FrmTcpServer.frx":01E1
         TabIndex        =   72
         Text            =   "TCP"
         Top             =   2550
         Width           =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   90
         X2              =   3165
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dispaly"
         Height          =   210
         Index           =   17
         Left            =   135
         TabIndex        =   88
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GateName"
         Height          =   210
         Index           =   16
         Left            =   135
         TabIndex        =   87
         Top             =   645
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LPR"
         Height          =   210
         Index           =   15
         Left            =   135
         TabIndex        =   86
         Top             =   1215
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Device Control"
         Height          =   210
         Index           =   14
         Left            =   135
         TabIndex        =   85
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Relay"
         Height          =   210
         Index           =   13
         Left            =   135
         TabIndex        =   84
         Top             =   3435
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   90
         X2              =   3165
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " LANE3 "
      ForeColor       =   &H00FF0000&
      Height          =   5730
      Index           =   2
      Left            =   6750
      TabIndex        =   53
      Top             =   2010
      Width           =   3285
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "Cap"
         Height          =   330
         Index           =   2
         Left            =   435
         TabIndex        =   122
         Top             =   5280
         Width           =   630
      End
      Begin VB.ComboBox cmb_Disp2 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":01F4
         Left            =   2550
         List            =   "FrmTcpServer.frx":0201
         TabIndex        =   111
         Text            =   "녹"
         Top             =   4890
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":0211
         Left            =   2550
         List            =   "FrmTcpServer.frx":021E
         TabIndex        =   110
         Text            =   "녹"
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "Nml"
         Height          =   330
         Index           =   2
         Left            =   2505
         TabIndex        =   109
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "Emg"
         Height          =   330
         Index           =   2
         Left            =   1815
         TabIndex        =   108
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "Gate"
         Height          =   330
         Index           =   2
         Left            =   1125
         TabIndex        =   107
         Top             =   5280
         Width           =   630
      End
      Begin VB.TextBox txt_Disp2 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   106
         Text            =   "주차장내 절대 서행"
         Top             =   4890
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   105
         Text            =   "일단 정지..!!"
         Top             =   4560
         Width           =   2430
      End
      Begin VB.TextBox txt_LPRIP 
         Height          =   330
         Index           =   2
         Left            =   105
         TabIndex        =   65
         Text            =   "192.168.111.111"
         Top             =   1815
         Width           =   1515
      End
      Begin VB.TextBox txt_LPRPort 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   2
         Left            =   1650
         TabIndex        =   64
         Text            =   "10000"
         Top             =   1815
         Width           =   630
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   2
         Left            =   975
         TabIndex        =   63
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.CommandButton cmd_OK 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   2295
         TabIndex        =   62
         Top             =   3810
         Width           =   825
      End
      Begin VB.CheckBox chk_UseYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   960
      End
      Begin VB.ComboBox cmb_LPRMode 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":022E
         Left            =   120
         List            =   "FrmTcpServer.frx":023B
         TabIndex        =   60
         Text            =   "TCP"
         Top             =   1440
         Width           =   1530
      End
      Begin VB.TextBox txt_DispPort 
         Height          =   330
         Index           =   2
         Left            =   1665
         TabIndex        =   59
         Text            =   "10000"
         Top             =   2925
         Width           =   630
      End
      Begin VB.TextBox txt_WizIP 
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Text            =   "192.168.111.111"
         Top             =   2925
         Width           =   1515
      End
      Begin VB.ComboBox cmb_DispComPort 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":0252
         Left            =   2325
         List            =   "FrmTcpServer.frx":0271
         TabIndex        =   57
         Top             =   2925
         Width           =   810
      End
      Begin VB.TextBox txt_RelayPort 
         Height          =   330
         Index           =   2
         Left            =   1665
         TabIndex        =   56
         Text            =   "10000"
         Top             =   3315
         Width           =   630
      End
      Begin VB.ComboBox cmb_RelayComPort 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":0290
         Left            =   2325
         List            =   "FrmTcpServer.frx":02AF
         TabIndex        =   55
         Top             =   3315
         Width           =   810
      End
      Begin VB.ComboBox cmb_WizMode 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":02CE
         Left            =   120
         List            =   "FrmTcpServer.frx":02DB
         TabIndex        =   54
         Text            =   "TCP"
         Top             =   2550
         Width           =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   90
         X2              =   3165
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dispaly"
         Height          =   210
         Index           =   12
         Left            =   135
         TabIndex        =   70
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GateName"
         Height          =   210
         Index           =   11
         Left            =   135
         TabIndex        =   69
         Top             =   645
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LPR"
         Height          =   210
         Index           =   10
         Left            =   135
         TabIndex        =   68
         Top             =   1215
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Device Control"
         Height          =   210
         Index           =   9
         Left            =   135
         TabIndex        =   67
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Relay"
         Height          =   210
         Index           =   8
         Left            =   135
         TabIndex        =   66
         Top             =   3435
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   90
         X2              =   3165
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " LANE2 "
      ForeColor       =   &H00FF0000&
      Height          =   5730
      Index           =   1
      Left            =   3405
      TabIndex        =   35
      Top             =   2010
      Width           =   3285
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "Cap"
         Height          =   330
         Index           =   1
         Left            =   450
         TabIndex        =   121
         Top             =   5280
         Width           =   630
      End
      Begin VB.TextBox txt_Disp1 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   104
         Text            =   "일단 정지..!!"
         Top             =   4560
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp2 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   103
         Text            =   "주차장내 절대 서행"
         Top             =   4890
         Width           =   2430
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "Gate"
         Height          =   330
         Index           =   1
         Left            =   1125
         TabIndex        =   102
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "Emg"
         Height          =   330
         Index           =   1
         Left            =   1815
         TabIndex        =   101
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "Nml"
         Height          =   330
         Index           =   1
         Left            =   2505
         TabIndex        =   100
         Top             =   5280
         Width           =   630
      End
      Begin VB.ComboBox cmb_Disp1 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":02EE
         Left            =   2550
         List            =   "FrmTcpServer.frx":02FB
         TabIndex        =   99
         Text            =   "녹"
         Top             =   4560
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp2 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":030B
         Left            =   2550
         List            =   "FrmTcpServer.frx":0318
         TabIndex        =   98
         Text            =   "녹"
         Top             =   4890
         Width           =   615
      End
      Begin VB.TextBox txt_LPRIP 
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   47
         Text            =   "192.168.111.111"
         Top             =   1815
         Width           =   1515
      End
      Begin VB.TextBox txt_LPRPort 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   1
         Left            =   1650
         TabIndex        =   46
         Text            =   "10000"
         Top             =   1815
         Width           =   630
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   1
         Left            =   975
         TabIndex        =   45
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.CommandButton cmd_OK 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2280
         TabIndex        =   44
         Top             =   3825
         Width           =   825
      End
      Begin VB.CheckBox chk_UseYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   960
      End
      Begin VB.ComboBox cmb_LPRMode 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":0328
         Left            =   120
         List            =   "FrmTcpServer.frx":0335
         TabIndex        =   42
         Text            =   "TCP"
         Top             =   1440
         Width           =   1530
      End
      Begin VB.TextBox txt_DispPort 
         Height          =   330
         Index           =   1
         Left            =   1665
         TabIndex        =   41
         Text            =   "10000"
         Top             =   2925
         Width           =   630
      End
      Begin VB.TextBox txt_WizIP 
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Text            =   "192.168.111.111"
         Top             =   2925
         Width           =   1515
      End
      Begin VB.ComboBox cmb_DispComPort 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":034C
         Left            =   2325
         List            =   "FrmTcpServer.frx":036B
         TabIndex        =   39
         Top             =   2925
         Width           =   810
      End
      Begin VB.TextBox txt_RelayPort 
         Height          =   330
         Index           =   1
         Left            =   1665
         TabIndex        =   38
         Text            =   "10000"
         Top             =   3315
         Width           =   630
      End
      Begin VB.ComboBox cmb_RelayComPort 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":038A
         Left            =   2325
         List            =   "FrmTcpServer.frx":03A9
         TabIndex        =   37
         Top             =   3315
         Width           =   810
      End
      Begin VB.ComboBox cmb_WizMode 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":03C8
         Left            =   120
         List            =   "FrmTcpServer.frx":03D5
         TabIndex        =   36
         Text            =   "TCP"
         Top             =   2550
         Width           =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   90
         X2              =   3165
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dispaly"
         Height          =   210
         Index           =   7
         Left            =   135
         TabIndex        =   52
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GateName"
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   51
         Top             =   645
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LPR"
         Height          =   210
         Index           =   5
         Left            =   135
         TabIndex        =   50
         Top             =   1215
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Device Control"
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   49
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Relay"
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   48
         Top             =   3435
         Width           =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   90
         X2              =   3165
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin LPR_PARKING_HOST.Server Server 
      Left            =   18105
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Net Stat "
      Height          =   1020
      Left            =   3405
      TabIndex        =   18
      Top             =   615
      Width           =   3285
      Begin VB.CommandButton Command2 
         Caption         =   "Client"
         Height          =   330
         Left            =   2250
         TabIndex        =   22
         Top             =   615
         Width           =   960
      End
      Begin VB.Label lblConnections 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "0 Current Connections"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   135
         TabIndex        =   21
         Top             =   225
         Width           =   1620
      End
      Begin VB.Label lblServerIP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server IP:"
         Height          =   210
         Left            =   135
         TabIndex        =   20
         Top             =   465
         Width           =   705
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server State:"
         Height          =   210
         Left            =   135
         TabIndex        =   19
         Top             =   705
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " LANE1 "
      ForeColor       =   &H00FF0000&
      Height          =   5730
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   2010
      Width           =   3285
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "Cap"
         Height          =   330
         Index           =   0
         Left            =   435
         TabIndex        =   120
         Top             =   5280
         Width           =   630
      End
      Begin VB.ComboBox cmb_Disp2 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":03E8
         Left            =   2550
         List            =   "FrmTcpServer.frx":03F5
         TabIndex        =   97
         Text            =   "녹"
         Top             =   4890
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":0405
         Left            =   2550
         List            =   "FrmTcpServer.frx":0412
         TabIndex        =   96
         Text            =   "녹"
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "Nml"
         Height          =   330
         Index           =   0
         Left            =   2505
         TabIndex        =   95
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "Emg"
         Height          =   330
         Index           =   0
         Left            =   1815
         TabIndex        =   94
         Top             =   5280
         Width           =   630
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "Gate"
         Height          =   330
         Index           =   0
         Left            =   1125
         TabIndex        =   93
         Top             =   5280
         Width           =   630
      End
      Begin VB.TextBox txt_Disp2 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   92
         Text            =   "주차장내 절대 서행"
         Top             =   4890
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Text            =   "일단 정지..!!"
         Top             =   4560
         Width           =   2430
      End
      Begin VB.ComboBox cmb_WizMode 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":0422
         Left            =   120
         List            =   "FrmTcpServer.frx":042F
         TabIndex        =   33
         Text            =   "TCP"
         Top             =   2550
         Width           =   1530
      End
      Begin VB.ComboBox cmb_RelayComPort 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":0442
         Left            =   2325
         List            =   "FrmTcpServer.frx":0461
         TabIndex        =   32
         Top             =   3315
         Width           =   810
      End
      Begin VB.TextBox txt_RelayPort 
         Height          =   330
         Index           =   0
         Left            =   1665
         TabIndex        =   31
         Text            =   "10000"
         Top             =   3315
         Width           =   630
      End
      Begin VB.ComboBox cmb_DispComPort 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":0480
         Left            =   2325
         List            =   "FrmTcpServer.frx":049F
         TabIndex        =   29
         Top             =   2925
         Width           =   810
      End
      Begin VB.TextBox txt_WizIP 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Text            =   "192.168.111.111"
         Top             =   2925
         Width           =   1515
      End
      Begin VB.TextBox txt_DispPort 
         Height          =   330
         Index           =   0
         Left            =   1665
         TabIndex        =   27
         Text            =   "10000"
         Top             =   2925
         Width           =   630
      End
      Begin VB.ComboBox cmb_LPRMode 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":04BE
         Left            =   120
         List            =   "FrmTcpServer.frx":04CB
         TabIndex        =   25
         Text            =   "TCP"
         Top             =   1440
         Width           =   1530
      End
      Begin VB.CheckBox chk_UseYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmd_OK 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2295
         TabIndex        =   16
         Top             =   3855
         Width           =   825
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   0
         Left            =   975
         TabIndex        =   14
         Text            =   "정문"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_LPRPort 
         Height          =   330
         Index           =   0
         Left            =   1650
         TabIndex        =   12
         Text            =   "10000"
         Top             =   1815
         Width           =   630
      End
      Begin VB.TextBox txt_LPRIP 
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Text            =   "192.168.111.111"
         Top             =   1815
         Width           =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   90
         X2              =   3165
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   90
         X2              =   3165
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Relay"
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   30
         Top             =   3435
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Device Control"
         Height          =   210
         Index           =   25
         Left            =   135
         TabIndex        =   26
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LPR"
         Height          =   210
         Index           =   24
         Left            =   135
         TabIndex        =   24
         Top             =   1215
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GateName"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   645
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dispaly"
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   3240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Remote IP / Port "
      Height          =   1020
      Left            =   6750
      TabIndex        =   7
      Top             =   615
      Width           =   3270
      Begin VB.CheckBox chk_RemoteYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use"
         Height          =   315
         Left            =   195
         TabIndex        =   89
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton CmdSvr 
         Caption         =   "SET"
         Height          =   360
         Left            =   2415
         TabIndex        =   17
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox TxtSvrIp 
         Height          =   315
         Left            =   195
         TabIndex        =   9
         Text            =   "255.255.255.255"
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox TxtSvrPort 
         Height          =   315
         Left            =   1620
         TabIndex        =   8
         Text            =   "10000"
         Top             =   570
         Width           =   615
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Refresh"
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   10500
      Value           =   1  '확인
      Width           =   945
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   -12765
      TabIndex        =   5
      Top             =   4530
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   12075
      TabIndex        =   4
      Top             =   90
      Width           =   1185
   End
   Begin VB.ListBox ListData 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2580
      Left            =   45
      TabIndex        =   3
      Top             =   7800
      Width           =   16680
   End
   Begin VB.Frame frameLocalInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Server IP/ Port "
      ForeColor       =   &H00000000&
      Height          =   1020
      Left            =   60
      TabIndex        =   0
      Top             =   615
      Width           =   3285
      Begin VB.CommandButton cmd_Svr 
         Caption         =   "SET"
         Height          =   360
         Left            =   2400
         TabIndex        =   119
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox txtIP 
         Height          =   315
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "255.255.255.255"
         Top             =   585
         Width           =   1380
      End
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Text            =   "10000"
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "Local IP and Port"
         Height          =   195
         Left            =   255
         TabIndex        =   90
         Top             =   285
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   17610
      Top             =   240
   End
   Begin MSCommLib.MSComm MSCommDisp 
      Index           =   1
      Left            =   18225
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommDisp 
      Index           =   2
      Left            =   18795
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommDisp 
      Index           =   3
      Left            =   19365
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommGate 
      Index           =   1
      Left            =   18225
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommGate 
      Index           =   2
      Left            =   18795
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommGate 
      Index           =   3
      Left            =   19365
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock Disp2_sock 
      Left            =   18075
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp3_sock 
      Left            =   18495
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp4_sock 
      Left            =   18915
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock RemoteR_sock 
      Left            =   19035
      Top             =   255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock RemoteS_sock 
      Left            =   18615
      Top             =   255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   10350
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock MvrSock 
      Left            =   4800
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock LPR5_sock 
      Left            =   19320
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp5_sock 
      Left            =   19320
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate5_sock 
      Left            =   19320
      Top             =   2085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSCommDisp 
      Index           =   4
      Left            =   19920
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommGate 
      Index           =   4
      Left            =   19920
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "세대통보"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   10320
      TabIndex        =   127
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BackStyle       =   0  '투명
      Caption         =   " # 주의 #  시스템 관리자 외에 절대 수정 금지..!!"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   180
      TabIndex        =   34
      Top             =   165
      Width           =   5145
   End
End
Attribute VB_Name = "FrmTcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const white = &H80000005
Const grey = &H8000000F

Private Sub sOutput(strIP As String, strText As String)
    If (Check2.value = 1) Then
        ListData.AddItem Format(Now, "YYYY-MM-DD HH:NN:SS") & "    " & strIP & "     " & strText, 0
    End If
End Sub

'LPR 통신방법 변경시
Private Sub cmb_LPRMode_Click(Index As Integer)

    Select Case Index
        Case 0
            Select Case cmb_LPRMode(0).Text
                Case "TCP"
                    txt_LPRPort(0).Text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(0).Text = Trim(LANE1_LPRPort)
            End Select
        Case 1
            Select Case cmb_LPRMode(1).Text
                Case "TCP"
                    txt_LPRPort(1).Text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(1).Text = Trim(LANE2_LPRPort)
            End Select
        Case 2
            Select Case cmb_LPRMode(2).Text
                Case "TCP"
                    txt_LPRPort(2).Text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(2).Text = Trim(LANE3_LPRPort)
            End Select
        Case 3
            Select Case cmb_LPRMode(3).Text
                Case "TCP"
                    txt_LPRPort(3).Text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(3).Text = Trim(LANE4_LPRPort)
            End Select
        Case 4
            Select Case cmb_LPRMode(4).Text
                Case "TCP"
                    txt_LPRPort(4).Text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(4).Text = Trim(LANE5_LPRPort)
            End Select
    End Select

    If (cmb_LPRMode(Index).Text = "TCP") Then
        txt_LPRPort(Index).Locked = True
        txt_LPRPort(Index).BackColor = &HE0E0E0
    Else
        txt_LPRPort(Index).Locked = False
        txt_LPRPort(Index).BackColor = &H80000005
    End If

End Sub

Private Sub cmd_CapTest_Click(Index As Integer)
    
    'Capture Test
    Call DataLogger("[Get Frame TEST]  Target Gate = " & Index)
    Call Relay_Out(1, Index)
    'Call Delay_Time(1000)
    
End Sub

Private Sub cmd_EmgTest_Click(Index As Integer)
    
    'Display Emg Test
    Call DataLogger("[DISPLAY Emg TEST]  Target Gate = " & Index)
    Call GL_Emergency("System Test", "System Test", 0, 30, 10, 1, 2, 1, Index)

End Sub

Private Sub cmd_GateTest_Click(Index As Integer)
    
    'Gate Test
    Call DataLogger("[GATE TEST]  Target Gate = " & Index)
    Call Relay_Out(0, Index)

End Sub

Private Sub cmd_HomeSet_Click()
    Call Put_Ini("System Config", "LANE1_Disp1Msg", txt_Disp1(0))
End Sub

Private Sub cmd_NmlTest_Click(Index As Integer)
    
    'Display Nomal Save
    Call DataLogger("[DISPLAY Nomal Save]  Target Gate = " & Index)
    Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, cmb_Disp1(Index).ListIndex, cmb_Disp2(Index).ListIndex, Index)
    
    Select Case Index
        Case 0
            Call Put_Ini("System Config", "LANE1_Disp1Msg", txt_Disp1(0))
            Call Put_Ini("System Config", "LANE1_Disp2Msg", txt_Disp2(0))
            Call Put_Ini("System Config", "LANE1_Disp1Color ", CStr(cmb_Disp1(0).ListIndex))
            Call Put_Ini("System Config", "LANE1_Disp2Color ", CStr(cmb_Disp2(0).ListIndex))
        
        Case 1
            Call Put_Ini("System Config", "LANE2_Disp1Msg", txt_Disp1(1))
            Call Put_Ini("System Config", "LANE2_Disp2Msg", txt_Disp2(1))
            Call Put_Ini("System Config", "LANE2_Disp1Color ", CStr(cmb_Disp1(1).ListIndex))
            Call Put_Ini("System Config", "LANE2_Disp2Color ", CStr(cmb_Disp2(1).ListIndex))
        
        Case 2
            Call Put_Ini("System Config", "LANE3_Disp1Msg", txt_Disp1(2))
            Call Put_Ini("System Config", "LANE3_Disp2Msg", txt_Disp2(2))
            Call Put_Ini("System Config", "LANE3_Disp1Color ", CStr(cmb_Disp1(2).ListIndex))
            Call Put_Ini("System Config", "LANE3_Disp2Color ", CStr(cmb_Disp2(2).ListIndex))
        
        Case 3
            Call Put_Ini("System Config", "LANE4_Disp1Msg", txt_Disp1(3))
            Call Put_Ini("System Config", "LANE4_Disp2Msg", txt_Disp2(3))
            Call Put_Ini("System Config", "LANE4_Disp1Color ", CStr(cmb_Disp1(3).ListIndex))
            Call Put_Ini("System Config", "LANE4_Disp2Color ", CStr(cmb_Disp2(3).ListIndex))
            
        Case 4
            Call Put_Ini("System Config", "LANE5_Disp1Msg", txt_Disp1(4))
            Call Put_Ini("System Config", "LANE5_Disp2Msg", txt_Disp2(4))
            Call Put_Ini("System Config", "LANE5_Disp1Color ", CStr(cmb_Disp1(4).ListIndex))
            Call Put_Ini("System Config", "LANE5_Disp2Color ", CStr(cmb_Disp2(4).ListIndex))
    
    End Select

End Sub

'LANE Config Save & Effect
Private Sub cmd_OK_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If chk_UseYN(0).value = "1" Then
                LANE1_YN = "Y"
            Else
                LANE1_YN = "N"
            End If
            LANE1_Name = Trim(txt_GateName(0))
            LANE1_LPRMode = cmb_LPRMode(0).ListIndex
            LANE1_LPRIP = Trim(txt_LPRIP(0))
            LANE1_LPRPort = Trim(txt_LPRPort(0))
            LANE1_DeviceMode = cmb_WizMode(0).ListIndex
            LANE1_DeviceIP = Trim(txt_WizIP(0))
            LANE1_DispPort = Trim(txt_DispPort(0))
            LANE1_RelayPort = Trim(txt_RelayPort(0))
            LANE1_DispComPort = cmb_DispComPort(0).Text
            LANE1_RelayComPort = cmb_RelayComPort(0).Text
            Call Put_Ini("System Config", "LANE1_YN ", LANE1_YN)
            Call Put_Ini("System Config", "LANE1_Name ", LANE1_Name)
            Call Put_Ini("System Config", "LANE1_LPRMode ", LANE1_LPRMode)
            Call Put_Ini("System Config", "LANE1_LPRIP ", LANE1_LPRIP)
            Call Put_Ini("System Config", "LANE1_LPRPort ", CStr(LANE1_LPRPort))
            Call Put_Ini("System Config", "LANE1_DeviceMode ", LANE1_DeviceMode)
            Call Put_Ini("System Config", "LANE1_DeviceIP ", LANE1_DeviceIP)
            Call Put_Ini("System Config", "LANE1_DispPort ", CStr(LANE1_DispPort))
            Call Put_Ini("System Config", "LANE1_RelayPort ", CStr(LANE1_RelayPort))
            Call Put_Ini("System Config", "LANE1_DispComPort ", CStr(LANE1_DispComPort))
            Call Put_Ini("System Config", "LANE1_RelayComPort ", CStr(LANE1_RelayComPort))
    
        Case 1
            If chk_UseYN(1).value = "1" Then
                LANE2_YN = "Y"
            Else
                LANE2_YN = "N"
            End If
            LANE2_Name = Trim(txt_GateName(1))
            LANE2_LPRMode = cmb_LPRMode(1).ListIndex
            LANE2_LPRIP = Trim(txt_LPRIP(1))
            LANE2_LPRPort = Trim(txt_LPRPort(1))
            LANE2_DeviceMode = cmb_WizMode(1).ListIndex
            LANE2_DeviceIP = Trim(txt_WizIP(1))
            LANE2_DispPort = Trim(txt_DispPort(1))
            LANE2_RelayPort = Trim(txt_RelayPort(1))
            LANE2_DispComPort = cmb_DispComPort(1).Text
            LANE2_RelayComPort = cmb_RelayComPort(1).Text
            Call Put_Ini("System Config", "LANE2_YN ", LANE2_YN)
            Call Put_Ini("System Config", "LANE2_Name ", LANE2_Name)
            Call Put_Ini("System Config", "LANE2_LPRMode ", LANE2_LPRMode)
            Call Put_Ini("System Config", "LANE2_LPRIP ", LANE2_LPRIP)
            Call Put_Ini("System Config", "LANE2_LPRPort ", CStr(LANE2_LPRPort))
            Call Put_Ini("System Config", "LANE2_DeviceMode ", LANE2_DeviceMode)
            Call Put_Ini("System Config", "LANE2_DeviceIP ", LANE2_DeviceIP)
            Call Put_Ini("System Config", "LANE2_DispPort ", CStr(LANE2_DispPort))
            Call Put_Ini("System Config", "LANE2_RelayPort ", CStr(LANE2_RelayPort))
            Call Put_Ini("System Config", "LANE2_DispComPort ", CStr(LANE2_DispComPort))
            Call Put_Ini("System Config", "LANE2_RelayComPort ", CStr(LANE2_RelayComPort))
    
        Case 2
            If chk_UseYN(2).value = "1" Then
                LANE3_YN = "Y"
            Else
                LANE3_YN = "N"
            End If
            LANE3_Name = Trim(txt_GateName(2))
            LANE3_LPRMode = cmb_LPRMode(2).ListIndex
            LANE3_LPRIP = Trim(txt_LPRIP(2))
            LANE3_LPRPort = Trim(txt_LPRPort(2))
            LANE3_DeviceMode = cmb_WizMode(2).ListIndex
            LANE3_DeviceIP = Trim(txt_WizIP(2))
            LANE3_DispPort = Trim(txt_DispPort(2))
            LANE3_RelayPort = Trim(txt_RelayPort(2))
            LANE3_DispComPort = cmb_DispComPort(2).Text
            LANE3_RelayComPort = cmb_RelayComPort(2).Text
            Call Put_Ini("System Config", "LANE3_YN ", LANE3_YN)
            Call Put_Ini("System Config", "LANE3_Name ", LANE3_Name)
            Call Put_Ini("System Config", "LANE3_LPRMode ", LANE3_LPRMode)
            Call Put_Ini("System Config", "LANE3_LPRIP ", LANE3_LPRIP)
            Call Put_Ini("System Config", "LANE3_LPRPort ", CStr(LANE3_LPRPort))
            Call Put_Ini("System Config", "LANE3_DeviceMode ", LANE3_DeviceMode)
            Call Put_Ini("System Config", "LANE3_DeviceIP ", LANE3_DeviceIP)
            Call Put_Ini("System Config", "LANE3_DispPort ", CStr(LANE3_DispPort))
            Call Put_Ini("System Config", "LANE3_RelayPort ", CStr(LANE3_RelayPort))
            Call Put_Ini("System Config", "LANE3_DispComPort ", CStr(LANE3_DispComPort))
            Call Put_Ini("System Config", "LANE3_RelayComPort ", CStr(LANE3_RelayComPort))
    
        Case 3
            If chk_UseYN(3).value = "1" Then
                LANE4_YN = "Y"
            Else
                LANE4_YN = "N"
            End If
            LANE4_Name = Trim(txt_GateName(3))
            LANE4_LPRMode = cmb_LPRMode(3).ListIndex
            LANE4_LPRIP = Trim(txt_LPRIP(3))
            LANE4_LPRPort = Trim(txt_LPRPort(3))
            LANE4_DeviceMode = cmb_WizMode(3).ListIndex
            LANE4_DeviceIP = Trim(txt_WizIP(3))
            LANE4_DispPort = Trim(txt_DispPort(3))
            LANE4_RelayPort = Trim(txt_RelayPort(3))
            LANE4_DispComPort = cmb_DispComPort(3).Text
            LANE4_RelayComPort = cmb_RelayComPort(3).Text
            Call Put_Ini("System Config", "LANE4_YN ", LANE4_YN)
            Call Put_Ini("System Config", "LANE4_Name ", LANE4_Name)
            Call Put_Ini("System Config", "LANE4_LPRMode ", LANE4_LPRMode)
            Call Put_Ini("System Config", "LANE4_LPRIP ", LANE4_LPRIP)
            Call Put_Ini("System Config", "LANE4_LPRPort ", CStr(LANE4_LPRPort))
            Call Put_Ini("System Config", "LANE4_DeviceMode ", LANE4_DeviceMode)
            Call Put_Ini("System Config", "LANE4_DeviceIP ", LANE4_DeviceIP)
            Call Put_Ini("System Config", "LANE4_DispPort ", CStr(LANE4_DispPort))
            Call Put_Ini("System Config", "LANE4_RelayPort ", CStr(LANE4_RelayPort))
            Call Put_Ini("System Config", "LANE4_DispComPort ", CStr(LANE4_DispComPort))
            Call Put_Ini("System Config", "LANE4_RelayComPort ", CStr(LANE4_RelayComPort))
            
        Case 4
            If chk_UseYN(4).value = "1" Then
                LANE5_YN = "Y"
            Else
                LANE5_YN = "N"
            End If
            LANE5_Name = Trim(txt_GateName(4))
            LANE5_LPRMode = cmb_LPRMode(4).ListIndex
            LANE5_LPRIP = Trim(txt_LPRIP(4))
            LANE5_LPRPort = Trim(txt_LPRPort(4))
            LANE5_DeviceMode = cmb_WizMode(4).ListIndex
            LANE5_DeviceIP = Trim(txt_WizIP(4))
            LANE5_DispPort = Trim(txt_DispPort(4))
            LANE5_RelayPort = Trim(txt_RelayPort(4))
            LANE5_DispComPort = cmb_DispComPort(4).Text
            LANE5_RelayComPort = cmb_RelayComPort(4).Text
            Call Put_Ini("System Config", "LANE5_YN ", LANE5_YN)
            Call Put_Ini("System Config", "LANE5_Name ", LANE5_Name)
            Call Put_Ini("System Config", "LANE5_LPRMode ", LANE5_LPRMode)
            Call Put_Ini("System Config", "LANE5_LPRIP ", LANE5_LPRIP)
            Call Put_Ini("System Config", "LANE5_LPRPort ", CStr(LANE5_LPRPort))
            Call Put_Ini("System Config", "LANE5_DeviceMode ", LANE5_DeviceMode)
            Call Put_Ini("System Config", "LANE5_DeviceIP ", LANE5_DeviceIP)
            Call Put_Ini("System Config", "LANE5_DispPort ", CStr(LANE5_DispPort))
            Call Put_Ini("System Config", "LANE5_RelayPort ", CStr(LANE5_RelayPort))
            Call Put_Ini("System Config", "LANE5_DispComPort ", CStr(LANE5_DispComPort))
            Call Put_Ini("System Config", "LANE5_RelayComPort ", CStr(LANE5_RelayComPort))
    
    End Select
        
    'MsgBox ("COM Prot 변경 시 프로그램을 다시 시작해주세요..!!")
    
    'Sever Refresh
    Call Server.StopServer
    Call Server.StartServer(Server_Port, Server.ServerIP)
    
    
End Sub

'Localhost Server Config Save
Private Sub cmd_Svr_Click()
    Dim i As Integer
        
    'Sever Refresh
    Call Server.StopServer
        
    Server_Port = Trim(txtPort)
    Call Put_Ini("System Config", "Server_Port", CStr(Server_Port))
        
    For i = 0 To 4
        If cmb_LPRMode(i).Text = "TCP" Then
            txt_LPRPort(i).Text = Trim(Server_Port)
            txt_LPRPort(i).Locked = True
            txt_LPRPort(i).BackColor = &HE0E0E0
        Else
            txt_LPRPort(i).Locked = False
            txt_LPRPort(i).BackColor = &H80000005
        End If
    Next i

    Call Server.StartServer(Server_Port, Server.ServerIP)
End Sub

'Remote Svr Config Save
Private Sub CmdSvr_Click()

'    Glo_Remote_IP = Trim(TxtSvrIp)
'    Glo_Remote_Port = Val(TxtSvrPort)
'
'    If chk_RemoteYN.value = "1" Then
'        Call Put_Ini("System Config", "Remote_YN", "Y")
'    Else
'        Call Put_Ini("System Config", "Remote_YN", "N")
'    End If
'    Call Put_Ini("System Config", "Remote_YN", Glo_Remote_IP)
'    Call Put_Ini("System Config", "Remote_YN", CStr(Glo_Remote_Port))

End Sub

'TCP Svr Hide
Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command2_Click()
    FrmClient.Show 0
End Sub



Private Sub Command3_Click()
    HomeNet_Dong = txt_Dong.Text
    HomeNet_Ho = txt_Ho.Text
    HomeNet_CarNo = "서울01가1234"
    
    HomeNet_Str = HomeNet_Dong & HomeNet_Ho & HomeNet_CarNo
    FrmTcpServer.HomeSock.SendData (HomeNet_Str)
End Sub


Public Sub InitSock()

'    If (adoConn <> "") Then
'        Call DataBaseClose(adoConn) '실행할 경우, 프로그램 종료시 오류메세지 발생
'    End If


    Call Server.StopServer

    If (LANE1_YN = "Y") Then
        LPR1_sock.Close
        Disp1_sock.Close
        Gate1_sock.Close
    End If
    
    If (LANE2_YN = "Y") Then
        LPR2_sock.Close
        Disp2_sock.Close
        Gate2_sock.Close
    End If
    
    If (LANE3_YN = "Y") Then
        LPR3_sock.Close
        Disp3_sock.Close
        Gate3_sock.Close
    End If
    
    If (LANE4_YN = "Y") Then
        LPR4_sock.Close
        Disp4_sock.Close
        Gate4_sock.Close
    End If
    
    If (LANE5_YN = "Y") Then
        LPR5_sock.Close
        Disp5_sock.Close
        Gate5_sock.Close
    End If

    If (MSCommGate(0).PortOpen = True) Then
        MSCommGate(0).PortOpen = False
    End If
    If (MSCommGate(1).PortOpen = True) Then
        MSCommGate(1).PortOpen = False
    End If
    If (MSCommGate(2).PortOpen = True) Then
        MSCommGate(2).PortOpen = False
    End If
    If (MSCommGate(3).PortOpen = True) Then
        MSCommGate(3).PortOpen = False
    End If
    If (MSCommGate(4).PortOpen = True) Then
        MSCommGate(4).PortOpen = False
    End If

    RemoteS_sock.Close
    RemoteR_sock.Close
    HomeSock.Close
    MvrSock.Close
'
End Sub


Private Sub Form_Load()
    Dim Port As Integer
    
On Error GoTo Err_Proc
    
    Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    Call InitSock
    Call Server.StartServer(Server_Port, Server.ServerIP)
    
    txtIP = Server.ServerIP
    txtPort = Server_Port
    
'    If Glo_Remote_YN = "Y" Then
'        chk_RemoteYN.value = 1
'    End If
'    TxtSvrIp = Glo_Remote_IP
'    TxtSvrPort = Glo_Remote_Port
    
    'Lane Config
    If LANE1_YN = "Y" Then
        chk_UseYN(0).value = 1
    End If
    txt_GateName(0).Text = LANE1_Name
    cmb_LPRMode(0).ListIndex = LANE1_LPRMode
    txt_LPRIP(0) = LANE1_LPRIP
    txt_LPRPort(0) = LANE1_LPRPort
    cmb_WizMode(0).ListIndex = LANE1_DeviceMode
    txt_WizIP(0).Text = LANE1_DeviceIP
    txt_DispPort(0).Text = LANE1_DispPort
    txt_RelayPort(0).Text = LANE1_RelayPort
    cmb_DispComPort(0).ListIndex = (LANE1_DispComPort - 1)
    cmb_RelayComPort(0).ListIndex = (LANE1_RelayComPort - 1)
    txt_Disp1(0) = LANE1_Disp1Msg
    txt_Disp2(0) = LANE1_Disp2Msg
    cmb_Disp1(0).ListIndex = LANE1_Disp1Color
    cmb_Disp2(0).ListIndex = LANE1_Disp2Color
    
    If LANE2_YN = "Y" Then
        chk_UseYN(1).value = 1
    End If
    txt_GateName(1).Text = LANE2_Name
    cmb_LPRMode(1).ListIndex = LANE2_LPRMode
    txt_LPRIP(1) = LANE2_LPRIP
    txt_LPRPort(1) = LANE2_LPRPort
    cmb_WizMode(1).ListIndex = LANE2_DeviceMode
    txt_WizIP(1).Text = LANE2_DeviceIP
    txt_DispPort(1).Text = LANE2_DispPort
    txt_RelayPort(1).Text = LANE2_RelayPort
    cmb_DispComPort(1).ListIndex = (LANE2_DispComPort - 1)
    cmb_RelayComPort(1).ListIndex = (LANE2_RelayComPort - 1)
    txt_Disp1(1) = LANE2_Disp1Msg
    txt_Disp2(1) = LANE2_Disp2Msg
    cmb_Disp1(1).ListIndex = LANE2_Disp1Color
    cmb_Disp2(1).ListIndex = LANE2_Disp2Color
    
    If LANE3_YN = "Y" Then
        chk_UseYN(2).value = 1
    End If
    txt_GateName(2).Text = LANE3_Name
    cmb_LPRMode(2).ListIndex = LANE3_LPRMode
    txt_LPRIP(2) = LANE3_LPRIP
    txt_LPRPort(2) = LANE3_LPRPort
    cmb_WizMode(2).ListIndex = LANE3_DeviceMode
    txt_WizIP(2).Text = LANE3_DeviceIP
    txt_DispPort(2).Text = LANE3_DispPort
    txt_RelayPort(2).Text = LANE3_RelayPort
    cmb_DispComPort(2).ListIndex = (LANE3_DispComPort - 1)
    cmb_RelayComPort(2).ListIndex = (LANE3_RelayComPort - 1)
    txt_Disp1(2) = LANE3_Disp1Msg
    txt_Disp2(2) = LANE3_Disp2Msg
    cmb_Disp1(2).ListIndex = LANE3_Disp1Color
    cmb_Disp2(2).ListIndex = LANE3_Disp2Color
    
    If LANE4_YN = "Y" Then
        chk_UseYN(3).value = 1
    End If
    txt_GateName(3).Text = LANE4_Name
    cmb_LPRMode(3).ListIndex = LANE4_LPRMode
    txt_LPRIP(3) = LANE4_LPRIP
    txt_LPRPort(3) = LANE4_LPRPort
    cmb_WizMode(3).ListIndex = LANE4_DeviceMode
    txt_WizIP(3).Text = LANE4_DeviceIP
    txt_DispPort(3).Text = LANE4_DispPort
    txt_RelayPort(3).Text = LANE4_RelayPort
    cmb_DispComPort(3).ListIndex = (LANE4_DispComPort - 1)
    cmb_RelayComPort(3).ListIndex = (LANE4_RelayComPort - 1)
    txt_Disp1(3) = LANE4_Disp1Msg
    txt_Disp2(3) = LANE4_Disp2Msg
    cmb_Disp1(3).ListIndex = LANE4_Disp1Color
    cmb_Disp2(3).ListIndex = LANE4_Disp2Color
    
    
    If LANE5_YN = "Y" Then
        chk_UseYN(4).value = 1
    End If
    txt_GateName(4).Text = LANE5_Name
    cmb_LPRMode(4).ListIndex = LANE5_LPRMode
    txt_LPRIP(4) = LANE5_LPRIP
    txt_LPRPort(4) = LANE5_LPRPort
    cmb_WizMode(4).ListIndex = LANE5_DeviceMode
    txt_WizIP(4).Text = LANE5_DeviceIP
    txt_DispPort(4).Text = LANE5_DispPort
    txt_RelayPort(4).Text = LANE5_RelayPort
    cmb_DispComPort(4).ListIndex = (LANE5_DispComPort - 1)
    cmb_RelayComPort(4).ListIndex = (LANE5_RelayComPort - 1)
    txt_Disp1(4) = LANE5_Disp1Msg
    txt_Disp2(4) = LANE5_Disp2Msg
    cmb_Disp1(4).ListIndex = LANE5_Disp1Color
    cmb_Disp2(4).ListIndex = LANE5_Disp2Color
    
    
    'Communication Config  ======================================================================================
    'Pcocess 통신을 하는 놈이 있는지?
    If (LANE1_YN = "Y" And LANE1_LPRMode = "2") Or (LANE2_YN = "Y" And LANE2_LPRMode = "2") Or (LANE3_YN = "Y" And LANE3_LPRMode = "2") Or (LANE4_YN = "Y" And LANE4_LPRMode = "2") Or (LANE5_YN = "Y" And LANE5_LPRMode = "2") Then
        gHW = Me.hwnd
    End If
    If (LANE1_YN = "Y") Then
        'LPR Engine
        Select Case LANE1_LPRMode
            Case "0"    'TCP
                LPR1_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Receive
                LPR1_sock.Protocol = sckUDPProtocol
                LPR1_sock.LocalPort = LANE1_LPRPort
                LPR1_sock.Bind
            Case "2"    'Process
                LANE1_Handle = FindWindow(vbNullString, "Lane1")
                SendMess WM_HOST_HANDLE & gHW, LANE1_Handle
        End Select
        
        'Wiznet Or COM Connection
        Select Case LANE1_DeviceMode
            Case "0"    'TCP
                Disp1_sock.Protocol = sckTCPProtocol
                Gate1_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Send
                Disp1_sock.Protocol = sckUDPProtocol
                Disp1_sock.RemoteHost = LANE1_DeviceIP
                Disp1_sock.RemotePort = LANE1_DispPort
                Gate1_sock.Protocol = sckUDPProtocol
                Gate1_sock.RemoteHost = LANE1_DeviceIP
                Gate1_sock.RemotePort = LANE1_RelayPort
            Case "2"    'COM
                MSCommDisp(0).CommPort = LANE1_DispComPort
                MSCommDisp(0).Settings = "115200,n,8,1"
                MSCommDisp(0).InputLen = 0
                MSCommDisp(0).InputMode = comInputModeBinary
                MSCommDisp(0).PortOpen = True
                If (MSCommDisp(0).PortOpen = True) Then
                    Call DataLogger("[전광판 Port 정상적으로 OPEN] COM Port = " & MSCommDisp(0).CommPort)
                Else
                    Call DataLogger("[전광판 Port OPEN 실패] Port번호 = " & MSCommDisp(0).CommPort & " 에러내용 : " & Err.Description)
                End If
                MSCommGate(0).CommPort = LANE1_RelayComPort
                MSCommGate(0).Settings = "9600,n,8,1"
                MSCommGate(0).InputLen = 0
                MSCommGate(0).InputMode = comInputModeBinary
                MSCommGate(0).PortOpen = True
                If (MSCommGate(0).PortOpen = True) Then
                    Call DataLogger("[IO Port 정상적으로 OPEN] Port번호 = " & MSCommGate(0).CommPort)
                Else
                    Call DataLogger("[IO Port OPEN 실패] Port번호 = " & MSCommGate(0).CommPort & " 에러내용 : " & Err.Description)
                End If
        End Select
    End If
    
    If (LANE2_YN = "Y") Then
        'LPR Engine
        Select Case LANE2_LPRMode
            Case "0"    'TCP
                LPR2_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Receive
                LPR2_sock.Protocol = sckUDPProtocol
                LPR2_sock.LocalPort = LANE2_LPRPort
                LPR2_sock.Bind
            Case "2"    'Process
                LANE2_Handle = FindWindow(vbNullString, "Lane2")
                SendMess WM_HOST_HANDLE & gHW, LANE2_Handle
        End Select
        
        'Wiznet Or COM Connection
        Select Case LANE2_DeviceMode
            Case "0"    'TCP
                Disp2_sock.Protocol = sckTCPProtocol
                Gate2_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Send
                Disp2_sock.Protocol = sckUDPProtocol
                Disp2_sock.RemoteHost = LANE2_DeviceIP
                Disp2_sock.RemotePort = LANE2_DispPort
                Gate2_sock.Protocol = sckUDPProtocol
                Gate2_sock.RemoteHost = LANE2_DeviceIP
                Gate2_sock.RemotePort = LANE2_RelayPort
            Case "2"    'COM
                MSCommDisp(1).CommPort = LANE2_DispComPort
                MSCommDisp(1).Settings = "115200,n,8,1"
                MSCommDisp(1).InputLen = 0
                MSCommDisp(1).InputMode = comInputModeBinary
                MSCommDisp(1).PortOpen = True
                If (MSCommDisp(1).PortOpen = True) Then
                    Call DataLogger("[전광판 Port 정상적으로 OPEN] COM Port = " & MSCommDisp(1).CommPort)
                Else
                    Call DataLogger("[전광판 Port OPEN 실패] Port번호 = " & MSCommDisp(1).CommPort & " 에러내용 : " & Err.Description)
                End If
                MSCommGate(1).CommPort = LANE2_RelayComPort
                MSCommGate(1).Settings = "9600,n,8,1"
                MSCommGate(1).InputLen = 0
                MSCommGate(1).InputMode = comInputModeBinary
                MSCommGate(1).PortOpen = True
                If (MSCommGate(1).PortOpen = True) Then
                    Call DataLogger("[IO Port 정상적으로 OPEN] Port번호 = " & MSCommGate(1).CommPort)
                Else
                    Call DataLogger("[IO Port OPEN 실패] Port번호 = " & MSCommGate(1).CommPort & " 에러내용 : " & Err.Description)
                End If
        End Select
    End If
    
    If (LANE3_YN = "Y") Then
        'LPR Engine
        Select Case LANE3_LPRMode
            Case "0"    'TCP
                LPR3_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Receive
                LPR3_sock.Protocol = sckUDPProtocol
                LPR3_sock.LocalPort = LANE3_LPRPort
                LPR3_sock.Bind
            Case "2"    'Process
                LANE3_Handle = FindWindow(vbNullString, "Lane3")
                SendMess WM_HOST_HANDLE & gHW, LANE3_Handle
        End Select
        
        'Wiznet Or COM Connection
        Select Case LANE3_DeviceMode
            Case "0"    'TCP
                Disp3_sock.Protocol = sckTCPProtocol
                Gate3_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Send
                Disp3_sock.Protocol = sckUDPProtocol
                Disp3_sock.RemoteHost = LANE3_DeviceIP
                Disp3_sock.RemotePort = LANE3_DispPort
                Gate3_sock.Protocol = sckUDPProtocol
                Gate3_sock.RemoteHost = LANE3_DeviceIP
                Gate3_sock.RemotePort = LANE3_RelayPort
            Case "2"    'COM
                MSCommDisp(2).CommPort = LANE3_DispComPort
                MSCommDisp(2).Settings = "115200,n,8,1"
                MSCommDisp(2).InputLen = 0
                MSCommDisp(2).InputMode = comInputModeBinary
                MSCommDisp(2).PortOpen = True
                If (MSCommDisp(2).PortOpen = True) Then
                    Call DataLogger("[전광판 Port 정상적으로 OPEN] COM Port = " & MSCommDisp(2).CommPort)
                Else
                    Call DataLogger("[전광판 Port OPEN 실패] Port번호 = " & MSCommDisp(2).CommPort & " 에러내용 : " & Err.Description)
                End If
                MSCommGate(2).CommPort = LANE3_RelayComPort
                MSCommGate(2).Settings = "9600,n,8,1"
                MSCommGate(2).InputLen = 0
                MSCommGate(2).InputMode = comInputModeBinary
                MSCommGate(2).PortOpen = True
                If (MSCommGate(2).PortOpen = True) Then
                    Call DataLogger("[IO Port 정상적으로 OPEN] Port번호 = " & MSCommGate(2).CommPort)
                Else
                    Call DataLogger("[IO Port OPEN 실패] Port번호 = " & MSCommGate(2).CommPort & " 에러내용 : " & Err.Description)
                End If
        End Select
    End If
    
    If (LANE4_YN = "Y") Then
        'LPR Engine
        Select Case LANE4_LPRMode
            Case "0"    'TCP
                LPR4_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Receive
                LPR4_sock.Protocol = sckUDPProtocol
                LPR4_sock.LocalPort = LANE4_LPRPort
                LPR4_sock.Bind
            Case "2"    'Process
                LANE4_Handle = FindWindow(vbNullString, "Lane4")
                SendMess WM_HOST_HANDLE & gHW, LANE4_Handle
        End Select
        
        'Wiznet Or COM Connection
        Select Case LANE4_DeviceMode
            Case "0"    'TCP
                Disp4_sock.Protocol = sckTCPProtocol
                Gate4_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Send
                Disp4_sock.Protocol = sckUDPProtocol
                Disp4_sock.RemoteHost = LANE4_DeviceIP
                Disp4_sock.RemotePort = LANE4_DispPort
                Gate4_sock.Protocol = sckUDPProtocol
                Gate4_sock.RemoteHost = LANE4_DeviceIP
                Gate4_sock.RemotePort = LANE4_RelayPort
            Case "2"    'COM
                MSCommDisp(3).CommPort = LANE4_DispComPort
                MSCommDisp(3).Settings = "115200,n,8,1"
                MSCommDisp(3).InputLen = 0
                MSCommDisp(3).InputMode = comInputModeBinary
                MSCommDisp(3).PortOpen = True
                If (MSCommDisp(3).PortOpen = True) Then
                    Call DataLogger("[전광판 Port 정상적으로 OPEN] COM Port = " & MSCommDisp(3).CommPort)
                Else
                    Call DataLogger("[전광판 Port OPEN 실패] Port번호 = " & MSCommDisp(3).CommPort & " 에러내용 : " & Err.Description)
                End If
                MSCommGate(3).CommPort = LANE4_RelayComPort
                MSCommGate(3).Settings = "9600,n,8,1"
                MSCommGate(3).InputLen = 0
                MSCommGate(3).InputMode = comInputModeBinary
                MSCommGate(3).PortOpen = True
                If (MSCommGate(3).PortOpen = True) Then
                    Call DataLogger("[IO Port 정상적으로 OPEN] Port번호 = " & MSCommGate(3).CommPort)
                Else
                    Call DataLogger("[IO Port OPEN 실패] Port번호 = " & MSCommGate(3).CommPort & " 에러내용 : " & Err.Description)
                End If
        End Select
    End If
    
    If (LANE5_YN = "Y") Then
        'LPR Engine
        Select Case LANE5_LPRMode
            Case "0"    'TCP
                LPR5_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Receive
                LPR5_sock.Protocol = sckUDPProtocol
                LPR5_sock.LocalPort = LANE5_LPRPort
                LPR5_sock.Bind
            Case "2"    'Process
                LANE5_Handle = FindWindow(vbNullString, "Lane5")
                SendMess WM_HOST_HANDLE & gHW, LANE5_Handle
        End Select
        
        'Wiznet Or COM Connection
        Select Case LANE5_DeviceMode
            Case "0"    'TCP
                Disp5_sock.Protocol = sckTCPProtocol
                Gate5_sock.Protocol = sckTCPProtocol
            Case "1"    'UDP Only Send
                Disp5_sock.Protocol = sckUDPProtocol
                Disp5_sock.RemoteHost = LANE5_DeviceIP
                Disp5_sock.RemotePort = LANE5_DispPort
                Gate5_sock.Protocol = sckUDPProtocol
                Gate5_sock.RemoteHost = LANE5_DeviceIP
                Gate5_sock.RemotePort = LANE5_RelayPort
            Case "2"    'COM
                MSCommDisp(4).CommPort = LANE5_DispComPort
                MSCommDisp(4).Settings = "115200,n,8,1"
                MSCommDisp(4).InputLen = 0
                MSCommDisp(4).InputMode = comInputModeBinary
                MSCommDisp(4).PortOpen = True
                If (MSCommDisp(4).PortOpen = True) Then
                    Call DataLogger("[전광판 Port 정상적으로 OPEN] COM Port = " & MSCommDisp(4).CommPort)
                Else
                    Call DataLogger("[전광판 Port OPEN 실패] Port번호 = " & MSCommDisp(4).CommPort & " 에러내용 : " & Err.Description)
                End If
                MSCommGate(4).CommPort = LANE4_RelayComPort
                MSCommGate(4).Settings = "9600,n,8,1"
                MSCommGate(4).InputLen = 0
                MSCommGate(4).InputMode = comInputModeBinary
                MSCommGate(4).PortOpen = True
                If (MSCommGate(4).PortOpen = True) Then
                    Call DataLogger("[IO Port 정상적으로 OPEN] Port번호 = " & MSCommGate(4).CommPort)
                Else
                    Call DataLogger("[IO Port OPEN 실패] Port번호 = " & MSCommGate(4).CommPort & " 에러내용 : " & Err.Description)
                End If
        End Select
    End If
    
    
'   'HOON 반드시 복구할 것
'    If (LANE1_YN = "Y" And LANE1_LPRMode = "2") Or (LANE2_YN = "Y" And LANE2_LPRMode = "2") Or (LANE3_YN = "Y" And LANE3_LPRMode = "2") Or (LANE4_YN = "Y" And LANE4_LPRMode = "2") Then
'        Call Hook
'    End If
   
    'Remote UDP 설정
    If (Glo_RemoteS_YN = "Y") Then
        RemoteS_sock.Protocol = sckUDPProtocol
        RemoteS_sock.RemoteHost = Glo_RemoteS_IP
        RemoteS_sock.RemotePort = Glo_RemoteS_Port
    End If
    
    If (Glo_RemoteR_YN = "Y") Then
        RemoteR_sock.Protocol = sckUDPProtocol
        RemoteR_sock.LocalPort = Glo_RemoteR_Port
        RemoteR_sock.Bind
    End If
    
    If (HomeNet_YN = "Y") Then
        HomeSock.Protocol = sckUDPProtocol
        HomeSock.RemoteHost = HomeNet_IP
        HomeSock.RemotePort = HomeNet_Port
    End If
    
    If (MVR_YN = "Y") Then
        MvrSock.Protocol = sckUDPProtocol
        MvrSock.RemoteHost = MVR_IP
        MvrSock.RemotePort = MVR_Port
    End If
    
    Timer1.Enabled = True

Exit Sub

Err_Proc:
    MsgBox ("[FormLoad_Proc]  " & Err.Description)
    Call DataLogger(" [TCP Server Load Proc]  " & Err.Description)


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call InitSock
End Sub
'
'Private Sub Server_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
'Dim sdata As String
'Dim Tmp_Path As String
'Dim car_num As String
'Dim Lpr_Cmd As String * 20
'Dim Lpr_CarNum As String * 20
'Dim Lpr_NumType As String * 2
'Dim Lpr_Path As String
'Dim Dns_Path As String
'Dim Tcp_Lpr_Path As String * 100
'Dim Lpr_Color As String * 10
'Dim image_name As String
'Dim Image_Path As String
'Dim url_name As String
'Dim fso As New FileSystemObject
'Dim tmp_gatenum As Integer
'Dim Mcnt As Integer
'Dim Pos As Integer
'Dim Loopcnt As Long
'
''New
'Dim i, GateNo As Integer
'Dim CarNum As String
'
'
'On Error GoTo Err_P
'
'    Call sOutput(RemoteIP, FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved.")
'    If Data = "GET_TIME" Then
'        Server.SendData Format(Time, "HH:MM:SS"), SckIndex
'        Exit Sub
'    End If
'    If Data = "GET_DATE" Then
'        Server.SendData Format(Date, "MM/DD/YYYY"), SckIndex
'        Exit Sub
'    End If
'    If (Mid(Data, 1, 8) = "LPR_TEST") Then
'        RemoteIP = Trim(Mid(Data, 21, 20))
'        Data = Mid(Data, 41, LenH(Data) - 40)
'        Server.SendData "LPR Test Command.", SckIndex
'    End If
'    Call sOutput(RemoteIP, Data)
'
'
'With Jung
'    Call DataLogger(" [Server DataArrival]  " & "---------------------------------------------------------------------")
'    'Call Err_doc("---------------------------------------------------------------------")
'    Call DataLogger(" [Server DataArrival]  " & "[LPR 데이터 수신(tcp)]  " & Data)
'    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LPR 데이터 수신(tcp)]  " & Data)
'    If (.Check1.value = 1) Then
'        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LPR 데이터 수신(tcp)] " & Data, 0
'    End If
'    If (Len(Data) > 100) Then
'        Exit Sub
'    End If
'
'    i = InStr(1, Data, "_", 1)
'    Glo_GateNo = Val(Left(Data, (i - 1)))
'    Select Case Glo_GateNo
'        Case 0
'            Glo_Lpr_IP = LANE1_LPRIP
'        Case 1
'            Glo_Lpr_IP = LANE2_LPRIP
'        Case 2
'            Glo_Lpr_IP = LANE3_LPRIP
'        Case 3
'            Glo_Lpr_IP = LANE4_LPRIP
'    End Select
'    If Glo_GateNo Mod 2 = 0 Then
'        Glo_GateGubun = 0
'    Else
'        Glo_GateGubun = 1
'    End If
'
'    s = InStr(4, Data, "_", 1)
'    CarNum = Mid(Data, (i + 1), (s - i - 1))
'    Glo_CarNum = CarNum
'    i = Len(Data)
'    Tmp_Path = Mid(Data, (s + 1), i)
'
'    If (HostType = 0) Then
'        'Call Form_Two_InOut(Data)
'        Call LPRIn_Proc(CarNum, Tmp_Path)
'        Call Jung_Show(CarNum)
'    Else
'        'Call Form_Two_In(Data)
'    End If
'    'Server.SendData Format(Now, "yyyymmddhhnnss"), SckIndex
'
'End With
'
'Exit Sub
'
'Err_P:
'    Call DataLogger(" [Server DataArrival]    " & Err.Description)
'    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [Server_DataArrival]  " & Err.Description)
'
'End Sub
'
'Private Sub Server_Error(ByVal SckIndex As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String)
'    Call sOutput("N/A", "Server Error! (" & Description & ")")
'End Sub
'
'Private Sub Server_ServerStarted()
'    Call sOutput("N/A", "Server Started! (" & Format(Time, "H:MM AM/PM") & ")")
'    txtIP.Locked = True
'    txtIP.BackColor = grey
'    'txtPort.Locked = True
'    txtPort.BackColor = grey
'End Sub
'
'Private Sub Server_ServerStopped()
'    Call sOutput("N/A", "Server Stopped! (" & Format(Time, "H:MM AM/PM") & ")")
'End Sub
'
'Private Sub Server_SocketClosed(ByVal SckIndex As Integer, ByVal LocalPort As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
'    Call sOutput(RemoteIP, "Connection closed. ")
'End Sub
'
'Private Sub Server_SocketOpened(ByVal SckIndex As Integer, ByVal LocalPort As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
'    Call sOutput(RemoteIP, "Connection opened. ")
'End Sub
'
'Private Sub Server_StartFailed()
'    Call sOutput("N/A", "Failed to Start Server! ")
'End Sub

Private Sub Timer1_Timer()
    
    Dim qry As String
    Dim rs As ADODB.Recordset

    If (Format(Now, "NNSS") = "0001") Then
        Call Time_Sync
        Call Inout_Reduce
    Else
        'List1.AddItem "  " & Format(Now, "yyyy-mm-dd hh:nn:ss"), 0
        'tmp = Format(Now, "HHNNSS")
    End If
    
    lblConnections = Server.ConnectionCount & " Current Connections"
    lblServerIP = "Server IP: " & Server.ServerIP
    lblState = "Server State: " & Server.State

End Sub
Public Sub Inout_Reduce()

Dim qry As String
Dim rs As ADODB.Recordset
Dim iDelDate As String
Dim iUseDay As Long

On Error GoTo Err_P

    If (Glo_INOUT_USING_DATE <> 99) Then
        
        iDelDate = DateAdd("m", Glo_INOUT_USING_DATE * (-1), Format(Now, "yyyy-mm-dd"))
'        Qry = "Delete from tb_inout WHERE PASS_DATE < '" & iDelDate & "'"
        adoConn.Execute "Delete from tb_inout WHERE PASS_DATE < '" & iDelDate & "'"
    
        Call DataLogger("DB Delete Table tb_inout")

    End If
On Error GoTo Err_P

Err_P:
    Call DataLogger("Inout_Reduce Proc Error")
End Sub

Public Sub Jung_Show(CarNum As String)
    Dim qry As String
    'Dim rs As ADODB.Recordset
    Dim Tmp_Path As String
    Dim itmX As ListItem
    Dim GateNo As Integer
    Dim inout As String
    Dim Gubun As String
    Dim i, s As Integer
    Dim ECHO As ICMP_ECHO_REPLY

On Error GoTo Err_P

    CarNum = Trim(CarNum)

    qry = "Select * From tb_inout Where PASS_GATE = '" & Glo_GateNo & "' And CAR_NO = '" & CarNum & "' And(PASS_DATE >= '" & Format(Now, "yyyy-mm-dd") & " " & "00:00:00" & "' AND PASS_DATE <= '" & Format(Now, "yyyy-mm-dd") & " " & "23:59:59" & "') Order By PASS_DATE Desc"
    'Debug.Print QRY
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn

With Jung
If (HostType = 0) Then
    If Not (rs.EOF) Then
        If (rs!PASS_INOUT = "IN") Then
            .lbl_title_in(0).Caption = "GATE : "
            
            If (User_Type = 0) Then
                .lbl_title_in(1).Caption = "이  름 : "
                .lbl_title_in(2).Caption = "연락처 : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_NAME
                .lbl_info_in(2).Caption = "" & rs!DRIVER_PHONE
            Else
                .lbl_info_in(1).Caption = "동    : "
                .lbl_info_in(2).Caption = "호    : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
            End If
            
            .lbl_title_in(3).Caption = "인식번호 : "
            .lbl_title_in(4).Caption = "종료일 : "
            .lbl_title_in(5).Caption = "입출상태 : " '성훈
            .lbl_info_in(0).Caption = "" & rs!PASS_GATE
            .lbl_info_in(3).Caption = "" & rs!REC_NO
            .lbl_info_in(4).Caption = "" & rs!End_Date
            .lbl_info_in(5).Caption = "" & rs!PASS_RESULT '성훈
            Select Case Trim(rs!PASS_RESULT)
                Case "정상입차"
                    .Proc_Type(0).Caption = " " & "정기권입차"
                    .Proc_Type(0).ForeColor = vbBlue '성훈
                Case Else
                    .Proc_Type(0).Caption = " " & rs!PASS_RESULT
                    .Proc_Type(0).ForeColor = vbRed '성훈
            End Select
            '==================================================================================================
            Call Ping(rs!PASS_IP, ECHO)
            If Left$(ECHO.Data, 1) <> Chr$(0) Then
                Tmp_Path = Dir(rs!PASS_IMAGE)
                If (Tmp_Path <> "") Then
                    .ImageIn(0).Picture = LoadPicture(rs!PASS_IMAGE)
                Else
                    .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            Else
                .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                Call DataLogger(" [Jung Show]    " & "Ping Test Failure...!!")
                Call DataLogger(" [Jung Show]    " & CarNum & "  " & Tmp_Path)
            End If
            '==================================================================================================
            .lbl_carno(0).Caption = rs!CAR_NO
            .lbl_time_now(0).Caption = rs!PASS_DATE
        Else
            .lbl_title_out(0).Caption = "GATE : "
            
            If (User_Type = 0) Then
                .lbl_title_out(1).Caption = "이  름 : "
                .lbl_title_out(2).Caption = "연락처 : "
                .lbl_info_out(1).Caption = "" & rs!DRIVER_NAME
                .lbl_info_out(2).Caption = "" & rs!DRIVER_PHONE
            Else
                .lbl_title_out(1).Caption = "동    : "
                .lbl_title_out(2).Caption = "호    : "
                .lbl_info_out(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_out(2).Caption = "" & rs!DRIVER_CLASS
            End If
            
            .lbl_title_out(3).Caption = "인식번호 : "
            .lbl_title_out(4).Caption = "종료일 : "
            .lbl_title_out(5).Caption = "입출상태 : " '성훈
            .lbl_info_out(0).Caption = "" & rs!PASS_GATE
            .lbl_info_out(3).Caption = "" & rs!REC_NO
            .lbl_info_out(4).Caption = "" & rs!End_Date
            .lbl_info_out(5).Caption = "" & rs!PASS_RESULT '성훈
            Select Case Trim(rs!PASS_RESULT)
                Case "정상출차"
                    .Proc_Type(1).Caption = " " & "정기권출차"
                    .Proc_Type(1).ForeColor = vbBlue '성훈
                Case Else
                    .Proc_Type(1).Caption = " " & rs!PASS_RESULT
                    .Proc_Type(1).ForeColor = vbRed '성훈
            End Select
            '==================================================================================================
            Call Ping(rs!PASS_IP, ECHO)
            If Left$(ECHO.Data, 1) <> Chr$(0) Then
                Tmp_Path = Dir(rs!PASS_IMAGE)
                If (Tmp_Path <> "") Then
                    .ImageIn(1).Picture = LoadPicture(rs!PASS_IMAGE)
                Else
                    .ImageIn(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            Else
                .ImageIn(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                Call DataLogger(" [Jung Show]    " & "Ping Test Failure...!!")
                Call DataLogger(" [Jung Show]    " & CarNum & "  " & Tmp_Path)
            End If
            '==================================================================================================
            .lbl_carno(1).Caption = rs!CAR_NO
            .lbl_time_now(1).Caption = rs!PASS_DATE
        End If
        Set itmX = .ListView2.ListItems.Add(, , "" & rs!CAR_NO)
        itmX.SubItems(1) = "" & rs!PASS_GATE
        itmX.SubItems(2) = "" & rs!DRIVER_NAME
        itmX.SubItems(3) = "" & rs!DRIVER_PHONE
        itmX.SubItems(4) = "" & rs!REC_NO
        itmX.SubItems(5) = "" & rs!End_Date
        itmX.SubItems(6) = "" & rs!PASS_RESULT
        itmX.SubItems(7) = "" & rs!PASS_DATE
        itmX.SubItems(8) = "" & rs!PASS_INOUT
        itmX.SubItems(9) = "" & rs!PASS_IMAGE
        '.ListView2.Sorted = False
        .ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
        .ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible
    End If
Else
    If Not (rs.EOF) Then
        If (rs!PASS_GATE = "0") Then
            .lbl_title_in(0).Caption = "GATE : "
            
            If (User_Type = 0) Then
                .lbl_title_in(1).Caption = "이  름 : "
                .lbl_title_in(2).Caption = "연락처 : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_NAME
                .lbl_info_in(2).Caption = "" & rs!DRIVER_PHONE
            Else
                .lbl_info_in(1).Caption = "동    : "
                .lbl_info_in(2).Caption = "호    : "
                .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
            End If
            
            .lbl_title_in(3).Caption = "인식번호 : "
            .lbl_title_in(4).Caption = "종료일 : "
            .lbl_title_in(5).Caption = "입출상태 : " '성훈
            .lbl_info_in(0).Caption = "" & rs!PASS_GATE
'            .lbl_info_in(1).Caption = "" & rs!DRIVER_DEPT
'            .lbl_info_in(2).Caption = "" & rs!DRIVER_CLASS
            .lbl_info_in(3).Caption = "" & rs!REC_NO
            .lbl_info_in(4).Caption = "" & rs!End_Date
            .lbl_info_in(5).Caption = "" & rs!PASS_RESULT '성훈
            Select Case Trim(rs!PASS_RESULT)
                Case "정상입차"
                    .Proc_Type(0).Caption = " " & "정기권입차"
                    .Proc_Type(0).ForeColor = vbBlue '성훈
                Case Else
                    .Proc_Type(0).Caption = " " & rs!PASS_RESULT
                    .Proc_Type(0).ForeColor = vbRed '성훈
            End Select
            '==================================================================================================
            Call Ping(rs!PASS_IP, ECHO)
            If Left$(ECHO.Data, 1) <> Chr$(0) Then
                Tmp_Path = Dir(rs!PASS_IMAGE)
                If (Tmp_Path <> "") Then
                    .ImageIn(0).Picture = LoadPicture(rs!PASS_IMAGE)
                Else
                    .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            Else
                .ImageIn(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                Call DataLogger(" [Jung Show]    " & "Ping Test Failure...!!")
                Call DataLogger(" [Jung Show]    " & CarNum & "  " & Tmp_Path)
            End If
            '==================================================================================================
            .lbl_carno(0).Caption = rs!CAR_NO
            .lbl_time_now(0).Caption = rs!PASS_DATE
        Else
            .lbl_title_out(0).Caption = "GATE : "
            
            If (User_Type = 0) Then
                .lbl_title_out(1).Caption = "이  름 : "
                .lbl_title_out(2).Caption = "연락처 : "
                .lbl_info_out(1).Caption = "" & rs!DRIVER_NAME
                .lbl_info_out(2).Caption = "" & rs!DRIVER_PHONE
            Else
                .lbl_title_out(1).Caption = "동    : "
                .lbl_title_out(2).Caption = "호    : "
                .lbl_info_out(1).Caption = "" & rs!DRIVER_DEPT
                .lbl_info_out(2).Caption = "" & rs!DRIVER_CLASS
            End If
            
            .lbl_title_out(3).Caption = "인식번호 : "
            .lbl_title_out(4).Caption = "종료일 : "
            .lbl_title_out(5).Caption = "입출상태 : " '성훈
            .lbl_info_out(0).Caption = "" & rs!PASS_GATE
'            .lbl_info_out(1).Caption = "" & rs!DRIVER_DEPT
'            .lbl_info_out(2).Caption = "" & rs!DRIVER_CLASS
            .lbl_info_out(3).Caption = "" & rs!REC_NO
            .lbl_info_out(4).Caption = "" & rs!End_Date
            .lbl_info_out(5).Caption = "" & rs!PASS_RESULT '성훈
            Select Case Trim(rs!PASS_RESULT)
                Case "정상입차"
                    .Proc_Type(1).Caption = " " & "정기권입차"
                    .Proc_Type(1).ForeColor = vbBlue '성훈
                Case Else
                    .Proc_Type(1).Caption = " " & rs!PASS_RESULT
                    .Proc_Type(1).ForeColor = vbRed '성훈
            End Select
            '==================================================================================================
            Call Ping(rs!PASS_IP, ECHO)
            If Left$(ECHO.Data, 1) <> Chr$(0) Then
                Tmp_Path = Dir(rs!PASS_IMAGE)
                If (Tmp_Path <> "") Then
                    .ImageIn(1).Picture = LoadPicture(rs!PASS_IMAGE)
                Else
                    .ImageIn(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                End If
            Else
                .ImageIn(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
                Call DataLogger(" [Jung Show]    " & "Ping Test Failure...!!")
                Call DataLogger(" [Jung Show]    " & CarNum & "  " & Tmp_Path)
            End If
            '==================================================================================================
            .lbl_carno(1).Caption = rs!CAR_NO
            .lbl_time_now(1).Caption = rs!PASS_DATE
        End If
        Set itmX = .ListView2.ListItems.Add(, , "" & rs!CAR_NO)
        itmX.SubItems(1) = "" & rs!PASS_GATE
        itmX.SubItems(2) = "" & rs!DRIVER_NAME
        itmX.SubItems(3) = "" & rs!DRIVER_PHONE
        itmX.SubItems(4) = "" & rs!REC_NO
        itmX.SubItems(5) = "" & rs!End_Date
        itmX.SubItems(6) = "" & rs!PASS_RESULT
        itmX.SubItems(7) = "" & rs!PASS_DATE
        itmX.SubItems(8) = "" & rs!PASS_INOUT
        itmX.SubItems(9) = "" & rs!PASS_IMAGE
        '.ListView2.Sorted = False
        .ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
        .ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible
    End If
End If
Set rs = Nothing

End With

Exit Sub

Err_P:
    Call DataLogger(" [Jung Show Proc]  " & Err.Description)

End Sub


'
''화면 두개 입구/출구 경우
'Private Sub Form_Two_InOut(Data As String)
'Dim i As Integer
'Dim GateNo As Integer
'Dim GateName As String
'Dim Result As String
'Dim CarNo As String
'Dim rs As Recordset
'Dim qry As String
'
'With main
'    GateNo = Left(Data, 1)
'    i = LenH(Data)
'    CarNo = Mid(Data, 3, (i - 2))
'
'    qry = "Select * From tb_inout Where PASS_GATE = '" & GateNo & "' And CAR_NO = '" & CarNo & "' And(PASS_DATE >= '" & Format(Now, "yyyy-mm-dd") & " " & "00:00:00" & "' AND PASS_DATE <= '" & Format(Now, "yyyy-mm-dd") & " " & "23:59:59" & "') Order By PASS_DATE Desc"
'
'    Set rs = New ADODB.Recordset
'    rs.Open qry, Ora_AdoConn
'
'    If Not (rs.EOF) Then
'        .ImageIn(GateNo).Picture = LoadPicture(rs!PASS_IMAGE)
'        .Proc_Type(GateNo).Caption = "" & rs!PASS_RESULT
'        If rs!PASS_YN = "Y" Then
'            .Proc_Type(GateNo).ForeColor = &HFF0000
'        Else
'            .Proc_Type(GateNo).ForeColor = &HFF&
'        End If
'
'        .lbl_carno(GateNo).Caption = rs!CAR_NO
'        .lbl_time_now(GateNo).Caption = rs!PASS_DATE
'
'        Select Case GateNo
'            Case 0
'                GateName = Glo_InGateName
'            Case 1
'                GateName = Glo_OutGateName
'            Case 2
'                GateName = Glo_InGateName1
'            Case 3
'                GateName = Glo_OutGateName1
'            Case 4
'                GateName = Glo_InGateName2
'            Case 5
'                GateName = Glo_OutGateName2
'        End Select
'
'        If rs!PASS_RESULT = "Y" Then
'            Result = "OPEN"
'        Else
'            Result = "Close"
'        End If
'
'        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " 입/출구 명칭 : " & GateName & ", 차량번호 : " & rs!CAR_NO & ", 처리결과 : " & Result, 0
'
'
'        '================================
'        Select Case GateNo
'            Case 0
'                For i = 0 To 6
'                    '.lbl_title_in(i).Caption = ""
'                    .lbl_info_in(i).Caption = ""
'                Next i
'            Case 1
'                For i = 0 To 6
'                    '.lbl_title_Out(i).Caption = ""
'                    .lbl_info_out(i).Caption = ""
'                Next i
'        End Select
'
'        For i = 0 To 1
'            .LblRecStat(i).Caption = " "
'        Next i
'        .LblRecStat(GateNo).Caption = "Detect..!!"
'
'        Select Case GateNo
'            Case 0
'                .lbl_info_in(0).Caption = "" & rs!DRIVER_NAME
'                .lbl_info_in(1).Caption = "" & rs!DRIVER_PHONE
'                .lbl_info_in(2).Caption = "" & rs!DRIVER_DEPT
'                .lbl_info_in(3).Caption = "" & rs!DRIVER_CLASS
'                .lbl_info_in(4).Caption = "" & rs!Start_Date & " - " & rs!End_Date
'                .lbl_info_in(5).Caption = "" & rs!CAR_MODEL
'            Case 1
'                .lbl_info_out(0).Caption = "" & rs!DRIVER_NAME
'                .lbl_info_out(1).Caption = "" & rs!DRIVER_PHONE
'                .lbl_info_out(2).Caption = "" & rs!DRIVER_DEPT
'                .lbl_info_out(3).Caption = "" & rs!DRIVER_CLASS
'                .lbl_info_out(4).Caption = "" & rs!Start_Date & " - " & rs!End_Date
'                .lbl_info_out(5).Caption = "" & rs!CAR_MODEL
'        End Select
'
'        Set itmX = .ListView2.ListItems.Add(1, , "" & Format(Now, "yyyy-mm-dd hh:nn:ss"))
'        itmX.SubItems(1) = "" & rs!PASS_RESULT
'        itmX.SubItems(2) = "" & rs!CAR_NO
'        itmX.SubItems(3) = "" & rs!CAR_GUBUN
'        itmX.SubItems(4) = "" & rs!DRIVER_NAME
'        itmX.SubItems(5) = "" & rs!DRIVER_PHONE
'        itmX.SubItems(6) = "" & rs!DRIVER_DEPT
'        itmX.SubItems(7) = "" & rs!DRIVER_CLASS
'        itmX.SubItems(8) = "" & rs!Start_Date
'        itmX.SubItems(9) = "" & rs!End_Date
'        itmX.SubItems(10) = "" & rs!PASS_IMAGE
'        itmX.SubItems(11) = "" & GateNo
'        itmX.SubItems(12) = "" & rs!PASS_IP
'            .ListView2.Sorted = True
'            .ListView2.ListItems(1).Selected = True
'            .ListView2.ListItems(1).EnsureVisible
'            '.ListView2.SetFocus
'    Else
'        Beep
'    End If
'End With
'
'Set rs = Nothing
'
'End Sub

''화면 두개 입구만 둘일 경우
'Private Sub Form_Two_In(Data As String)
'Dim i As Integer
'Dim GateNo As Integer
'Dim inout As Integer
'Dim GateName As String
'Dim Result As String
'Dim CarNo As String
'Dim rs As Recordset
'Dim qry As String
'Dim Tmp_File As String
'
'With main
'
'    GateNo = Left(Data, 1)
'    i = LenH(Data)
'    CarNo = Mid(Data, 3, (i - 2))
'
'    qry = "Select * From tb_inout Where PASS_GATE = '" & GateNo & "' And CAR_NO = '" & CarNo & "' And(PASS_DATE >= '" & Format(Now, "yyyy-mm-dd") & " " & "00:00:00" & "' AND PASS_DATE <= '" & Format(Now, "yyyy-mm-dd") & " " & "23:59:59" & "') Order By PASS_DATE Desc"
'
'    Set rs = New ADODB.Recordset
'    rs.Open qry, Ora_AdoConn
'
'    If Not (rs.EOF) Then
'
'        Select Case rs!PASS_GATE
'            Case 0
'                If GuestGate = 0 Then
'                    inout = 0
'                Else
'                    inout = 1
'                End If
'            Case 1
'                Set rs = Nothing
'                Exit Sub
'            Case 2
'                If GuestGate = 2 Then
'                    inout = 0
'                Else
'                    inout = 1
'                End If
'            Case 3
'                Set rs = Nothing
'                Exit Sub
'            Case 4
'                If GuestGate = 4 Then
'                    inout = 0
'                Else
'                    inout = 1
'                End If
'            Case 5
'                Set rs = Nothing
'                Exit Sub
'        End Select
'
'        Tmp_File = Dir(rs!PASS_IMAGE)
'        If (Tmp_File <> "") Then
'            .ImageIn(inout).Picture = LoadPicture(rs!PASS_IMAGE)
'        Else
'
'        End If
'
'        .Proc_Type(inout).Caption = "" & rs!PASS_RESULT
'
'        If (rs!PASS_YN = "Y") Then
'            .Proc_Type(inout).ForeColor = &HFF0000
'        Else
'            .Proc_Type(inout).ForeColor = &HFF&
'        End If
'
'        .lbl_carno(inout).Caption = rs!CAR_NO
'        .lbl_time_now(inout).Caption = rs!PASS_DATE
'
'        Select Case GateNo
'            Case 0
'                GateName = Glo_InGateName
'            Case 1
'                GateName = Glo_OutGateName
'            Case 2
'                GateName = Glo_InGateName1
'            Case 3
'                GateName = Glo_OutGateName1
'            Case 4
'                GateName = Glo_InGateName2
'            Case 5
'                GateName = Glo_OutGateName2
'        End Select
'
'        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " 입/출구 명칭 : " & GateName & ", 차량번호 : " & rs!CAR_NO & ", 처리결과 : " & rs!PASS_RESULT, 0
'
'        '================================
'        Select Case inout
'            Case 0
'                For i = 0 To 6
'                    '.lbl_title_in(i).Caption = ""
'                    .lbl_info_in(i).Caption = ""
'                Next i
'            Case 1
'                For i = 0 To 6
'                    '.lbl_title_Out(i).Caption = ""
'                    .lbl_info_out(i).Caption = ""
'                Next i
'        End Select
'
'        For i = 0 To 1
'            .LblRecStat(i).Caption = " "
'        Next i
'        .LblRecStat(inout).Caption = "Detect..!!"
'
'        Select Case inout
'            Case 0
'                .lbl_info_in(0).Caption = "" & rs!DRIVER_NAME
'                .lbl_info_in(1).Caption = "" & rs!DRIVER_PHONE
'                .lbl_info_in(2).Caption = "" & rs!DRIVER_DEPT
'                .lbl_info_in(3).Caption = "" & rs!DRIVER_CLASS
'                .lbl_info_in(4).Caption = "" & rs!Start_Date & " - " & rs!End_Date
'                .lbl_info_in(5).Caption = "" & rs!CAR_MODEL
'            Case 1
'                .lbl_info_out(0).Caption = "" & rs!DRIVER_NAME
'                .lbl_info_out(1).Caption = "" & rs!DRIVER_PHONE
'                .lbl_info_out(2).Caption = "" & rs!DRIVER_DEPT
'                .lbl_info_out(3).Caption = "" & rs!DRIVER_CLASS
'                .lbl_info_out(4).Caption = "" & rs!Start_Date & " - " & rs!End_Date
'                .lbl_info_out(5).Caption = "" & rs!CAR_MODEL
'        End Select
'
'        Set itmX = .ListView2.ListItems.Add(1, , "" & Format(Now, "yyyy-mm-dd hh:nn:ss"))
'        itmX.SubItems(1) = "" & rs!PASS_RESULT
'        itmX.SubItems(2) = "" & rs!CAR_NO
'        itmX.SubItems(3) = "" & rs!CAR_GUBUN
'        itmX.SubItems(4) = "" & rs!DRIVER_NAME
'        itmX.SubItems(5) = "" & rs!DRIVER_PHONE
'        itmX.SubItems(6) = "" & rs!DRIVER_DEPT
'        itmX.SubItems(7) = "" & rs!DRIVER_CLASS
'        itmX.SubItems(8) = "" & rs!Start_Date
'        itmX.SubItems(9) = "" & rs!End_Date
'        itmX.SubItems(10) = "" & rs!PASS_IMAGE
'        itmX.SubItems(11) = "" & GateNo
'        itmX.SubItems(12) = "" & rs!PASS_IP
'        .ListView2.Sorted = True
'        .ListView2.ListItems(1).Selected = True
'        .ListView2.ListItems(1).EnsureVisible
'        .ListView2.SetFocus
'
'        '세대통보
'        If (HomeAlarm_Mode <> 0) Then
'            If (IsNumeric(rs!DRIVER_DEPT) = True) And (IsNumeric(rs!DRIVER_CLASS) = True) Then
'                Call HomeAlarm_Proc(rs!DRIVER_DEPT, rs!DRIVER_CLASS, rs!CAR_NO)
'            End If
'        End If
'
'    Else
'        'Beep
'    End If
'
'End With
'
'Set rs = Nothing
'
'On Error Resume Next
'
'End Sub

'
''세대통보 프로세스
'Public Sub HomeAlarm_Proc(Dong As Integer, Ho As Integer, CarNo As String)
'    Dim i As Integer
'    Dim qry As String
'    Dim rs As ADODB.Recordset
'    Dim tmpRS As ADODB.Recordset
'    Dim day_Err_F As Boolean
'    Dim Today_Num As Integer
'    Dim Today_End_Num As Integer
'    Dim Car_End_No As Integer
'    Dim Car_Kind As String
'    Dim itmX As ListItem
'    Dim Car_i As Integer
'    Dim Car_Num_Str As String
'    Dim Image_Path As String
'    Dim inout As Integer
'
'On Error GoTo Err_Proc
'
'With main
'
'    Select Case HomeAlarm_Mode
'        Case 1  '현대정보통신
'            'Call HyunDaeCom_Proc(Dong, Ho, CarNo)
'        Case 2
'
'        Case 3
'
'        Case Else
'
'    End Select
'
'End With
'
'Exit Sub
'
'Err_Proc:
'    Call DataLogger(" [Home Alarm Proc]    " & Err.Description)
'    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LPRIn_Proc]  " & Err.Description)
'End Sub


Public Sub Disp1_sock_Connect()
Dim bData() As Byte

'GlO_TcpDataDisp = "0"

'ReDim BData(Len(GlO_TcpDataDisp) - 1) As Byte
'BData = StrConv(GlO_TcpDataDisp, vbFromUnicode)

Disp1_sock.SendData GloDisp_BData
'Call sOutput(Glo_Disp_IP & " [전광판 Device 송신] ")
'Call Write_log("    [전광판 Device 송신] ")
'Disp1_sock.Close
End Sub

Public Sub Disp1_sock_DataArrival(ByVal bytesTotal As Long)

Dim strData As String
Dim bData() As Byte
Dim i As Integer


ReDim bData(bytesTotal)

Disp1_sock.GetData bData, , bytesTotal

'For i = 0 To bytesTotal - 1
'    Debug.Print Hex(bdata(i))
'Next i


'Disp1_sock.GetData strData, , bytesTotal
'Call sOutput(Glo_Disp_IP & " [전광판 Device 수신] ")


Disp1_sock.Close
End Sub
Public Sub Disp1_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Call sOutput(Glo_Disp_IP & " [전광판 Device 소켓] " & "에러 : " & Description)
'Call Write_log("    [전광판 Device 소켓] " & "에러 : " & Description)

End Sub

Public Sub Disp1_sock_SendComplete()
'Disp1_sock.Close
End Sub

Public Sub Gate1_sock_SendComplete()
'Gate1_sock.Close
End Sub


Public Sub Gate1_sock_Connect()
Dim bData() As Byte

ReDim bData(Len(GlO_TcpDataGate) - 1) As Byte
bData = StrConv(GlO_TcpDataGate, vbFromUnicode)
Gate1_sock.SendData bData
'Call sOutput(Glo_Gate_IP & " [Gate Device 송신] " & GlO_TcpDataGate)
'Call Write_log("    [Gate Device 송신] " & GlO_TcpDataGate)

End Sub

Public Sub Gate1_sock_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

Gate1_sock.GetData strData, , bytesTotal
'Call sOutput(Glo_Gate_IP & " [Gate Device 수신] " & strData)
Gate1_sock.Close

End Sub

Public Sub Gate1_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Call sOutput(Glo_Gate_IP & " [Gate Device 소켓] " & "에러 : " & Description)
'Call Write_log("    [Gate Device 소켓] " & "에러 : " & Description)
End Sub

Public Sub UDP_Proc(sdata As String)
    'Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P

With FrmG4Mini
    If (.Check1.value = 1) Then
        .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & " [LPR 데이터 수신(tcp)] " & Data, 0
    End If
    If (Len(Data) > 100) Then
        Exit Sub
    End If
    
    i = InStr(1, sdata, "_", 1)
    Glo_GateNo = Val(Left(sdata, (i - 1)))
    Select Case Glo_GateNo
        Case 0
            Glo_Lpr_IP = LANE1_LPRIP
            Glo_Lane_Inout = LANE1_Inout
            Glo_FreePass = Glo_FreePassLane1_YN
        Case 1
            Glo_Lpr_IP = LANE2_LPRIP
            Glo_Lane_Inout = LANE2_Inout
            Glo_FreePass = Glo_FreePassLane2_YN
        Case 2
            Glo_Lpr_IP = LANE3_LPRIP
            Glo_Lane_Inout = LANE3_Inout
            Glo_FreePass = Glo_FreePassLane3_YN
        Case 3
            Glo_Lpr_IP = LANE4_LPRIP
            Glo_Lane_Inout = LANE4_Inout
            Glo_FreePass = Glo_FreePassLane4_YN
    End Select
    
    Select Case HostType
        Case 2
            Glo_GateGubun = 0
        
        Case 3
            If Glo_GateNo Mod 2 = 0 Then
                Glo_GateGubun = 0
            Else
                Glo_GateGubun = 1
            End If
    End Select
    
    s = InStr(4, sdata, "_", 1)
    CarNum = Mid(sdata, (i + 1), (s - i - 1))
    Glo_CarNum = CarNum
    i = Len(sdata)
    Tmp_Path = Mid(sdata, (s + 1), i)
    
    Call LPRIn_Proc(CarNum, Tmp_Path)
    
    
    '스크린 수에 따라서 분기
    If (Glo_Screen_No = 4) Then
        If (Glo_GateNo < Glo_Screen_No) Then
            Call G4Mini_4INShow(CarNum)
        End If
    
    ElseIf (Glo_Screen_No = 2) Then
        If (Glo_GateNo < Glo_Screen_No) Then
            Call Jung_Show(CarNum)
        End If
    End If
    
    'UDP끊으면 안돼~~~
    'LPR1_sock.Close

    'Remote Send Data
    If Glo_RemoteS_YN = "Y" Then
        Glo_Remote_Str = Glo_GateNo & "_" & CarNum
        RemoteS_sock.SendData (Glo_Remote_Str)
        Call DataLogger("[Remote UDP 전송]  DATA = " & Glo_Remote_Str)
    End If

    If MVR_YN = "Y" Then
        MVR_Str = (Glo_GateNo + 1) & " " & CarNum
        MvrSock.SendData (Trim(MVR_Str))
        
        Call DataLogger("[MVR UDP 전송]  DATA = " & MVR_Str)
    End If


End With

Exit Sub

Err_P:
    Call DataLogger(" [UDP Proc]  " & Err.Description)

End Sub

Public Sub LPR1_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    LPR1_sock.GetData sdata, , bytesTotal
    Call DataLogger("Lane1 UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    Call UDP_Proc(sdata)
    
Exit Sub

Err_P:
    Call DataLogger(" [Lane1 UDP DataArrival]  " & Err.Description)

End Sub

Public Sub LPR1_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [Lane1 UDP Error]  " & Description)
End Sub

Public Sub LPR2_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    LPR2_sock.GetData sdata, , bytesTotal
    Call DataLogger("Lane2 UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    Call UDP_Proc(sdata)
    
Exit Sub

Err_P:
    Call DataLogger(" [Lane2 UDP DataArrival]  " & Err.Description)

End Sub

Public Sub LPR2_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [Lane2 UDP Error]  " & Description)
End Sub

Public Sub LPR3_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    LPR3_sock.GetData sdata, , bytesTotal
    Call DataLogger("Lane3 UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    Call UDP_Proc(sdata)
    
Exit Sub

Err_P:
    Call DataLogger(" [Lane3 UDP DataArrival]  " & Err.Description)

End Sub

Public Sub LPR3_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [Lane3 UDP Error]  " & Description)
End Sub

Public Sub LPR4_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    LPR4_sock.GetData sdata, , bytesTotal
    Call DataLogger("Lane4 UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    Call UDP_Proc(sdata)
    
Exit Sub

Err_P:
    Call DataLogger(" [Lane4 UDP DataArrival]  " & Err.Description)

End Sub

Public Sub LPR4_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [Lane4 UDP Error]  " & Description)
End Sub

Public Sub LPR5_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    LPR1_sock.GetData sdata, , bytesTotal
    Call DataLogger("Lane1 UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    Call UDP_Proc(sdata)
    
Exit Sub

Err_P:
    Call DataLogger(" [Lane5 UDP DataArrival]  " & Err.Description)

End Sub

Public Sub LPR5_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [Lane5 UDP Error]  " & Description)
End Sub

'Reomte_UDP 받기
Public Sub RemoteR_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, GateNo As Integer
    Dim CarNum As String
    
On Error GoTo Err_P
    
    RemoteR_sock.GetData sdata, , bytesTotal
    Call DataLogger("RemoteR_sock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    
    Glo_GateNo = Left(sdata, 1)
    i = Len(sdata)
    CarNum = Mid(sdata, 3, i - 2)
    
    
    '스크린 수에 따라서 분기
    If (Glo_Screen_No = 4) Then
        If (Glo_GateNo < Glo_Screen_No) Then
            Call G4Mini_4INShow(CarNum)
        End If
    
    ElseIf (Glo_Screen_No = 2) Then
        If (Glo_GateNo < Glo_Screen_No) Then
            Call Jung_Show(CarNum)
        End If
    End If

Exit Sub

Err_P:
    Call DataLogger(" [RemoteR_sock UDP DataArrival]  " & Err.Description)

End Sub

Public Sub RemoteR_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [RemoteR_sock UDP Error]  " & Description)
End Sub



Public Sub G4Mini_4INShow(Data As String)
Dim i As Integer
Dim GateNo As Integer
Dim GateName As String
Dim CarNo As String
Dim rs As Recordset
Dim qry As String
Dim Tmp_File As String

With FrmG4Mini
'        GateNo = Left(Data, 1)
'        i = LenH(Data)
'        CarNo = Mid(Data, 3, (i - 2))
        GateNo = Glo_GateNo
        CarNo = Data

        qry = "Select * From tb_inout Where PASS_GATE = '" & Glo_GateNo & "' And CAR_NO = '" & CarNo & "' Order By PASS_DATE Desc Limit 1"
        Set rs = New ADODB.Recordset
        rs.Open qry, adoConn

        If Not (rs.EOF) Then
                .lbl_carno(GateNo).Caption = "" & rs!CAR_NO
                Tmp_File = Dir(rs!PASS_IMAGE)
                If (Tmp_File <> "") Then
                    .ImageIn(GateNo).Picture = LoadPicture(rs!PASS_IMAGE)
                End If
                For i = 0 To 3
                    .Shp_Rec(i).Visible = False
                Next i
                .Shp_Rec(GateNo).Visible = True
                .lbl_time_now(GateNo).Caption = "" & rs!PASS_DATE
                .lbl_RecState(GateNo).Caption = "" & rs!PASS_RESULT
                If rs!PASS_YN = "Y" Then
                    .lbl_RecState(GateNo).ForeColor = vbBlue
                Else
                    .lbl_RecState(GateNo).ForeColor = vbRed
                End If
                .List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "   " & " GateNo : " & GateNo & ", 차량번호 : " & rs!CAR_NO & ", 처리결과 : " & rs!PASS_RESULT, 0
            
            
                Set itmX = .ListView2.ListItems.Add(, , "" & rs!PASS_DATE)
                itmX.SubItems(1) = "" & rs!CAR_NO
                itmX.SubItems(2) = "" & rs!CAR_GUBUN
                itmX.SubItems(3) = "" & rs!DRIVER_NAME
                itmX.SubItems(4) = "" & rs!DRIVER_PHONE
                itmX.SubItems(5) = "" & rs!Start_Date
                itmX.SubItems(6) = "" & rs!End_Date
                itmX.SubItems(7) = "" & rs!PASS_RESULT
                'itmX.SubItems(7) = "" & rs!PASS_DATE
                itmX.SubItems(8) = "" & rs!PASS_IMAGE
                '.ListView2.Sorted = False
                '.ListView2.ListItems(.ListView2.ListItems.Count).Selected = True
                '.ListView2.ListItems(.ListView2.ListItems.Count).EnsureVisible
        
        Else
            'Beep
        End If
        Set rs = Nothing

End With

On Error Resume Next

End Sub




